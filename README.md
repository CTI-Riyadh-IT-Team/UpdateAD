# دليل إدارة حسابات طلاب CTI في Active Directory

**الإصدار:** 2.1  
**تاريخ التحديث:** 17 سبتمبر 2026  
**النطاق:** حسابات الطلاب داخل OU `students` في Domain `cti.org`

> هذه الحزمة مخصصة للموظفين المسؤولين عن استبدال حسابات الطلاب في Active Directory من ملف `users.xlsx`. لا يحتاج الموظف إلى تعديل أي سطر برمجي. في الاستخدام المعتاد: جهّز ملف Excel، شغّل التحقق، ثم شغّل السكربت الرئيسي.

---

## 1. ملخص سريع

المسار المستهدف في Active Directory:

```text
cti.org
└── CTI
    └── Saudis
        └── Users
            └── students
```

Distinguished Name:

```text
OU=students,OU=Users,OU=Saudis,OU=CTI,DC=cti,DC=org
```

العملية الكاملة:

```text
1) VERIFY
   فحص users.xlsx وبيئة AD بدون أي تعديل

2) DELETE
   حذف جميع User Objects الموجودة تحت OU الطلاب

3) IMPORT
   إنشاء حسابات الطلاب من users.xlsx من جديد
```

إذا فشل **VERIFY** فلن يبدأ الحذف.  
إذا فشل **DELETE** فلن يبدأ الاستيراد.  
قبل الحذف يتم إنشاء CSV يحتوي بيانات تعريفية للحسابات الموجودة، لكنه **لا يحتوي كلمات المرور**.

---

## 2. ملفات الحزمة

يجب أن يكون المجلد بالشكل التالي:

```text
C:\ADImport\
├── users.xlsx
├── 0-Run-All-CTI-Student-Replacement.ps1
├── 1-Verify-CTI-Users.ps1
├── 2-Delete-CTI-Students.ps1
├── 3-Import-CTI-Users.ps1
├── README.md
└── CTI-AD-Student-Management-Guide.pdf
```

وظيفة الملفات:

| الملف | الوظيفة |
|---|---|
| `0-Run-All-CTI-Student-Replacement.ps1` | يشغّل المراحل الثلاث بالترتيب ويوقف عند أي خطأ. |
| `1-Verify-CTI-Users.ps1` | يتحقق من ملف Excel والبيانات والـAD فقط. لا يحذف ولا ينشئ حسابات. |
| `2-Delete-CTI-Students.ps1` | يحذف User Objects الموجودة تحت OU الطلاب فقط. |
| `3-Import-CTI-Users.ps1` | ينشئ الحسابات الجديدة من `users.xlsx`. |
| `users.xlsx` | ملف مصدر الطلاب، ويجب أن يكون بهذا الاسم بالضبط. |

---

## 3. تجهيز users.xlsx

تقرأ الحزمة أول Worksheet في الملف.

ترتيب الأعمدة ثابت:

| العمود | البيانات | Active Directory |
|---|---|---|
| **A** | البريد الإلكتروني | `mail` |
| **B** | الاسم بالإنجليزي | `description` |
| **C** | الاسم بالعربي | `displayName` |
| **D** | كلمة المرور | Password |
| **E** | الرقم التدريبي / Username | `Name/CN` + `sAMAccountName` + `userPrincipalName` + `employeeID` |

### مثال

```text
A = 446149588@tvtc.edu.sa
B = Ahmed Mohammed Ali
C = أحمد محمد علي
D = Example@1234
E = 446149588
```

الحساب الناتج:

```text
Name / CN          = 446149588
displayName        = أحمد محمد علي
description        = Ahmed Mohammed Ali
mail               = 446149588@tvtc.edu.sa
sAMAccountName     = 446149588
userPrincipalName  = 446149588@cti.org
employeeID         = 446149588
```

### لماذا CN هو الرقم التدريبي؟

أسماء الطلاب قد تتكرر، بينما الرقم التدريبي يجب أن يكون فريداً. لذلك تستخدم الحزمة الرقم التدريبي كـ `Name/CN` حتى لا يحدث تعارض عند وجود طالبين بنفس الاسم.

**تكرار الاسم الإنجليزي مسموح.**

### لماذا employeeID مهم؟

في البيئة الحالية:

- الرقم التدريبي المستخدم كاسم دخول موجود في `sAMAccountName`.
- `userPrincipalName` يصبح بالشكل `رقم@cti.org`.
- منصة **سم** تحتاج قيمة الرقم التدريبي في خاصية `employeeID` لكي يظهر/يُتعرف عليه كرقم تدريبي في بيانات الطالب.

إذا كان الحساب يعمل ولكن سم يعرض أن الرقم التدريبي غير مسجل في الدليل، أول خاصية يجب فحصها هي:

```text
employeeID
```

ويجب أن تساوي الرقم التدريبي في العمود E.

---

## 4. الحقول المطلوبة

كل الأعمدة من **A إلى E مطلوبة**.

يفشل التحقق إذا كان أحد الحقول التالية فارغاً:

- A: البريد الإلكتروني.
- B: الاسم بالإنجليزي.
- C: الاسم بالعربي.
- D: كلمة المرور.
- E: الرقم التدريبي / Username.

كما يتحقق السكربت من:

- صيغة البريد الإلكتروني.
- عدم تكرار البريد داخل `users.xlsx`.
- عدم تكرار الرقم التدريبي/Username داخل `users.xlsx`.
- صلاحية القيمة للاستخدام كـ `sAMAccountName`.
- فحص مبدئي لكلمة المرور مقابل Default Domain Password Policy.
- عدم وجود Username أو UPN متعارض **خارج** OU الطلاب.

### ملاحظة مهمة عن الحسابات الموجودة مسبقاً

وجود نفس الطالب مسبقاً **داخل OU الطلاب ليس خطأ**؛ لأن الخطوة الثانية ستحذف الحسابات القديمة قبل إنشاء الحسابات الجديدة.

أما إذا كان نفس `sAMAccountName` أو UPN موجوداً خارج OU الطلاب، فيعتبر ذلك خطأ ويوقف العملية، لأن سكربت الحذف لن يلمس الحساب الموجود خارج النطاق.

---

## 5. تنسيق Excel

ينصح بتنسيق الأعمدة التالية كـ **Text** قبل إدخال البيانات:

```text
A - Email
D - Password
E - Training Number
```

هذا مهم خصوصاً للرقم التدريبي وكلمة المرور، لأن Excel قد يحذف صفراً في البداية أو يعيد تفسير بعض القيم كأرقام/تواريخ/صيغ.

يمكن أن يحتوي الصف الأول على Header مثل:

```text
Email | English Name | Arabic Name | Password | Username
```

والسكربت يتجاهله عند اكتشافه.

---

## 6. المتطلبات والصلاحيات

يجب أن يعمل الجهاز على Windows ويستطيع الوصول إلى Domain:

```text
cti.org
```

ويلزم توفر PowerShell Active Directory Module.

يمكن اختبار وجود الوحدة:

```powershell
Get-Module -ListAvailable ActiveDirectory
```

ويجب أن يملك الحساب المستخدم صلاحيات:

- قراءة OU الطلاب.
- إنشاء User Objects.
- تعيين كلمات المرور.
- تفعيل الحسابات.
- حذف User Objects داخل OU الطلاب.

شغّل Windows PowerShell بصلاحية:

```text
Run as administrator
```

---

## 7. طريقة التشغيل الموصى بها

افتح PowerShell كمسؤول ثم:

```powershell
cd C:\ADImport
```

اسمح للسكربتات في الجلسة الحالية فقط:

```powershell
Set-ExecutionPolicy -Scope Process Bypass -Force
```

### أولاً: التحقق فقط

```powershell
.\1-Verify-CTI-Users.ps1
```

يجب أن تنتهي النتيجة بـ:

```text
RESULT: PASS
```

إذا ظهرت `FAIL` لا تشغّل الحذف. أصلح ملف Excel حسب التقرير ثم أعد التحقق.

### ثانياً: تشغيل العملية الكاملة

بعد نجاح التحقق:

```powershell
.\0-Run-All-CTI-Student-Replacement.ps1
```

السكربت الرئيسي يعيد التحقق ثم ينفذ:

```text
VERIFY -> DELETE -> IMPORT
```

---

## 8. ما الذي يفعله VERIFY؟

`1-Verify-CTI-Users.ps1` لا يغيّر أي شيء في Active Directory.

يفحص:

1. وجود `users.xlsx`.
2. إمكانية الوصول إلى `cti.org`.
3. أن OU المستهدف موجود ومطابق للمسار الثابت في السكربت.
4. وجود صفوف طلاب فعلية.
5. اكتمال A-E.
6. صحة البريد.
7. تكرار البريد داخل الملف.
8. تكرار Username/Training Number داخل الملف.
9. طول وصلاحية `sAMAccountName`.
10. Password Policy بشكل مبدئي.
11. تعارض Username أو UPN خارج OU الطلاب.

### أمثلة

سليم:

```text
[125] OK    446149588 | أحمد محمد علي
```

خطأ:

```text
[126] FAIL  446149589 | أحمد خالد
    - D: Password is empty
```

أو:

```text
[127] FAIL  446149590 | محمد علي
    - A: Email is empty
```

أو:

```text
[128] FAIL  446149591 | خالد محمد
    - E: Duplicate username inside users.xlsx
```

---

## 9. ما الذي يفعله DELETE؟

`2-Delete-CTI-Students.ps1` هو المرحلة الحساسة.

يحذف:

```text
User Objects
```

الموجودة تحت:

```text
OU=students,OU=Users,OU=Saudis,OU=CTI,DC=cti,DC=org
```

ويستخدم `SearchScope Subtree`، لذلك يشمل مستخدمين داخل Child OUs تحت `students` إن وجدت.

لا يحذف:

- OU `students`.
- Groups.
- Computers.
- Contacts.
- أي حساب خارج OU الطلاب.

قبل الحذف ينشئ تقريراً مثل:

```text
DELETED-STUDENTS-BACKUP-20260917-170000.csv
```

ويحتوي مثلاً على:

```text
Name
DisplayName
SamAccountName
UserPrincipalName
employeeID
mail
Enabled
DistinguishedName
ObjectGUID
SID
```

**هذا ليس Backup كاملاً لـAD ولا يحتوي كلمات المرور.**

بعد الحذف، يفحص السكربت أن عدد User Objects المتبقية في النطاق أصبح صفراً. إذا بقي حساب، تتوقف العملية ولا يبدأ الاستيراد.

---

## 10. ما الذي يفعله IMPORT؟

`3-Import-CTI-Users.ps1` يعيد التحقق من البيانات ثم ينشئ الحسابات.

الحساب ينشأ أولاً وهو:

```text
Disabled
```

ثم يقوم السكربت بما يلي:

1. إنشاء User Object.
2. كتابة:
   - CN/Name.
   - DisplayName.
   - Description.
   - `sAMAccountName`.
   - UPN.
   - `employeeID`.
   - `mail`.
3. قراءة الحساب مرة أخرى والتأكد أن الحقول الحرجة حُفظت فعلياً.
4. تعيين كلمة المرور.
5. تفعيل الحساب.

إذا فشل أي جزء بعد إنشاء User Object، يحاول السكربت إبقاء الحساب:

```text
Disabled
```

حتى لا يترك حساباً فعالاً بحالة ناقصة.

### فحص employeeID

النسخة الحالية تتحقق بعد الإنشاء من أن:

```text
employeeID = Training Number
```

قبل تفعيل الحساب.

هذا الفحص أضيف لأن `employeeID` مطلوب في التكامل الحالي مع سم لبيانات الرقم التدريبي.

---

## 11. التحقق اليدوي من حساب في الواجهة الرسومية

لفحص طالب واحد من **Active Directory Users and Computers**:

1. افتح ADUC.
2. من الأعلى:
   `View -> Advanced Features`
3. انتقل يدوياً إلى:
   `CTI -> Saudis -> Users -> students`
4. افتح خصائص حساب الطالب.
5. افتح:
   `Attribute Editor`
6. تحقق من:

```text
employeeID
sAMAccountName
userPrincipalName
displayName
mail
description
```

للطالب رقم `446149588` يجب أن يكون تقريباً:

```text
Name/CN           = 446149588
sAMAccountName    = 446149588
userPrincipalName = 446149588@cti.org
employeeID        = 446149588
```

إذا كان `employeeID` فارغاً فقد يعمل الحساب في AD، لكن سم قد يعرض أن الرقم التدريبي غير مسجل في الدليل.

---

## 12. التقارير الناتجة

الحزمة تنشئ CSV تلقائياً في نفس المجلد.

### تقرير التحقق

```text
VERIFY-AD-Users-YYYYMMDD-HHMMSS.csv
```

أهم القيم:

```text
READY
FAILED
```

### تقرير الحذف

```text
DELETE-STUDENTS-YYYYMMDD-HHMMSS.csv
```

### نسخة بيانات قبل الحذف

```text
DELETED-STUDENTS-BACKUP-YYYYMMDD-HHMMSS.csv
```

### تقرير الاستيراد

```text
IMPORT-AD-Users-YYYYMMDD-HHMMSS.csv
```

أهم القيم:

```text
CREATED
FAILED
```

احتفظ بالتقارير مؤقتاً لأغراض المراجعة، ثم تعامل معها كبيانات حساسة لأنها تحتوي معلومات الطلاب.

---

## 13. Exit Codes

| المرحلة | Exit Code | المعنى |
|---|---:|---|
| Verify | 0 | نجاح |
| Verify | 10 | خطأ أساسي في الملف/AD/OU |
| Verify | 11 | لا توجد صفوف طلاب |
| Verify | 20 | بيانات غير صالحة |
| Delete | 0 | نجاح |
| Delete | 30 | فشل التحقق من Domain/OU |
| Delete | 31 | فشل حذف حساب أو أكثر |
| Delete | 32 | بقيت حسابات بعد الحذف |
| Import | 0 | نجاح |
| Import | 40 | خطأ أساسي |
| Import | 41 | لا توجد صفوف طلاب |
| Import | 42 | فشل Pre-check |
| Import | 43 | فشل إنشاء حساب أو أكثر |

السكربت الرئيسي يتوقف عند أول Exit Code غير صفر.

---

## 14. مشاكل شائعة

### `The term ... is not recognized`

تأكد من أنك داخل المجلد الصحيح:

```powershell
cd C:\ADImport
Get-ChildItem -Name *.ps1
```

ثم شغّل الملف باسمه كما يظهر.

### ظهور `>>`

PowerShell ينتظر إكمال أمر أو إغلاق quotation.

اضغط:

```text
Ctrl + C
```

ثم أعد كتابة الأمر.

### العربي يظهر بأحرف غريبة في Console

سكربت التحقق يحاول ضبط Console على UTF-8. إذا بقي العرض غير صحيح، فهذا غالباً مشكلة عرض في Windows PowerShell القديم، وليس دليلاً وحده على فساد قيمة `displayName` داخل AD. تحقق من الحساب في ADUC.

### الحساب يعمل لكن سم يقول الرقم التدريبي غير مسجل

افحص:

```text
employeeID
```

ويجب أن يساوي الرقم التدريبي.

### Password Policy

قد يرفض AD كلمة مرور حتى لو اجتاز الفحص المبدئي، لأن بعض السياسات قد تكون أدق من الفحص المحلي. في هذه الحالة يظهر الخطأ في تقرير الاستيراد ويبقى الحساب غير مفعل قدر الإمكان.

---

## 15. تشغيل المراحل يدوياً

للتشخيص فقط يمكن تشغيل كل مرحلة منفردة:

```powershell
.\1-Verify-CTI-Users.ps1
```

ثم:

```powershell
.\2-Delete-CTI-Students.ps1
```

ثم:

```powershell
.\3-Import-CTI-Users.ps1
```

**لا تشغّل Delete يدوياً إلا بعد نجاح Verify ومعرفة أنك تريد حذف الحسابات الحالية.**

---

## 16. إجراءات الأمان

`users.xlsx` يحتوي كلمات مرور بنص صريح. لذلك:

- لا ترفعه إلى GitHub.
- لا ترسله بالبريد دون حماية.
- لا تحفظه في مجلد مشاركة عام.
- احذفه أو انقله إلى موقع محمي بعد انتهاء العملية حسب سياسة الجهة.

أيضاً لا تنشر:

```text
VERIFY-AD-Users-*.csv
IMPORT-AD-Users-*.csv
DELETE-STUDENTS-*.csv
DELETED-STUDENTS-BACKUP-*.csv
```

لأنها تحتوي معلومات تعريفية للطلاب.

إذا تم وضع المشروع في Git، أضف على الأقل:

```gitignore
users.xlsx
*.xlsx
*.csv
*.log
```

---

## 17. الاسترجاع إذا حدثت مشكلة بعد الحذف

أفضل مصدر لإعادة إنشاء الحسابات هو `users.xlsx` الصحيح.

إذا نجح الحذف ثم فشل الاستيراد:

1. لا تعِد تشغيل DELETE.
2. راجع تقرير `IMPORT-AD-Users-*.csv`.
3. أصلح السبب.
4. شغّل Verify من جديد.
5. شغّل Import فقط:

```powershell
.\3-Import-CTI-Users.ps1
```

ملف `DELETED-STUDENTS-BACKUP-*.csv` مفيد للمراجعة، لكنه لا يعيد كلمات المرور أو جميع خصائص AD.

إذا كانت **Active Directory Recycle Bin** مفعلة، يمكن لمسؤول AD استخدام آليات الاستعادة الرسمية عند الحاجة.

---

## 18. Checklist قبل التشغيل

قبل بدء العملية تأكد من:

- [ ] `users.xlsx` موجود في نفس مجلد السكربتات.
- [ ] الأعمدة A-E مرتبة بشكل صحيح.
- [ ] لا توجد خلايا ناقصة في البيانات المطلوبة.
- [ ] العمود E يحتوي الرقم التدريبي الصحيح.
- [ ] العمود D يحتوي كلمات مرور مطابقة للسياسة.
- [ ] تم إغلاق Excel.
- [ ] أنت متصل بـ `cti.org`.
- [ ] حساب التشغيل لديه صلاحيات كافية.
- [ ] تم تشغيل `1-Verify-CTI-Users.ps1`.
- [ ] النتيجة `RESULT: PASS`.
- [ ] تمت مراجعة أي تحذيرات قبل بدء الحذف.

بعد العملية:

- [ ] النتيجة النهائية SUCCESS.
- [ ] عدد الحسابات المنشأة منطقي.
- [ ] تم اختبار حساب طالب.
- [ ] `employeeID` في الحساب يساوي الرقم التدريبي.
- [ ] تم اختبار تسجيل الدخول للأنظمة المرتبطة عند الحاجة.
- [ ] تم حماية أو إزالة ملف كلمات المرور بعد الانتهاء.

---

## 19. التغيير في الإصدار 2.1

تم اعتماد الرقم التدريبي في العمود **E** كقيمة موحدة للخصائص التالية:

```text
Name / CN
sAMAccountName
userPrincipalName prefix
employeeID
```

وتمت إضافة فحص بعد إنشاء الحساب للتأكد من حفظ `employeeID` قبل تفعيل المستخدم.

سبب التغيير: تم التأكد عملياً أن الحساب يمكنه تسجيل الدخول بينما لا تتعرف منصة سم على الرقم التدريبي إذا كانت خاصية `employeeID` فارغة.

---

**نهاية الدليل**
