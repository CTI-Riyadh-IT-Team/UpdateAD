#requires -Version 5.1
#requires -Modules ActiveDirectory

$ErrorActionPreference = "Stop"

# ============================================================
# CTI Active Directory bulk users
#
# users.xlsx must be in the SAME folder as this script.
#
# Excel mapping:
# A = Email
# B = English Name          -> Name / CN
# C = Arabic Name           -> DisplayName
# D = Password
# E = Training Number       -> sAMAccountName
#                             UserPrincipalName = E@cti.org
#
# Target OU:
# OU=students,OU=Users,OU=Saudis,OU=CTI,DC=cti,DC=org
# ============================================================

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ExcelPath = Join-Path $ScriptDir "users.xlsx"
$TargetOU  = "OU=students,OU=Users,OU=Saudis,OU=CTI,DC=cti,DC=org"
$UPNSuffix = "cti.org"

function Write-Status {
    param(
        [string]$Text,
        [ConsoleColor]$Color = [ConsoleColor]::Gray
    )
    Write-Host $Text -ForegroundColor $Color
}

function Get-ColumnNumber {
    param([string]$CellReference)

    $letters = ($CellReference -replace '[^A-Za-z]', '').ToUpperInvariant()
    $number = 0

    foreach ($ch in $letters.ToCharArray()) {
        $number = ($number * 26) + ([int][char]$ch - [int][char]'A' + 1)
    }

    return $number
}

function Get-ZipEntryXml {
    param(
        [System.IO.Compression.ZipArchive]$Zip,
        [string]$EntryName
    )

    $entry = $Zip.GetEntry($EntryName)
    if (-not $entry) { return $null }

    $stream = $entry.Open()
    try {
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
        try {
            $text = $reader.ReadToEnd()
        }
        finally {
            $reader.Dispose()
        }

        $xml = New-Object System.Xml.XmlDocument
        $xml.PreserveWhitespace = $false
        $xml.LoadXml($text)
        return $xml
    }
    finally {
        $stream.Dispose()
    }
}

function Get-ExcelRows {
    param([string]$Path)

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)

    try {
        $workbook = Get-ZipEntryXml -Zip $zip -EntryName "xl/workbook.xml"
        if (-not $workbook) { throw "Invalid XLSX: xl/workbook.xml was not found." }

        $wbNs = New-Object System.Xml.XmlNamespaceManager($workbook.NameTable)
        $wbNs.AddNamespace("m", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
        $wbNs.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")

        $firstSheet = $workbook.SelectSingleNode("//m:sheets/m:sheet[1]", $wbNs)
        if (-not $firstSheet) { throw "No worksheet was found in users.xlsx." }

        $relId = $firstSheet.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
        if ([string]::IsNullOrWhiteSpace($relId)) { throw "Could not resolve the first worksheet relationship." }

        $rels = Get-ZipEntryXml -Zip $zip -EntryName "xl/_rels/workbook.xml.rels"
        if (-not $rels) { throw "Invalid XLSX: workbook relationships were not found." }

        $relNs = New-Object System.Xml.XmlNamespaceManager($rels.NameTable)
        $relNs.AddNamespace("p", "http://schemas.openxmlformats.org/package/2006/relationships")

        $relNode = $rels.SelectSingleNode("//p:Relationship[@Id='$relId']", $relNs)
        if (-not $relNode) { throw "Could not locate the first worksheet XML." }

        $sheetTarget = [string]$relNode.Target
        if ($sheetTarget.StartsWith("/")) {
            $sheetEntryName = $sheetTarget.TrimStart("/")
        }
        elseif ($sheetTarget.StartsWith("xl/")) {
            $sheetEntryName = $sheetTarget
        }
        else {
            $sheetEntryName = "xl/" + $sheetTarget.TrimStart("/")
        }

        $parts = New-Object System.Collections.Generic.List[string]
        foreach ($part in $sheetEntryName.Split("/")) {
            if ($part -eq "..") {
                if ($parts.Count -gt 0) { $parts.RemoveAt($parts.Count - 1) }
            }
            elseif ($part -ne "." -and $part -ne "") {
                $parts.Add($part)
            }
        }
        $sheetEntryName = ($parts -join "/")

        $sharedStrings = @()
        $sharedXml = Get-ZipEntryXml -Zip $zip -EntryName "xl/sharedStrings.xml"

        if ($sharedXml) {
            $ssNs = New-Object System.Xml.XmlNamespaceManager($sharedXml.NameTable)
            $ssNs.AddNamespace("m", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")

            foreach ($item in $sharedXml.SelectNodes("//m:sst/m:si", $ssNs)) {
                $value = ""
                foreach ($textNode in $item.SelectNodes(".//m:t", $ssNs)) {
                    $value += $textNode.InnerText
                }
                $sharedStrings += $value
            }
        }

        $sheet = Get-ZipEntryXml -Zip $zip -EntryName $sheetEntryName
        if (-not $sheet) { throw "Could not open worksheet entry: $sheetEntryName" }

        $sheetNs = New-Object System.Xml.XmlNamespaceManager($sheet.NameTable)
        $sheetNs.AddNamespace("m", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")

        foreach ($rowNode in $sheet.SelectNodes("//m:sheetData/m:row", $sheetNs)) {
            $values = @("", "", "", "", "")

            foreach ($cell in $rowNode.SelectNodes("./m:c", $sheetNs)) {
                $col = Get-ColumnNumber -CellReference ([string]$cell.r)
                if ($col -lt 1 -or $col -gt 5) { continue }

                $type = [string]$cell.t
                $value = ""

                if ($type -eq "inlineStr") {
                    foreach ($textNode in $cell.SelectNodes("./m:is//m:t", $sheetNs)) {
                        $value += $textNode.InnerText
                    }
                }
                else {
                    $vNode = $cell.SelectSingleNode("./m:v", $sheetNs)
                    if ($vNode) {
                        $raw = [string]$vNode.InnerText
                        switch ($type) {
                            "s" {
                                $idx = 0
                                if ([int]::TryParse($raw, [ref]$idx) -and $idx -ge 0 -and $idx -lt $sharedStrings.Count) {
                                    $value = [string]$sharedStrings[$idx]
                                }
                                else {
                                    $value = $raw
                                }
                            }
                            "b" {
                                if ($raw -eq "1") { $value = "TRUE" } else { $value = "FALSE" }
                            }
                            default {
                                $value = $raw
                            }
                        }
                    }
                }

                $values[$col - 1] = $value
            }

            [pscustomobject]@{
                ExcelRow = [int]$rowNode.r
                A = [string]$values[0]
                B = [string]$values[1]
                C = [string]$values[2]
                D = [string]$values[3]
                E = [string]$values[4]
            }
        }
    }
    finally {
        $zip.Dispose()
    }
}

function Test-HeaderRow {
    param($Row)

    $combined = "$($Row.A)|$($Row.B)|$($Row.C)|$($Row.D)|$($Row.E)"
    return (
        ([string]$Row.A) -match '(?i)email|e-mail|البريد' -or
        ([string]$Row.D) -match '(?i)password|pass|كلمة' -or
        ([string]$Row.E) -match '(?i)username|user.?logon|login|اسم.?المستخدم|training' -or
        $combined -match '(?i)english.?name|arabic.?name|display.?name'
    )
}

function Get-PreparedRows {
    param([object[]]$Rows)

    $prepared = New-Object System.Collections.Generic.List[object]
    $firstNonEmpty = $true

    foreach ($row in $Rows) {
        $email       = ([string]$row.A).Trim()
        $englishName = ([string]$row.B).Trim()
        $displayName = ([string]$row.C).Trim()
        $password    = [string]$row.D
        $username    = ([string]$row.E).Trim()

        $isEmpty = (
            [string]::IsNullOrWhiteSpace($email) -and
            [string]::IsNullOrWhiteSpace($englishName) -and
            [string]::IsNullOrWhiteSpace($displayName) -and
            [string]::IsNullOrWhiteSpace($password) -and
            [string]::IsNullOrWhiteSpace($username)
        )
        if ($isEmpty) { continue }

        if ($firstNonEmpty) {
            $firstNonEmpty = $false
            if (Test-HeaderRow $row) { continue }
        }

        $prepared.Add([pscustomobject]@{
            ExcelRow    = $row.ExcelRow
            Email       = $email
            EnglishName = $englishName
            DisplayName = $displayName
            Password    = $password
            Username    = $username
            UPN         = if ($username) { "$username@$UPNSuffix" } else { "" }
        })
    }

    return $prepared
}

function Test-PasswordApprox {
    param(
        [string]$Password,
        [string]$Username,
        [string]$EnglishName,
        $Policy
    )

    $problems = New-Object System.Collections.Generic.List[string]

    if ([string]::IsNullOrWhiteSpace($Password)) {
        $problems.Add("Password is empty")
        return $problems
    }

    if ($Policy -and $Password.Length -lt [int]$Policy.MinPasswordLength) {
        $problems.Add("Password is shorter than domain minimum length ($($Policy.MinPasswordLength))")
    }

    if ($Policy -and $Policy.ComplexityEnabled) {
        $classes = 0
        if ($Password -cmatch '[A-Z]') { $classes++ }
        if ($Password -cmatch '[a-z]') { $classes++ }
        if ($Password -match '\d') { $classes++ }
        if ($Password -match '[^A-Za-z0-9]') { $classes++ }

        # Windows complexity also has additional rules. This is a safe pre-check.
        if ($classes -lt 3) {
            $problems.Add("Password may fail domain complexity requirements")
        }

        if ($Username.Length -ge 3 -and $Password.ToLowerInvariant().Contains($Username.ToLowerInvariant())) {
            $problems.Add("Password contains the username")
        }

        foreach ($part in ($EnglishName -split '[\s,._-]+' | Where-Object { $_.Length -ge 3 })) {
            if ($Password.ToLowerInvariant().Contains($part.ToLowerInvariant())) {
                $problems.Add("Password contains part of the user's name: $part")
                break
            }
        }
    }

    return $problems
}

function Initialize-Environment {
    if (-not (Test-Path -LiteralPath $ExcelPath)) {
        throw "users.xlsx was not found beside the script: $ExcelPath"
    }

    Import-Module ActiveDirectory -ErrorAction Stop
    $domain = Get-ADDomain -Identity "cti.org" -ErrorAction Stop
    $null = Get-ADOrganizationalUnit -Identity $TargetOU -ErrorAction Stop
    $policy = Get-ADDefaultDomainPasswordPolicy -Identity "cti.org" -ErrorAction Stop

    return [pscustomobject]@{
        Domain = $domain
        Policy = $policy
    }
}

function Test-UserRow {
    param(
        $User,
        [object[]]$AllPreparedRows,
        $Policy
    )

    $issues = New-Object System.Collections.Generic.List[string]

    if ([string]::IsNullOrWhiteSpace($User.EnglishName)) { $issues.Add("B: English Name is empty") }
    if ([string]::IsNullOrWhiteSpace($User.DisplayName)) { $issues.Add("C: Display Name is empty") }
    if ([string]::IsNullOrWhiteSpace($User.Password))    { $issues.Add("D: Password is empty") }
    if ([string]::IsNullOrWhiteSpace($User.Username))    { $issues.Add("E: Username/Training Number is empty") }

    if (-not [string]::IsNullOrWhiteSpace($User.Email) -and $User.Email -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
        $issues.Add("A: Email format is invalid")
    }

    if (-not [string]::IsNullOrWhiteSpace($User.Username)) {
        if ($User.Username.Length -gt 20) {
            $issues.Add("E: sAMAccountName is longer than 20 characters")
        }

        if ($User.Username -match '[\"/\\\[\]:;|=,\+\*\?<>\s]') {
            $issues.Add("E: Username contains a character not suitable for sAMAccountName")
        }

        $dupUserRows = @($AllPreparedRows | Where-Object { $_.Username -eq $User.Username })
        if ($dupUserRows.Count -gt 1) {
            $issues.Add("Duplicate username inside users.xlsx")
        }

        $safeUsername = $User.Username.Replace("'", "''")
        if (Get-ADUser -Filter "SamAccountName -eq '$safeUsername'" -ErrorAction Stop) {
            $issues.Add("Username already exists in Active Directory")
        }

        $safeUPN = $User.UPN.Replace("'", "''")
        if (Get-ADUser -Filter "UserPrincipalName -eq '$safeUPN'" -ErrorAction Stop) {
            $issues.Add("UPN already exists in Active Directory")
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($User.EnglishName)) {
        $dupNameRows = @($AllPreparedRows | Where-Object { $_.EnglishName -eq $User.EnglishName })
        if ($dupNameRows.Count -gt 1) {
            $issues.Add("Duplicate English Name/CN inside users.xlsx")
        }

        $escapedName = $User.EnglishName.Replace("'", "''")
        if (Get-ADUser -SearchBase $TargetOU -SearchScope OneLevel -Filter "Name -eq '$escapedName'" -ErrorAction Stop) {
            $issues.Add("A user with the same Name/CN already exists in the target OU")
        }
    }

    foreach ($p in (Test-PasswordApprox -Password $User.Password -Username $User.Username -EnglishName $User.EnglishName -Policy $Policy)) {
        $issues.Add($p)
    }

    return $issues
}

# ============================================================
# VERIFY ONLY - NO AD CHANGES
# ============================================================

$ReportPath = Join-Path $ScriptDir ("VERIFY-AD-Users-{0}.csv" -f (Get-Date -Format "yyyyMMdd-HHmmss"))

Write-Host ""
Write-Status "============================================================" Cyan
Write-Status " CTI USERS - VERIFY ONLY (NO CHANGES)" Cyan
Write-Status "============================================================" Cyan
Write-Host ""

try {
    $envInfo = Initialize-Environment
    $rawRows = @(Get-ExcelRows -Path $ExcelPath)
    $users = @(Get-PreparedRows -Rows $rawRows)
}
catch {
    Write-Status "FATAL: $($_.Exception.Message)" Red
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Status "Domain    : $($envInfo.Domain.DNSRoot)" DarkCyan
Write-Status "Target OU : $TargetOU" DarkCyan
Write-Status "Excel     : $ExcelPath" DarkCyan
Write-Status "Mode      : VERIFY ONLY - nothing will be created or changed" Yellow
Write-Host ""

$results = New-Object System.Collections.Generic.List[object]
$valid = 0
$invalid = 0

foreach ($user in $users) {
    $issues = @(Test-UserRow -User $user -AllPreparedRows $users -Policy $envInfo.Policy)

    if ($issues.Count -eq 0) {
        Write-Status "[$($user.ExcelRow)] OK  $($user.Username) | $($user.DisplayName)" Green
        $valid++
        $status = "READY"
        $message = "Ready to import"
    }
    else {
        Write-Status "[$($user.ExcelRow)] FAIL  $($user.Username) | $($user.DisplayName)" Red
        foreach ($issue in $issues) {
            Write-Status "    - $issue" Yellow
        }
        $invalid++
        $status = "FAILED"
        $message = ($issues -join " | ")
    }

    $results.Add([pscustomobject]@{
        Row         = $user.ExcelRow
        Email       = $user.Email
        EnglishName = $user.EnglishName
        DisplayName = $user.DisplayName
        Username    = $user.Username
        UPN         = $user.UPN
        Status      = $status
        Message     = $message
    })
}

$results | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Status "============================================================" Cyan
Write-Status " VERIFY FINISHED - NO CHANGES WERE MADE" Cyan
Write-Status "============================================================" Cyan
Write-Status "Ready  : $valid" Green
Write-Status "Failed : $invalid" Red
Write-Status "Report : $ReportPath" DarkCyan
Write-Host ""

if ($invalid -eq 0 -and $valid -gt 0) {
    Write-Status "RESULT: PASS - file is ready for the import script." Green
}
elseif ($valid -eq 0) {
    Write-Status "RESULT: FAIL - no valid users found." Red
}
else {
    Write-Status "RESULT: FAIL - fix the failed rows before importing." Red
}

Write-Host ""
Read-Host "Press Enter to exit"
