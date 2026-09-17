#requires -Version 5.1
#requires -Modules ActiveDirectory

$ErrorActionPreference = "Stop"

# Better UTF-8 output for Arabic names in Windows PowerShell console
try { chcp 65001 > $null } catch {}
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$OutputEncoding = [Console]::OutputEncoding

# ============================================================
# CTI - VERIFY users.xlsx ONLY
#
# Purpose:
# - Validate that the NEW Excel data is complete and usable.
# - Existing users INSIDE the students OU are NOT an error,
#   because step 2 deletes them before step 3 imports fresh users.
# - A collision OUTSIDE the students OU IS an error because it
#   will remain after the students OU is cleared.
#
# Excel:
# A = Email                    REQUIRED
# B = English Name            REQUIRED -> Description
# C = Arabic DisplayName       REQUIRED -> DisplayName
# D = Password                 REQUIRED
# E = Training Number/Username REQUIRED -> Name/CN + sAMAccountName + UPN
#
# Target:
# OU=students,OU=Users,OU=Saudis,OU=CTI,DC=cti,DC=org
# ============================================================

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ExcelPath = Join-Path $ScriptDir "users.xlsx"
$TargetOU  = "OU=students,OU=Users,OU=Saudis,OU=CTI,DC=cti,DC=org"
$UPNSuffix = "cti.org"
$ReportPath = Join-Path $ScriptDir ("VERIFY-AD-Users-{0}.csv" -f (Get-Date -Format "yyyyMMdd-HHmmss"))

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
        try { $text = $reader.ReadToEnd() }
        finally { $reader.Dispose() }

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
                                else { $value = $raw }
                            }
                            "b" {
                                if ($raw -eq "1") { $value = "TRUE" } else { $value = "FALSE" }
                            }
                            default { $value = $raw }
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

function Test-IsInsideTargetOU {
    param([string]$DistinguishedName)

    if ([string]::IsNullOrWhiteSpace($DistinguishedName)) { return $false }

    return $DistinguishedName.EndsWith(
        "," + $TargetOU,
        [System.StringComparison]::OrdinalIgnoreCase
    )
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
        $problems.Add("D: Password is empty")
        return $problems
    }

    if ($Password.Length -lt [int]$Policy.MinPasswordLength) {
        $problems.Add("D: Password is shorter than the domain minimum ($($Policy.MinPasswordLength))")
    }

    if ($Policy.ComplexityEnabled) {
        $classes = 0
        if ($Password -cmatch '[A-Z]') { $classes++ }
        if ($Password -cmatch '[a-z]') { $classes++ }
        if ($Password -match '\d') { $classes++ }
        if ($Password -match '[^A-Za-z0-9]') { $classes++ }

        if ($classes -lt 3) {
            $problems.Add("D: Password may fail the domain complexity policy")
        }

        if ($Username.Length -ge 3 -and $Password.ToLowerInvariant().Contains($Username.ToLowerInvariant())) {
            $problems.Add("D: Password contains the username")
        }

        foreach ($part in ($EnglishName -split '[\s,._-]+' | Where-Object { $_.Length -ge 3 })) {
            if ($Password.ToLowerInvariant().Contains($part.ToLowerInvariant())) {
                $problems.Add("D: Password contains part of the English name: $part")
                break
            }
        }
    }

    return $problems
}

Write-Host ""
Write-Status "============================================================" Cyan
Write-Status " CTI USERS - VERIFY NEW EXCEL DATA ONLY" Cyan
Write-Status "============================================================" Cyan
Write-Host ""

try {
    if (-not (Test-Path -LiteralPath $ExcelPath)) {
        throw "users.xlsx was not found beside the script: $ExcelPath"
    }

    Import-Module ActiveDirectory -ErrorAction Stop

    $domain = Get-ADDomain -Identity "cti.org" -ErrorAction Stop
    $ou = Get-ADOrganizationalUnit -Identity $TargetOU -ErrorAction Stop
    $policy = Get-ADDefaultDomainPasswordPolicy -Identity "cti.org" -ErrorAction Stop

    if ($domain.DNSRoot -ne "cti.org") {
        throw "Connected domain is not cti.org."
    }

    if ($ou.DistinguishedName -ne $TargetOU) {
        throw "The resolved OU is not the configured students OU."
    }

    $rawRows = @(Get-ExcelRows -Path $ExcelPath)
    $users = @(Get-PreparedRows -Rows $rawRows)
}
catch {
    Write-Status "FATAL: $($_.Exception.Message)" Red
    exit 10
}

if ($users.Count -eq 0) {
    Write-Status "FAIL: users.xlsx contains no student rows." Red
    exit 11
}

Write-Status "Domain    : $($domain.DNSRoot)" DarkCyan
Write-Status "Target OU : $TargetOU" DarkCyan
Write-Status "Excel     : $ExcelPath" DarkCyan
Write-Status "Students  : $($users.Count)" DarkCyan
Write-Host ""

$results = New-Object System.Collections.Generic.List[object]
$ready = 0
$failed = 0
$existingInsideOU = 0

foreach ($user in $users) {
    $issues = New-Object System.Collections.Generic.List[string]
    $notes  = New-Object System.Collections.Generic.List[string]

    # A-E are all required
    if ([string]::IsNullOrWhiteSpace($user.Email))       { $issues.Add("A: Email is empty") }
    if ([string]::IsNullOrWhiteSpace($user.EnglishName)) { $issues.Add("B: English Name is empty") }
    if ([string]::IsNullOrWhiteSpace($user.DisplayName)) { $issues.Add("C: Arabic DisplayName is empty") }
    if ([string]::IsNullOrWhiteSpace($user.Password))    { $issues.Add("D: Password is empty") }
    if ([string]::IsNullOrWhiteSpace($user.Username))    { $issues.Add("E: Username/Training Number is empty") }

    # Email format
    if (-not [string]::IsNullOrWhiteSpace($user.Email) -and
        $user.Email -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
        $issues.Add("A: Email format is invalid")
    }

    # Username validation
    if (-not [string]::IsNullOrWhiteSpace($user.Username)) {
        if ($user.Username.Length -gt 20) {
            $issues.Add("E: Username is longer than 20 characters")
        }

        if ($user.Username -match '[\"/\\\[\]:;|=,\+\*\?<>\s]') {
            $issues.Add("E: Username contains invalid characters")
        }

        # Duplicate username inside the NEW Excel file
        if (@($users | Where-Object { $_.Username -eq $user.Username }).Count -gt 1) {
            $issues.Add("E: Duplicate username inside users.xlsx")
        }

        # Existing account check:
        # - inside students OU = OK; step 2 will delete it
        # - outside students OU = FAIL
        $safeUsername = $user.Username.Replace("'", "''")
        $existingBySam = @(
            Get-ADUser `
                -Filter "SamAccountName -eq '$safeUsername'" `
                -Properties DistinguishedName `
                -ErrorAction Stop
        )

        foreach ($existing in $existingBySam) {
            if (Test-IsInsideTargetOU -DistinguishedName $existing.DistinguishedName) {
                $notes.Add("Existing account in students OU - OK, it will be deleted in step 2")
                $existingInsideOU++
            }
            else {
                $issues.Add("Username already exists OUTSIDE students OU: $($existing.DistinguishedName)")
            }
        }

        $safeUPN = $user.UPN.Replace("'", "''")
        $existingByUPN = @(
            Get-ADUser `
                -Filter "UserPrincipalName -eq '$safeUPN'" `
                -Properties DistinguishedName `
                -ErrorAction Stop
        )

        foreach ($existing in $existingByUPN) {
            if (-not (Test-IsInsideTargetOU -DistinguishedName $existing.DistinguishedName)) {
                $issues.Add("UPN already exists OUTSIDE students OU: $($existing.DistinguishedName)")
            }
        }
    }

    # Duplicate English names are allowed.
    # CN/Name will use the unique training number from column E.
    if (-not [string]::IsNullOrWhiteSpace($user.EnglishName)) {
        if (@($users | Where-Object { $_.EnglishName -eq $user.EnglishName }).Count -gt 1) {
            $notes.Add("Duplicate English name is allowed; CN will use training number")
        }
    }

    # Password pre-check
    foreach ($p in (Test-PasswordApprox `
        -Password $user.Password `
        -Username $user.Username `
        -EnglishName $user.EnglishName `
        -Policy $policy)) {
        $issues.Add($p)
    }

    if ($issues.Count -eq 0) {
        Write-Status "[$($user.ExcelRow)] OK    $($user.Username) | $($user.DisplayName)" Green
        $ready++
        $status = "READY"
        $message = if ($notes.Count -gt 0) { $notes -join " | " } else { "Ready" }
    }
    else {
        Write-Status "[$($user.ExcelRow)] FAIL  $($user.Username) | $($user.DisplayName)" Red
        foreach ($issue in $issues) {
            Write-Status "    - $issue" Yellow
        }
        $failed++
        $status = "FAILED"
        $message = $issues -join " | "
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
Write-Status " VERIFY FINISHED - NO AD CHANGES WERE MADE" Cyan
Write-Status "============================================================" Cyan
Write-Status "Ready  : $ready" Green
Write-Status "Failed : $failed" Red
Write-Status "Report : $ReportPath" DarkCyan
Write-Host ""

if ($failed -gt 0) {
    Write-Status "RESULT: FAIL - fix the Excel data before delete/import." Red
    exit 20
}

Write-Status "RESULT: PASS - Excel data is ready for delete -> import." Green
exit 0
