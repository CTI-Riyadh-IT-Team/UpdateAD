#requires -Version 5.1
#requires -Modules ActiveDirectory

$ErrorActionPreference = "Stop"

# ============================================================
# 2 - DELETE ALL STUDENT USER ACCOUNTS
#
# Deletes USER objects only, and only under this exact OU:
# OU=students,OU=Users,OU=Saudis,OU=CTI,DC=cti,DC=org
#
# The OU itself is NOT deleted.
# Groups, computers, contacts, and other object types are NOT deleted.
# SearchScope = Subtree, so user accounts in child OUs are included.
# ============================================================

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$TargetOU  = "OU=students,OU=Users,OU=Saudis,OU=CTI,DC=cti,DC=org"
$BackupPath = Join-Path $ScriptDir ("DELETED-STUDENTS-BACKUP-{0}.csv" -f (Get-Date -Format "yyyyMMdd-HHmmss"))
$LogPath    = Join-Path $ScriptDir ("DELETE-STUDENTS-{0}.csv" -f (Get-Date -Format "yyyyMMdd-HHmmss"))

function Write-Status {
    param(
        [string]$Text,
        [ConsoleColor]$Color = [ConsoleColor]::Gray
    )
    Write-Host $Text -ForegroundColor $Color
}

Write-Host ""
Write-Status "============================================================" Cyan
Write-Status " 2/3 - DELETE EXISTING CTI STUDENT USERS" Cyan
Write-Status "============================================================" Cyan
Write-Host ""

try {
    Import-Module ActiveDirectory -ErrorAction Stop

    $domain = Get-ADDomain -Identity "cti.org" -ErrorAction Stop
    $ou = Get-ADOrganizationalUnit -Identity $TargetOU -ErrorAction Stop

    if ($domain.DNSRoot -ne "cti.org") {
        throw "Safety check failed: connected domain is not cti.org."
    }

    if ($ou.DistinguishedName -ne $TargetOU) {
        throw "Safety check failed: target OU does not exactly match the configured students OU."
    }
}
catch {
    Write-Status "FATAL: $($_.Exception.Message)" Red
    exit 30
}

$students = @(
    Get-ADUser `
        -SearchBase $TargetOU `
        -SearchScope Subtree `
        -Filter * `
        -Properties DisplayName,mail,UserPrincipalName,employeeID,Enabled,DistinguishedName `
        -ErrorAction Stop
)

Write-Status "Domain    : $($domain.DNSRoot)" DarkCyan
Write-Status "Target OU : $TargetOU" DarkCyan
Write-Status "Users     : $($students.Count)" Yellow
Write-Host ""

if ($students.Count -eq 0) {
    Write-Status "No existing student user accounts were found. Nothing to delete." Green
    @() | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8
    exit 0
}

# Export readable metadata before deletion.
# Note: Active Directory passwords cannot be exported/recovered here.
$students |
    Select-Object Name,DisplayName,SamAccountName,UserPrincipalName,employeeID,mail,Enabled,DistinguishedName,ObjectGUID,SID |
    Export-Csv -Path $BackupPath -NoTypeInformation -Encoding UTF8

Write-Status "Metadata backup created: $BackupPath" Yellow
Write-Status "Deleting ONLY user objects under the exact students OU..." Yellow
Write-Host ""

$results = New-Object System.Collections.Generic.List[object]
$deleted = 0
$failed = 0

foreach ($student in $students) {
    try {
        # Safety check each object immediately before removal.
        if (-not $student.DistinguishedName.EndsWith("," + $TargetOU, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "Safety check blocked deletion because DN is outside target OU."
        }

        Remove-ADUser -Identity $student.DistinguishedName -Confirm:$false -ErrorAction Stop

        Write-Status "DELETED  $($student.SamAccountName) | $($student.DisplayName)" Green
        $deleted++

        $results.Add([pscustomobject]@{
            Username = $student.SamAccountName
            DisplayName = $student.DisplayName
            DistinguishedName = $student.DistinguishedName
            Status = "DELETED"
            Message = "Deleted successfully"
        })
    }
    catch {
        $failed++
        $msg = $_.Exception.Message
        Write-Status "FAILED   $($student.SamAccountName) | $msg" Red

        $results.Add([pscustomobject]@{
            Username = $student.SamAccountName
            DisplayName = $student.DisplayName
            DistinguishedName = $student.DistinguishedName
            Status = "FAILED"
            Message = $msg
        })
    }
}

$results | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Status "============================================================" Cyan
Write-Status " DELETE FINISHED" Cyan
Write-Status "============================================================" Cyan
Write-Status "Deleted : $deleted" Green
Write-Status "Failed  : $failed" Red
Write-Status "Backup  : $BackupPath" DarkCyan
Write-Status "Log     : $LogPath" DarkCyan

if ($failed -gt 0) {
    Write-Status "RESULT: FAIL - import will NOT continue." Red
    exit 31
}

# Final check: no user objects may remain in the students OU.
$remaining = @(
    Get-ADUser -SearchBase $TargetOU -SearchScope Subtree -Filter * -ErrorAction Stop
)

if ($remaining.Count -gt 0) {
    Write-Status "RESULT: FAIL - $($remaining.Count) user account(s) still remain in the students OU." Red
    exit 32
}

Write-Status "RESULT: PASS - students OU contains no user accounts." Green
exit 0
