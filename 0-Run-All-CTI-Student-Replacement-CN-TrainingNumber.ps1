#requires -Version 5.1

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$VerifyScript = Join-Path $ScriptDir "1-Verify-CTI-Users-CN-TrainingNumber.ps1"
$DeleteScript = Join-Path $ScriptDir "2-Delete-CTI-Students.ps1"
$ImportScript = Join-Path $ScriptDir "3-Import-CTI-Users-CN-TrainingNumber.ps1"
$ExcelPath    = Join-Path $ScriptDir "users.xlsx"

function Write-Status {
    param(
        [string]$Text,
        [ConsoleColor]$Color = [ConsoleColor]::Gray
    )
    Write-Host $Text -ForegroundColor $Color
}

Write-Host ""
Write-Status "================================================================" Cyan
Write-Status " CTI STUDENT ACTIVE DIRECTORY - FULL REPLACEMENT WORKFLOW" Cyan
Write-Status "================================================================" Cyan
Write-Host ""
Write-Status "Sequence:" White
Write-Status "  1) Verify users.xlsx and AD environment - NO CHANGES" White
Write-Status "  2) Delete ALL user accounts under the students OU" Yellow
Write-Status "  3) Import users.xlsx as fresh AD accounts" White
Write-Host ""
Write-Status "Target OU:" DarkCyan
Write-Status "OU=students,OU=Users,OU=Saudis,OU=CTI,DC=cti,DC=org" DarkCyan
Write-Host ""

foreach ($required in @($VerifyScript, $DeleteScript, $ImportScript, $ExcelPath)) {
    if (-not (Test-Path -LiteralPath $required)) {
        Write-Status "FATAL: required file is missing: $required" Red
        Read-Host "Press Enter to exit"
        exit 100
    }
}

# STEP 1
Write-Host ""
Write-Status "----------------------------------------------------------------" DarkGray
Write-Status "STEP 1: VERIFY" Cyan
Write-Status "----------------------------------------------------------------" DarkGray

& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $VerifyScript
$verifyCode = $LASTEXITCODE

if ($verifyCode -ne 0) {
    Write-Host ""
    Write-Status "WORKFLOW STOPPED: verification failed (exit code $verifyCode)." Red
    Write-Status "Nothing was deleted and nothing was imported." Yellow
    Read-Host "Press Enter to exit"
    exit $verifyCode
}

# STEP 2
Write-Host ""
Write-Status "Verification PASSED." Green
Write-Status "WARNING: the next step permanently deletes existing student user objects." Yellow
Write-Status "A CSV metadata backup will be created before deletion; passwords cannot be backed up." Yellow
Write-Host ""
Write-Status "Starting deletion automatically in 5 seconds. Press Ctrl+C now to cancel." Yellow
Start-Sleep -Seconds 5

Write-Host ""
Write-Status "----------------------------------------------------------------" DarkGray
Write-Status "STEP 2: DELETE EXISTING STUDENT USERS" Cyan
Write-Status "----------------------------------------------------------------" DarkGray

& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $DeleteScript
$deleteCode = $LASTEXITCODE

if ($deleteCode -ne 0) {
    Write-Host ""
    Write-Status "WORKFLOW STOPPED: deletion failed (exit code $deleteCode)." Red
    Write-Status "Import was NOT started." Yellow
    Read-Host "Press Enter to exit"
    exit $deleteCode
}

# STEP 3
Write-Host ""
Write-Status "----------------------------------------------------------------" DarkGray
Write-Status "STEP 3: IMPORT FRESH USERS" Cyan
Write-Status "----------------------------------------------------------------" DarkGray

& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $ImportScript
$importCode = $LASTEXITCODE

if ($importCode -ne 0) {
    Write-Host ""
    Write-Status "WORKFLOW FINISHED WITH ERROR: import failed (exit code $importCode)." Red
    Write-Status "Review the IMPORT report and the DELETED-STUDENTS backup CSV in this folder." Yellow
    Read-Host "Press Enter to exit"
    exit $importCode
}

Write-Host ""
Write-Status "================================================================" Green
Write-Status " SUCCESS: VERIFY -> DELETE -> IMPORT completed successfully." Green
Write-Status "================================================================" Green
Write-Host ""
Read-Host "Press Enter to exit"
exit 0
