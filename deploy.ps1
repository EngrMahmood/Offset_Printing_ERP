# Deploy Offset ERP to the development server (Windows).
# Backup the database first, then run from PowerShell:
#   .\deploy.ps1
#
# If the window closes too fast, run:
#   powershell -NoExit -File .\deploy.ps1
#
# Options are passed through to scripts/deploy_dev.py, for example:
#   .\deploy.ps1 --git-pull
#   .\deploy.ps1 --run-tests

$ErrorActionPreference = 'Stop'
$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ProjectRoot

$LogDir = Join-Path $ProjectRoot 'logs'
if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir | Out-Null
}
$LogFile = Join-Path $LogDir ("deploy_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date))

$Python = $null
if ($env:VIRTUAL_ENV) {
    $venvPython = Join-Path $env:VIRTUAL_ENV 'Scripts\python.exe'
    if (Test-Path $venvPython) {
        $Python = $venvPython
    }
}
if (-not $Python) {
    $Python = 'python'
}

$deployArgs = @('scripts/deploy_dev.py', '--confirm-backup') + $args
$exitCode = 0

Write-Host ""
Write-Host "Offset ERP deploy"
Write-Host "Log file: $LogFile"
Write-Host ""

$previousErrorAction = $ErrorActionPreference
$ErrorActionPreference = 'Continue'

try {
    # Python/pip often write progress to stderr; do not treat that as a PowerShell error.
    $output = & $Python @deployArgs 2>&1
    $output | Tee-Object -FilePath $LogFile
    if ($null -ne $LASTEXITCODE -and $LASTEXITCODE -ne 0) {
        $exitCode = $LASTEXITCODE
    }
}
catch {
    $_ | Out-String | Tee-Object -FilePath $LogFile -Append
    Write-Host "Deploy failed: $_" -ForegroundColor Red
    $exitCode = 1
}
finally {
    $ErrorActionPreference = $previousErrorAction
}

Write-Host ""
if ($exitCode -eq 0) {
    Write-Host "Deploy finished successfully." -ForegroundColor Green
}
else {
    Write-Host "Deploy finished with errors. Exit code: $exitCode" -ForegroundColor Red
}
Write-Host "Full output saved to: $LogFile"
Write-Host ""
Read-Host "Press Enter to close"

exit $exitCode
