# Deploy Offset ERP to the development server (Windows).
# Backup the database first, then run:
#   .\deploy.ps1
#
# Options are passed through to scripts/deploy_dev.py, for example:
#   .\deploy.ps1 --git-pull
#   .\deploy.ps1 --run-tests

$ErrorActionPreference = 'Stop'
$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ProjectRoot

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
& $Python @deployArgs
exit $LASTEXITCODE
