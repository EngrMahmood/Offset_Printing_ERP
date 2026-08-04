@echo off
REM Pulls the latest code from GitHub onto the Oracle cloud server and
REM redeploys it. Double-click this file to run it.
REM
REM Does NOT touch the Windows production server or its database — this
REM only updates the cloud copy at https://offseterp.duckdns.org.

setlocal

set "SILENT=%~1"
set "KEY_PATH=%USERPROFILE%\.ssh\offset-erp-oracle.key"
set "VM_HOST=offseterp.duckdns.org"
set "VM_USER=ubuntu"

echo ============================================
echo  Offset ERP - Deploy latest from GitHub
echo ============================================
echo.

if not exist "%KEY_PATH%" (
    echo SSH key not found at %KEY_PATH%
    echo See DEPLOY_CLOUD.md / DISASTER_RECOVERY.md for how to restore it.
    goto :end
)

ssh -i "%KEY_PATH%" -o StrictHostKeyChecking=accept-new %VM_USER%@%VM_HOST% "bash ~/offset-erp/scripts/deploy_update.sh"
if errorlevel 1 (
    echo.
    echo Deploy failed - check the output above.
    goto :end
)

echo.
echo Done. https://%VM_HOST% is running the latest code from GitHub.

:end
echo.
if /i not "%SILENT%"=="silent" pause
