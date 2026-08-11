@echo off
REM One-shot full standby refresh: pulls the latest CODE from GitHub onto the
REM standby AND syncs the latest DATA (db + media) from the primary. Run this
REM any time you want the standby (offseterpbackup.duckdns.org) fully caught
REM up -- e.g. after pushing a change to GitHub, or periodically to keep the
REM failover candidate ready. Combines what deploy_from_github.bat does for
REM the primary with what sync_standby_from_primary.bat does for data, both
REM targeted at the standby.
REM
REM Double-click this file to run it.

setlocal

REM Pass "silent" as the first argument to skip the "press any key" prompt.
set "SILENT=%~1"

REM ---- Configuration ----
set "STANDBY_KEY=%USERPROFILE%\.ssh\offset-erp-oracle.key"
set "STANDBY_HOST=offseterpbackup.duckdns.org"
set "VM_USER=ubuntu"
set "REPO_DIR=%~dp0..\..\.."
set "LOG_FILE=%REPO_DIR%\backups\update_standby.log"

if not exist "%REPO_DIR%\backups" mkdir "%REPO_DIR%\backups"
echo [%date% %time%] Starting full standby update (code + data) >> "%LOG_FILE%"

echo ============================================
echo  Offset ERP - Full standby update (code + data)
echo ============================================
echo.

if not exist "%STANDBY_KEY%" (
    echo SSH key not found at %STANDBY_KEY%
    echo See DEPLOY_CLOUD.md / DISASTER_RECOVERY.md for how to restore it.
    goto :end
)

echo [1/2] Updating standby code from GitHub...
ssh -i "%STANDBY_KEY%" -o StrictHostKeyChecking=accept-new %VM_USER%@%STANDBY_HOST% "cd ~/offset-erp && git pull origin main && docker compose up -d --build web"
if errorlevel 1 (
    echo.
    echo Failed to update code on the standby - check the output above.
    echo Skipping the data sync so fresh data doesn't land on broken code.
    echo [%date% %time%] Full standby update FAILED - code step >> "%LOG_FILE%"
    goto :end
)

echo.
echo [2/2] Syncing database + media from the primary...
call "%~dp0sync_standby_from_primary.bat" silent

echo.
echo Done. The standby (https://%STANDBY_HOST%) now has the latest code AND
echo the latest data pulled from the primary. See the output above for the
echo data sync step's own success/failure detail (also logged separately to
echo backups\sync_standby.log).
echo [%date% %time%] Full standby update finished - see backups\sync_standby.log for the data step's result >> "%LOG_FILE%"

:end
echo.
if /i not "%SILENT%"=="silent" pause
