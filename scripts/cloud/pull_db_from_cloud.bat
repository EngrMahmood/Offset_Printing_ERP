@echo off
REM Pulls a read-only copy of the live database (and media) DOWN from the
REM Oracle cloud deployment to this PC. Does NOT touch or restart anything
REM on the VM - safe to run anytime, as many times as you like.
REM
REM Double-click this file to run it.

setlocal

REM Pass "silent" as the first argument to skip the "press any key" prompt.
set "SILENT=%~1"

REM ---- Configuration ----
set "KEY_PATH=%USERPROFILE%\.ssh\offset-erp-oracle-a1"
set "VM_HOST=offseterp.duckdns.org"
set "VM_USER=ubuntu"
set "REPO_DIR=%~dp0..\.."
set "OUT_DIR=%REPO_DIR%\backups\from_cloud"

for /f "delims=" %%s in ('powershell -NoProfile -Command "Get-Date -Format yyyy-MM-dd_HHmmss"') do set "STAMP=%%s"

echo ============================================
echo  Offset ERP - Pull latest data FROM the cloud
echo ============================================
echo.

if not exist "%KEY_PATH%" (
    echo SSH key not found at %KEY_PATH%
    echo See DEPLOY_CLOUD.md / DISASTER_RECOVERY.md for how to restore it.
    goto :end
)

if not exist "%OUT_DIR%" mkdir "%OUT_DIR%"

REM Piping "ssh ... cat file" straight to a local redirect shows no progress
REM meter (the shell just sits there for the whole transfer). Instead, copy
REM the file onto the VM's own filesystem first, then scp it down -- scp
REM prints a live percent/speed/ETA bar, same as the upload side.
echo [1/2] Downloading database from %VM_HOST%...
ssh -i "%KEY_PATH%" -o StrictHostKeyChecking=accept-new %VM_USER%@%VM_HOST% "docker compose -f ~/offset-erp/docker-compose.yml cp web:/data/db.sqlite3 ~/pull_db.sqlite3"
if errorlevel 1 (
    echo Failed to stage the database on the VM. Check your internet connection
    echo and that %VM_HOST% is reachable, then try again.
    goto :end
)
scp -i "%KEY_PATH%" -o StrictHostKeyChecking=accept-new %VM_USER%@%VM_HOST%:~/pull_db.sqlite3 "%OUT_DIR%\db_from_cloud_%STAMP%.sqlite3"
if errorlevel 1 (
    echo Failed to download the staged database file. Check your internet
    echo connection and that %VM_HOST% is reachable, then try again.
    goto :end
)
ssh -i "%KEY_PATH%" -o StrictHostKeyChecking=accept-new %VM_USER%@%VM_HOST% "rm -f ~/pull_db.sqlite3"

echo [2/2] Downloading media folder from %VM_HOST%...
ssh -i "%KEY_PATH%" -o StrictHostKeyChecking=accept-new %VM_USER%@%VM_HOST% "docker compose -f ~/offset-erp/docker-compose.yml exec -T web tar -czf /app/pull_media.tar.gz -C /app media && docker compose -f ~/offset-erp/docker-compose.yml cp web:/app/pull_media.tar.gz ~/pull_media.tar.gz && docker compose -f ~/offset-erp/docker-compose.yml exec -T web rm -f /app/pull_media.tar.gz"
if errorlevel 1 (
    echo Failed to stage the media folder on the VM. Database download above
    echo may have still succeeded - check %OUT_DIR%.
    goto :end
)
scp -i "%KEY_PATH%" -o StrictHostKeyChecking=accept-new %VM_USER%@%VM_HOST%:~/pull_media.tar.gz "%OUT_DIR%\media_from_cloud_%STAMP%.tar.gz"
if errorlevel 1 (
    echo Failed to download the staged media file. Database download above
    echo may have still succeeded - check %OUT_DIR%.
    goto :end
)
ssh -i "%KEY_PATH%" -o StrictHostKeyChecking=accept-new %VM_USER%@%VM_HOST% "rm -f ~/pull_media.tar.gz"

echo.
echo Done. Saved to:
echo   %OUT_DIR%\db_from_cloud_%STAMP%.sqlite3
echo   %OUT_DIR%\media_from_cloud_%STAMP%.tar.gz
echo.
echo This is a READ-ONLY copy - the live cloud site was not changed.
echo Do not overwrite your real production db.sqlite3 with this unless you
echo have confirmed it is actually newer/more complete than what you have.

:end
echo.
if /i not "%SILENT%"=="silent" pause
