@echo off
REM *** RETIRED as of 2026-08-11 ***
REM This pushed Windows PC -> cloud, back when the Windows PC was the
REM authoritative source. That is no longer true: the cloud VM
REM (offseterp.duckdns.org) is now production, and this Windows PC is
REM development-only. Running this script would OVERWRITE real production
REM data with the stale/older Windows copy.
REM
REM If you actually need this direction again for some reason, pass "force"
REM as the first argument. Otherwise use sync_standby_from_primary.bat
REM (primary -> standby) or pull_db_from_cloud.bat (cloud -> local, read-only)
REM instead. See DEPLOY_CLOUD.md section 6.
if /i not "%~1"=="force" (
    echo ============================================
    echo  This script is retired - see comment block at top of file.
    echo  Re-run with "force" as the first argument if you really mean it.
    echo ============================================
    goto :eof
)

setlocal

REM Pass "silent" as the second argument (first is "force") to skip the
REM "press any key" prompt at the end, since no one is there to press it.
set "SILENT=%~2"

REM ---- Configuration ----
set "KEY_PATH=%USERPROFILE%\.ssh\offset-erp-oracle-a1"
set "VM_HOST=offseterp.duckdns.org"
set "VM_USER=ubuntu"
set "REPO_DIR=%~dp0..\.."

REM Only the stdlib sqlite3 module is needed here (no third-party packages),
REM so the project's venv is just one option, not a requirement. Try a few
REM candidates and use whichever actually runs -- "python" on PATH can be a
REM broken/relocated venv shim on some machines (fails with a pyvenv.cfg
REM error), so existence alone isn't enough; the "py" launcher is more
REM reliable where available.
set "PYTHON="
for %%P in ("%REPO_DIR%\.venv\Scripts\python.exe" "py" "python") do (
    if not defined PYTHON (
        %%P -c "import sqlite3" >nul 2>&1
        if not errorlevel 1 set "PYTHON=%%~P"
    )
)
if not defined PYTHON (
    echo No working Python interpreter found ^(tried the project venv, "py", and "python"^).
    echo Install Python from https://python.org and ensure it's on PATH, then try again.
    goto :end
)

set "TMP_DB=%TEMP%\offset_erp_sync_db.sqlite3"
set "TMP_MEDIA=%TEMP%\offset_erp_sync_media.tar.gz"
set "LOG_FILE=%REPO_DIR%\backups\sync_to_cloud.log"

if not exist "%REPO_DIR%\backups" mkdir "%REPO_DIR%\backups"
echo [%date% %time%] Starting sync >> "%LOG_FILE%"

echo ============================================
echo  Offset ERP - Sync production data to cloud
echo ============================================
echo.

echo [1/5] Taking a safe live snapshot of db.sqlite3...
"%PYTHON%" -c "import sqlite3; s=sqlite3.connect(r'%REPO_DIR%\db.sqlite3'); d=sqlite3.connect(r'%TMP_DB%'); s.backup(d); d.close(); s.close(); print('snapshot ok')"
if errorlevel 1 (
    echo Failed to create database snapshot. Aborting.
    echo [%date% %time%] Sync FAILED - db snapshot step >> "%LOG_FILE%"
    goto :end
)

echo [2/5] Packing media folder...
"%SystemRoot%\System32\tar.exe" -czf "%TMP_MEDIA%" -C "%REPO_DIR%" media
if errorlevel 1 (
    echo Failed to pack media folder. Aborting.
    echo [%date% %time%] Sync FAILED - media pack step >> "%LOG_FILE%"
    goto :end
)

echo [3/5] Uploading to the cloud server...
scp -i "%KEY_PATH%" -o StrictHostKeyChecking=accept-new "%TMP_DB%" %VM_USER%@%VM_HOST%:~/sync_db.sqlite3
if errorlevel 1 goto :upload_failed
scp -i "%KEY_PATH%" -o StrictHostKeyChecking=accept-new "%TMP_MEDIA%" %VM_USER%@%VM_HOST%:~/sync_media.tar.gz
if errorlevel 1 goto :upload_failed

echo [4/5] Loading data into the running deployment...
scp -i "%KEY_PATH%" -o StrictHostKeyChecking=accept-new "%~dp0remote_sync.sh" %VM_USER%@%VM_HOST%:~/remote_sync.sh
if errorlevel 1 goto :upload_failed
REM Strip any CR bytes on the Linux side before running it -- Windows git
REM configs (core.autocrlf) can silently give this file CRLF line endings
REM on this machine, which breaks bash parsing on the VM. Sanitizing here
REM means it works no matter what this machine's local file looks like.
ssh -i "%KEY_PATH%" -o StrictHostKeyChecking=accept-new %VM_USER%@%VM_HOST% "sed -i 's/\r$//' ~/remote_sync.sh && bash ~/remote_sync.sh"
if errorlevel 1 (
    echo Remote load step failed - check the output above.
    echo [%date% %time%] Sync FAILED - remote load step >> "%LOG_FILE%"
    goto :end
)

echo [5/5] Cleaning up local temp files...
del "%TMP_DB%" "%TMP_MEDIA%" 2>nul

echo.
echo Done. https://%VM_HOST% now has your latest production data.
echo [%date% %time%] Sync succeeded >> "%LOG_FILE%"
goto :end

:upload_failed
echo Upload to the cloud server failed. Check your internet connection
echo and that %VM_HOST% is reachable, then try again.
echo [%date% %time%] Sync FAILED - upload step >> "%LOG_FILE%"
goto :end

:end
del "%TMP_DB%" "%TMP_MEDIA%" 2>nul
echo.
if /i not "%SILENT%"=="silent" pause
