@echo off
REM Refreshes the standby VM (offseterpbackup.duckdns.org, the old VM) with a
REM live snapshot pulled FROM the primary (offseterp.duckdns.org, the new A1
REM instance). This is the reverse-direction counterpart to
REM sync_db_to_cloud.bat, used now that the primary takes real production
REM writes directly and the old VM is a read-only standby, not a source of
REM truth. Safe to run anytime -- pulls a live snapshot from the primary via
REM Docker (no disruption there), then loads it into the standby.
REM
REM Double-click this file to run it.

setlocal

REM Pass "silent" as the first argument (used by the scheduled task) to skip
REM the "press any key" prompt at the end.
set "SILENT=%~1"

REM ---- Configuration ----
set "PRIMARY_KEY=%USERPROFILE%\.ssh\offset-erp-oracle-a1"
set "PRIMARY_HOST=offseterp.duckdns.org"
set "STANDBY_KEY=%USERPROFILE%\.ssh\offset-erp-oracle.key"
set "STANDBY_HOST=offseterpbackup.duckdns.org"
set "VM_USER=ubuntu"
set "REPO_DIR=%~dp0..\..\.."
set "LOG_FILE=%REPO_DIR%\backups\sync_standby.log"

if not exist "%REPO_DIR%\backups" mkdir "%REPO_DIR%\backups"
echo [%date% %time%] Starting standby sync >> "%LOG_FILE%"

echo ============================================
echo  Offset ERP - Refresh standby from primary
echo ============================================
echo.

if not exist "%PRIMARY_KEY%" (
    echo SSH key not found at %PRIMARY_KEY%
    goto :end
)
if not exist "%STANDBY_KEY%" (
    echo SSH key not found at %STANDBY_KEY%
    goto :end
)

echo [1/5] Staging a live snapshot on the primary...
REM Uses SQLite's online backup API (via a python one-liner inside the
REM container) rather than a raw "docker compose cp" of db.sqlite3 -- the
REM primary takes real live writes now (WAL mode), and a raw file copy can
REM grab a half-written/inconsistent page mid-transaction. The backup API
REM is safe to run against a live, actively-written database.
ssh -i "%PRIMARY_KEY%" -o StrictHostKeyChecking=accept-new %VM_USER%@%PRIMARY_HOST% "docker compose -f ~/offset-erp/docker-compose.yml exec -T web python -c \"import sqlite3; s=sqlite3.connect('/data/db.sqlite3'); d=sqlite3.connect('/tmp/standby_pull_db.sqlite3'); s.backup(d); d.close(); s.close()\" && docker compose -f ~/offset-erp/docker-compose.yml cp web:/tmp/standby_pull_db.sqlite3 ~/standby_pull_db.sqlite3 && docker compose -f ~/offset-erp/docker-compose.yml exec -T web rm -f /tmp/standby_pull_db.sqlite3 && docker compose -f ~/offset-erp/docker-compose.yml exec -T web tar -czf /app/standby_pull_media.tar.gz -C /app media && docker compose -f ~/offset-erp/docker-compose.yml cp web:/app/standby_pull_media.tar.gz ~/standby_pull_media.tar.gz && docker compose -f ~/offset-erp/docker-compose.yml exec -T web rm -f /app/standby_pull_media.tar.gz"
if errorlevel 1 (
    echo Failed to stage the snapshot on the primary. Aborting.
    echo [%date% %time%] Standby sync FAILED - stage step >> "%LOG_FILE%"
    goto :end
)

echo [2/5] Downloading snapshot to this PC...
set "TMP_DB=%TEMP%\offset_erp_standby_db.sqlite3"
set "TMP_MEDIA=%TEMP%\offset_erp_standby_media.tar.gz"
scp -i "%PRIMARY_KEY%" -o StrictHostKeyChecking=accept-new %VM_USER%@%PRIMARY_HOST%:~/standby_pull_db.sqlite3 "%TMP_DB%"
if errorlevel 1 goto :fail_download
scp -i "%PRIMARY_KEY%" -o StrictHostKeyChecking=accept-new %VM_USER%@%PRIMARY_HOST%:~/standby_pull_media.tar.gz "%TMP_MEDIA%"
if errorlevel 1 goto :fail_download
ssh -i "%PRIMARY_KEY%" -o StrictHostKeyChecking=accept-new %VM_USER%@%PRIMARY_HOST% "rm -f ~/standby_pull_db.sqlite3 ~/standby_pull_media.tar.gz"

echo [3/5] Uploading snapshot to the standby...
scp -i "%STANDBY_KEY%" -o StrictHostKeyChecking=accept-new "%TMP_DB%" %VM_USER%@%STANDBY_HOST%:~/sync_db.sqlite3
if errorlevel 1 goto :fail_upload
scp -i "%STANDBY_KEY%" -o StrictHostKeyChecking=accept-new "%TMP_MEDIA%" %VM_USER%@%STANDBY_HOST%:~/sync_media.tar.gz
if errorlevel 1 goto :fail_upload

echo [4/5] Loading data into the standby deployment...
scp -i "%STANDBY_KEY%" -o StrictHostKeyChecking=accept-new "%~dp0..\remote_sync.sh" %VM_USER%@%STANDBY_HOST%:~/remote_sync.sh
if errorlevel 1 goto :fail_upload
ssh -i "%STANDBY_KEY%" -o StrictHostKeyChecking=accept-new %VM_USER%@%STANDBY_HOST% "sed -i 's/\r$//' ~/remote_sync.sh && bash ~/remote_sync.sh"
if errorlevel 1 (
    echo Remote load step failed on the standby - check the output above.
    echo [%date% %time%] Standby sync FAILED - remote load step >> "%LOG_FILE%"
    goto :end
)

echo [5/5] Cleaning up local temp files...
del "%TMP_DB%" "%TMP_MEDIA%" 2>nul

echo.
echo Done. The standby (https://%STANDBY_HOST%) now mirrors the primary.
echo [%date% %time%] Standby sync succeeded >> "%LOG_FILE%"
goto :end

:fail_download
echo Failed to download the staged snapshot from the primary. Check your
echo internet connection and that %PRIMARY_HOST% is reachable, then try again.
echo [%date% %time%] Standby sync FAILED - download step >> "%LOG_FILE%"
goto :end

:fail_upload
echo Failed to upload to the standby. Check your internet connection and
echo that %STANDBY_HOST% is reachable, then try again.
echo [%date% %time%] Standby sync FAILED - upload step >> "%LOG_FILE%"
goto :end

:end
del "%TMP_DB%" "%TMP_MEDIA%" 2>nul
echo.
if /i not "%SILENT%"=="silent" pause
