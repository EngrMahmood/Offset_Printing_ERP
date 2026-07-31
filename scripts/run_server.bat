@echo off
REM ============================================================
REM  Offset ERP - Server launcher for Task Scheduler / service.
REM  Runs the Django server in THIS process (no new windows), so it
REM  works when no user is logged in. Keeping this process alive keeps
REM  the in-process auto-backup running.
REM ============================================================

set "PROJECT_DIR=E:\Offset_Printing_ERP"

REM --- No venv is used on this machine; run with the system Python.
set "PYTHON=python"

REM --- Single HTTPS listener on 8000 (self-signed cert; see DEPLOYMENT.md
REM     step 6). WebRTC camera/mic access requires a secure context, so this
REM     replaces the old plain-HTTP runserver entirely. Cert paths are
REM     relative to PROJECT_DIR because Twisted's endpoint-string parser
REM     splits on ':' and a Windows drive letter (E:\...) breaks it.
set "BIND=192.168.88.30:8000"

cd /d "%PROJECT_DIR%"

REM --- Companion redirect listener on port 80: catches plain-HTTP requests
REM     (old bookmarks, or someone typing the bare IP) and 301s them to the
REM     HTTPS site above, since a TLS-only socket can't itself explain
REM     anything to a plain-HTTP client. Needs admin rights to bind :80,
REM     which this task already runs with (RunLevel HighestAvailable).
start "" /min "%PYTHON%" scripts\http_redirect_server.py >> "%PROJECT_DIR%\backups\http_redirect.log" 2>&1

echo [%date% %time%] Starting ERP server (HTTPS) on %BIND% ... >> "%PROJECT_DIR%\backups\server.log"
"%PYTHON%" -m daphne -e ssl:8000:privateKey=certs/chat_key.pem:certKey=certs/chat_cert.pem:interface=192.168.88.30 Offset_ERP.asgi:application >> "%PROJECT_DIR%\backups\server.log" 2>&1
echo [%date% %time%] Server process exited (code %ERRORLEVEL%). >> "%PROJECT_DIR%\backups\server.log"
