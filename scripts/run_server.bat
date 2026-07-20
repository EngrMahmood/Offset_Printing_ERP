@echo off
REM ============================================================
REM  Offset ERP - Server launcher for Task Scheduler / service.
REM  Runs the Django server in THIS process (no new windows), so it
REM  works when no user is logged in. Keeping this process alive keeps
REM  the in-process auto-backup running.
REM ============================================================

set "PROJECT_DIR=E:\Offset_Printing_ERP"

REM --- Python from the virtual environment (matches your (.venv) prompt).
REM     If your venv lives elsewhere, edit this path. If you don't use a
REM     venv, set:  set "PYTHON=python"
set "PYTHON=%PROJECT_DIR%\.venv\Scripts\python.exe"
if not exist "%PYTHON%" set "PYTHON=python"

set "BIND=192.168.88.30:8000"

cd /d "%PROJECT_DIR%"

echo [%date% %time%] Starting ERP server on %BIND% ... >> "%PROJECT_DIR%\backups\server.log"
"%PYTHON%" manage.py runserver %BIND% --noreload >> "%PROJECT_DIR%\backups\server.log" 2>&1
echo [%date% %time%] Server process exited (code %ERRORLEVEL%). >> "%PROJECT_DIR%\backups\server.log"
