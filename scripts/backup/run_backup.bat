@echo off
REM ============================================================
REM  Offset ERP - Daily Database Backup
REM  Runs the Django management command `run_backup`.
REM  Called by Windows Task Scheduler (see ERP_AutoBackup_Task.xml).
REM  Runs completely independently of the ERP web server.
REM ============================================================

REM --- Production project directory (matches "start server.bat") ---
set "PROJECT_DIR=E:\Offset_Printing_ERP"

REM --- Python executable ---
REM If you use a virtual environment, set the full path to its python.exe, e.g.:
REM   set "PYTHON=E:\Offset_Printing_ERP\venv\Scripts\python.exe"
REM Otherwise this uses the python found on PATH:
set "PYTHON=python"

cd /d "%PROJECT_DIR%"

echo [%date% %time%] Starting scheduled backup... >> "%PROJECT_DIR%\backups\auto_backup.log"
"%PYTHON%" manage.py run_backup >> "%PROJECT_DIR%\backups\auto_backup.log" 2>&1
echo [%date% %time%] Finished (exit code %ERRORLEVEL%). >> "%PROJECT_DIR%\backups\auto_backup.log"
