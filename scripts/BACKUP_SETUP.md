# Reliable Auto-Backup Setup (Windows Task Scheduler)

## Why
The old auto-backup ran as a background thread *inside the Django web process*.
Your `start server.bat` launches the server with `--noreload`, which means Django
never sets `RUN_MAIN=true`, so the scheduler thread was **never started at all** —
that is the exact reason `backup_backuphistory` had zero AUTO records. Rather than
depend on the web process, Windows Task Scheduler runs the backup on its own.

## Do I need the server running?
**No.** Once the scheduled task is set up, the backup runs on its own every day at
20:05. `manage.py run_backup` opens its own short-lived process, reads the database
file, and writes the zip — it does not need `start server.bat` to be running. The
only requirement is that the **PC is powered on** at 20:05 (and if it isn't, the
`StartWhenAvailable` setting makes it run at the next power-on).

## What was added
- `scripts/run_backup.bat` — runs `python manage.py run_backup`, logs to `backups/auto_backup.log`.
- `scripts/ERP_AutoBackup_Task.xml` — importable scheduled task, daily at 20:05.
- `manage.py run_backup` now records the run as **AUTO** by default (`--type MANUAL` to override).

## One-time setup

1. **Check the Python path.** Open `scripts\run_backup.bat`. If you run the ERP from a
   virtual environment, set `PYTHON` to that venv's `python.exe`. Otherwise leave it as
   `python` (must be on PATH).

2. **Test the batch file** by double-clicking `run_backup.bat`, then confirm a new
   zip appears in `E:\Offset ERP DB Backup` and a line was written to
   `backups\auto_backup.log`.

3. **Import the task** (Admin PowerShell / Command Prompt):
   ```
   schtasks /Create /TN "Offset ERP Auto Backup" /XML "E:\Offset_Printing_ERP\scripts\ERP_AutoBackup_Task.xml"
   ```
   Or: Task Scheduler → Action → **Import Task…** → select the XML.

4. **Run it once on demand** to verify:
   ```
   schtasks /Run /TN "Offset ERP Auto Backup"
   ```
   Then check the backup folder, `auto_backup.log`, and Supply Chain → backup history.

## In-process scheduler (disabled)
The old background-thread scheduler is now **off by default** — the Windows scheduled
task is the single source of truth, so there's no risk of duplicate backups. If you
ever want the in-process thread back on a dev machine, add to `settings.py`:
```
BACKUP_INPROCESS_SCHEDULER = True
```

## Notes
- `StartWhenAvailable` is on, so if the PC is off at 20:05 the task runs at the next
  opportunity (catch-up).
- To run even when no user is logged in, edit the task → General → "Run whether user is
  logged on or not" (Windows will ask for the account password).
- The in-process thread scheduler can stay as-is (harmless) or be disabled later; the
  Task Scheduler job is now the source of truth.
