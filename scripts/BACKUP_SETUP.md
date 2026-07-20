# Auto-Backup — How It Works

Auto-backup now runs **inside the running server process**, the same way on every
machine (dev or production). As long as the Django server is running, a background
thread checks once a minute and creates the daily backup at the scheduled time
(Backup > Settings). No per-machine configuration and no separate scheduler needed.

## Why it wasn't working in production
Production starts the server with `runserver ... --noreload`. The old code only
started the scheduler when Django set `RUN_MAIN=true`, which **never happens with
`--noreload`** — so the thread never started and no auto-backup ever ran. That guard
has been fixed: the scheduler now starts under `runserver` (with or without
`--noreload`) and under WSGI servers (gunicorn/waitress/uWSGI/IIS), and only skips
the redundant reloader *parent* process to avoid running twice.

## What you need to do
Nothing in the code — just make sure the **server process stays running**. The backup
only happens while the server is up.

### Backup "whether a user is logged in or not"
A server started by double-clicking `start server.bat` dies when that user logs off.
To keep it running across logoff/reboot (so backups keep happening), run the server
as a background service or a startup task. Pick one:

**Option A — Task Scheduler at boot (ready-made, recommended):**
Two files are provided:
- `scripts/run_server.bat` — launches the server in-process (no pop-up windows,
  works with no user logged in), logging to `backups/server.log`.
- `scripts/ERP_Server_Startup_Task.xml` — a boot-time task set to "run whether
  logged on or not", highest privileges, auto-restart on failure.

Import it (Admin Command Prompt; substitute your Windows account + password):
```
schtasks /Create /TN "Offset ERP Server" /XML "E:\Offset_Printing_ERP\scripts\ERP_Server_Startup_Task.xml" /RU "%COMPUTERNAME%\YourUser" /RP "YourPassword"
```
Then start it now without rebooting, and confirm it's running:
```
schtasks /Run /TN "Offset ERP Server"
schtasks /Query /TN "Offset ERP Server"
```
Check `backups\server.log` and open the ERP in a browser to confirm. First edit
`scripts\run_server.bat` if your venv path or bind address differs.

Or do it in the GUI: Task Scheduler > Create Task > General: "Run whether user is
logged on or not" + "Run with highest privileges"; Triggers: "At startup";
Actions: Start a program > `E:\Offset_Printing_ERP\scripts\run_server.bat`.

**Option B — Windows Service via NSSM (most robust):**
```
nssm install OffsetERP "C:\path\to\python.exe" "E:\Offset_Printing_ERP\manage.py runserver 192.168.88.30:8000 --noreload"
nssm set OffsetERP AppDirectory E:\Offset_Printing_ERP
nssm start OffsetERP
```
A service runs regardless of login and restarts automatically.

## Verifying
- Set the backup time (Backup > Settings) a couple of minutes ahead and watch the
  Backup dashboard — a new **Auto / Success** row should appear.
- Or force one immediately in a clean process:
  ```
  cd E:\Offset_Printing_ERP
  python manage.py run_backup
  ```
  This is the exact code the scheduler runs; it should produce a Success row.

## Reliability behaviour (built in)
- One successful Auto backup per day; after it succeeds the scheduler stops retrying
  until the next day.
- If the server starts *after* the scheduled time and today's backup hasn't run yet,
  it runs within ~1 minute (catch-up).
- If a backup fails, it retries at most every 15 minutes — never the per-minute
  "storm" that filled the dashboard with Pending rows before.
- A run that is interrupted (e.g. server restart) is recorded as Failed, not left
  hanging as Pending.

## Optional: OS-driven backups instead of in-process
If you would rather NOT depend on the server being up, you can disable the in-process
scheduler and drive backups from Windows Task Scheduler instead:
- In `settings.py`: `BACKUP_INPROCESS_SCHEDULER = False`
- Import `scripts/ERP_AutoBackup_Task.xml` (runs `scripts/run_backup.bat`, i.e.
  `manage.py run_backup`, daily). Do not use both at once.
