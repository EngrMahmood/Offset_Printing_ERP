# Deploying the Chat Module to Production (E:\Offset_Printing_ERP)

Context for whoever (or whichever Claude Code session) is running this on the
production machine: the `chat` app (real-time messaging, attachments, WebRTC
calling, docked popup windows) was added on the dev machine and pushed to
`origin/main`. It introduces new runtime infrastructure that the production
server doesn't have yet. This doc is the one-time deployment checklist.

## What's new and why it matters here

- **Django Channels + Daphne**: `manage.py runserver` now needs `daphne` in
  `INSTALLED_APPS` (already committed) to serve WebSockets — no script change
  needed for this part, `runserver` picks it up automatically.
- **Redis**: required for both the Channels layer (`CHANNEL_LAYERS`) and the
  cache backend (`CACHES`), both in `Offset_ERP/settings.py`. **Not yet
  installed on production** as of this writing.
- **New Python packages**: `channels`, `channels-redis`, `daphne`,
  `djangorestframework`, `redis`, `Pillow` — see `requirements.txt`.
- **New migrations + permission seed commands** for the `chat` app and the
  RBAC system that shipped alongside it.

## Step-by-step

### 1. Install Redis

Production uses the same **unofficial Redis-for-Windows port** already
validated on the dev machine (not Memurai — its free tier is
production-prohibited and auto-shuts-down after 10 days, which would cause
mysterious outages).

Download the MSI from https://github.com/tporadowski/redis/releases
(`Redis-x64-5.0.14.1` or newer). During install, check:
- "Add the Redis installation folder to the PATH environment variable"
- "Redis Windows service"

This registers Redis as a Windows Service listening on `127.0.0.1:6379`,
auto-starting on boot. Verify in `services.msc` that its Startup type is
**Automatic**. No config file changes needed — `Offset_ERP/settings.py`
already points at `redis://127.0.0.1:6379` and forces RESP2 protocol
(`protocol: 2` in `CACHES`/`CHANNEL_LAYERS`) for compatibility with this
older Redis version.

### 2. Pull code and install dependencies

```bash
cd /d E:\Offset_Printing_ERP
git pull origin main
python -m pip install -r requirements.txt
```

### 3. One-time migrate + permission seed

```bash
python manage.py migrate
python manage.py seed_access_control
python manage.py seed_chat_permissions
```

`seed_access_control` must run before `seed_chat_permissions` (the latter
looks up `Role` rows the former creates).

### 4. Manual smoke test before touching the scheduled task

Run `start server.bat` as usual. Confirm:
- The site loads normally.
- `/chat/` loads, and the Chat nav link appears (permission-gated).
- Two accounts can send a message and see it arrive live with no page
  refresh (confirms Redis + Channels are actually wired up, not just that
  the server started).

### 5. Fix the boot-time Scheduled Task

The existing task ("Offset ERP Server") wasn't starting the server
automatically at boot. Root cause: its action ran `start server.bat`, which
opens an interactive console window (`cmd /k`) — a task set to "Run whether
user is logged on or not" has no desktop to display that window on before
anyone logs in, so it silently does nothing.

Fix: a new script, `start_server_task.bat` (in the same
`C:\Users\Universal\OneDrive\Development\` folder as the original), runs
headless, resolves `python.exe` explicitly, waits up to 60s for Redis to
become reachable before starting Django, and logs everything to
`E:\Offset_Printing_ERP\logs\server_startup.log`. `start server.bat` itself
is untouched and still fine for manual double-click runs.

In `taskschd.msc` → "Offset ERP Server" → Properties:
- **Actions tab**: change the action to point at `start_server_task.bat`
  instead of `start server.bat`.
- **Settings tab**: uncheck **"Stop the task if it runs longer than: 3
  days"** — a Task Scheduler default that would kill a long-running server
  process even after startup itself is fixed.
- **Settings tab**: uncheck **"Start the task only if the computer is on AC
  power"** if it's checked.
- **General tab**: confirm **"Run whether user is logged on or not"** is
  still selected.

Then re-run `setup_server_task.local.bat` to re-register the task, and do a
real reboot test — restart the machine and check
`logs\server_startup.log` for a clean startup with Redis reachable, with no
one logged in.

### 6. Enable HTTPS for calling (self-signed certificate)

WebRTC's camera/microphone access (`getUserMedia`) only works in a browser
"secure context" — `https:` or `localhost`/`127.0.0.1`. Since this server is
reached over plain `http://192.168.88.30:8000` from other LAN PCs, calling
silently fails everywhere except the server machine itself. Fix: serve a
second Daphne listener over HTTPS with a self-signed certificate.

Generate a cert (PowerShell, using OpenSSL if available, or the .NET
`New-SelfSignedCertificate` cmdlet — either works; example with OpenSSL):

```bash
openssl req -x509 -newkey rsa:2048 -nodes -keyout chat_key.pem -out chat_cert.pem -days 3650 -subj "/CN=192.168.88.30" -addext "subjectAltName=IP:192.168.88.30"
```

Store `chat_key.pem`/`chat_cert.pem` **outside the git repo** (e.g. next to
`logs/` under `E:\Offset_Printing_ERP\certs\`) — never commit private keys.

Add a second Daphne bind alongside the existing plain-HTTP one in
`start_server_task.bat` / `start server.bat`:

```bash
daphne -b 0.0.0.0 -p 8443 -e ssl:8443:privateKey=E:\Offset_Printing_ERP\certs\chat_key.pem:certKey=E:\Offset_Printing_ERP\certs\chat_cert.pem Offset_ERP.asgi:application
```

(Keep the plain `:8000` listener running too, or run both from one process
via Daphne's multi-endpoint `-e` flags — either is fine.) `Offset_ERP/settings.py`
already trusts `https://192.168.88.30:8443` via `CSRF_TRUSTED_ORIGINS`
(overridable with the `CSRF_TRUSTED_ORIGINS` env var if the port/IP differ).

On each LAN PC, the first visit to `https://192.168.88.30:8443` shows a
"Your connection is not private" warning (expected for a self-signed cert on
an internal LAN) — click **Advanced → Proceed**. This is a one-time step per
browser per machine. Bookmark/share the `https://` link, not the old `http://`
one, so calling works.

## If something breaks

- **Chat page loads but messages don't send/appear live**: check Redis is
  actually running (`services.msc`, or `Test-NetConnection 127.0.0.1 -Port
  6379`).
- **Server won't start at all after `git pull`**: check
  `pip install -r requirements.txt` actually completed — `Pillow` in
  particular can fail to build on some systems without a matching wheel;
  if so, `pip install Pillow` alone first to see the real error.
- **Boot task still doesn't start the server**: check
  `logs\server_startup.log` — it will say exactly where it stopped (couldn't
  find `python.exe`, `cd` failed, Redis unreachable, etc.) rather than
  failing silently like before.
