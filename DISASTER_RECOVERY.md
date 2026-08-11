# Disaster Recovery — If This PC Crashes / Is Replaced

**As of 2026-08-11, the Windows PC is dev-only.** Production is the cloud VM
at `offseterp.duckdns.org` (Oracle A1 instance). This doc originally assumed
the PC was the authoritative primary — it no longer is. See
`DEPLOY_CLOUD.md` section 6 for the current architecture.

This covers what to do if the Windows PC dies, is replaced, or otherwise
becomes unavailable.

## What's actually at risk

Most of this setup lives in places that survive a PC crash on their own —
and now that production runs on the cloud VM independently of this PC, a PC
crash is a *development environment* loss, not a data-loss event:

| Thing | Where it really lives | At risk if this PC dies? |
|---|---|---|
| App source code | GitHub (`EngrMahmood/Offset_Printing_ERP`) | No |
| Production server + database | Oracle Cloud A1 instance (`offseterp.duckdns.org`), independent VM | No — keeps running |
| Standby server + database | Oracle Cloud old VM (`offseterpbackup.duckdns.org`), independent VM, kept in sync via `scripts\cloud\standby\sync_standby_from_primary.bat` | No — keeps running |
| Cloud backups | Google Drive (`offseterp@gmail.com`) | No |
| This PC's local `db.sqlite3` | This PC (dev/test data only — not production) | Yes, but not production data |
| **SSH private key (primary)** | `%USERPROFILE%\.ssh\offset-erp-oracle-a1` | **Yes — only copy** |
| **SSH private key (standby)** | `%USERPROFILE%\.ssh\offset-erp-oracle.key` | **Yes — only copy** |
| **Android signing keystore** | `android-twa\android.keystore` + password | **Yes — only copy, and irreplaceable** |

The bolded "only copy" items are the real single points of failure — back
those up **before** anything happens, not after. Losing this PC does **not**
put production data at risk; it only costs you your dev environment and,
until restored, your ability to deploy/sync from a PC.

## Do this now, before any crash

1. **Back up both SSH keys**: copy `%USERPROFILE%\.ssh\offset-erp-oracle-a1`
   (primary) and `%USERPROFILE%\.ssh\offset-erp-oracle.key` (standby) to a
   password manager or another secure, separate location (a second device,
   encrypted USB drive, etc.). Anyone with these files can access the cloud
   servers, so treat them like passwords, not regular files.
2. **Back up the Android keystore**: same treatment for
   `android-twa\android.keystore` and `android-twa\keystore_password.txt`.
   If these are lost, you can never publish another update to the existing
   Play Store app listing — there is no recovery path for this one.
3. **Confirm OneDrive is actually syncing** — open OneDrive's tray icon and
   check `Offset ERP DB Backup` shows a green checkmark (fully synced), not
   a spinning/paused icon. This only backs up whatever local dev data is on
   this PC — it is not a production backup.

## If this PC crashes and you're setting up a new one

### 1. Get the app source code back
```bash
git clone https://github.com/EngrMahmood/Offset_Printing_ERP.git
cd Offset_Printing_ERP
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
```
Two things `requirements.txt` can't cover, since they're not Python packages:
- **Python 3.12** — install this version before creating the venv.
- **Redis** (v5+) — install and run it as a Windows service; required for
  chat/notifications/caching. `channels`/`cache` in `settings.py` are
  already set to talk RESP2 for compatibility with older Redis 5.x.

### 2. Get a local dev copy of the database (optional)
Production data lives on the cloud VM and was **not** affected by the PC
crash — there's nothing to "recover" there. If you want a fresh local copy
for dev/testing on the new PC, use the read-only pull script once the SSH
key is restored (next step):
```bash
scripts\cloud\pull_db_from_cloud.bat
```
This saves a timestamped copy under `backups\from_cloud\` — do **not**
overwrite the live cloud database from this PC (`sync_db_to_cloud.bat` is
retired for exactly this reason; see `DEPLOY_CLOUD.md` section 6).

### 3. Restore the SSH key(s)
Copy your backed-up `offset-erp-oracle-a1` (primary) to
`%USERPROFILE%\.ssh\offset-erp-oracle-a1`, and `offset-erp-oracle.key`
(standby) to `%USERPROFILE%\.ssh\offset-erp-oracle.key` on the new PC. If
you didn't back one up and it's truly gone, you'll need to add a new key via
**Oracle Cloud Console → Compute → Instances → [instance name] → Console
Connection** (browser-based serial console, doesn't need SSH) to log in and
add a new public key to `~/.ssh/authorized_keys` manually — do this for
whichever instance's key you lost.

### 4. Restore the Android keystore
Copy `android.keystore` and `keystore_password.txt` back into `android-twa/`
from your backup. Without these, `bubblewrap build` can still produce a
*new* signing key, but Google Play will reject it as a different
app — the original Play Store listing becomes permanently un-updatable.

### 5. Re-register the Windows Scheduled Tasks
These don't survive a PC replacement — recreate them from the XML files
already in `scripts/` (organized into `cloud/`, `backup/`, `server/`
subfolders — see each folder for what belongs there):
```bash
schtasks /Create /F /TN "Offset ERP Cloud Sync" /XML "scripts\cloud\ERP_CloudSync_Task.xml"
schtasks /Create /F /TN "Offset ERP Auto Backup" /XML "scripts\backup\ERP_AutoBackup_Task.xml"
schtasks /Create /F /TN "Offset ERP Server" /XML "scripts\server\ERP_Server_Startup_Task.xml"
```
Only register "Offset ERP Cloud Sync" if the Windows PC is still the
primary data source — once the cloud VM is primary, that task should stay
disabled (see `DEPLOY_CLOUD.md` §6); use `scripts\cloud\pull_db_from_cloud.bat`
instead if you want a local copy.
(Check each XML's `<Command>` path matches where you cloned the repo on the
new PC — they may reference the old PC's drive letter/path.)

### 6. Get the Windows server running again
Follow `DEPLOYMENT.md` for the normal server startup process (this hasn't
changed — same as before any of the cloud work).

### 7. Verify
- Log in locally against the dev server, confirm it comes up.
- Do **not** run `sync_db_to_cloud.bat` to "push data back" — that
  direction is retired and would overwrite real production data with stale
  dev data. If you need current data locally, use
  `scripts\cloud\pull_db_from_cloud.bat` instead (cloud → PC, read-only).
- Confirm `https://offseterp.duckdns.org` is still reachable (it should be
  — the cloud VM never went down, it's independent of this PC).

## The cloud VM itself never needs "recovery" from a PC crash

Worth remembering: **the live site at `https://offseterp.duckdns.org` keeps
running the whole time**, regardless of what happens to this PC — it's a
fully independent server. A PC crash affects your ability to *develop and
sync new changes*, not the live site's uptime.
