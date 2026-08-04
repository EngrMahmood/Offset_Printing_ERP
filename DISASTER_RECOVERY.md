# Disaster Recovery — If This PC Crashes / Is Replaced

This covers what to do if the Windows PC running the ERP (and used to manage
the cloud deployment) dies, is replaced, or otherwise becomes unavailable.

## What's actually at risk

Most of this setup lives in places that survive a PC crash on their own:

| Thing | Where it really lives | At risk if this PC dies? |
|---|---|---|
| App source code | GitHub (`EngrMahmood/Offset_Printing_ERP`) | No |
| Cloud server itself | Oracle Cloud (independent VM) | No — keeps running |
| Cloud database/media | Oracle VM's disk | No |
| Cloud backups | Google Drive (`offseterp@gmail.com`) | No |
| **Production database (freshest copy)** | **This PC's `db.sqlite3`** | **Yes** |
| Daily local backups | `E:\Offset ERP DB Backup`, mirrored to OneDrive | Only if OneDrive wasn't syncing |
| **SSH private key** (VM access) | `%USERPROFILE%\.ssh\offset-erp-oracle.key` | **Yes — only copy** |
| **Android signing keystore** | `android-twa\android.keystore` + password | **Yes — only copy, and irreplaceable** |

The two bolded "only copy" items are the real single points of failure —
back those up **before** anything happens, not after.

## Do this now, before any crash

1. **Back up the SSH key**: copy `%USERPROFILE%\.ssh\offset-erp-oracle.key`
   to a password manager or another secure, separate location (a second
   device, encrypted USB drive, etc.). Anyone with this file can access the
   cloud server, so treat it like a password, not a regular file.
2. **Back up the Android keystore**: same treatment for
   `android-twa\android.keystore` and `android-twa\keystore_password.txt`.
   If these are lost, you can never publish another update to the existing
   Play Store app listing — there is no recovery path for this one.
3. **Confirm OneDrive is actually syncing** — open OneDrive's tray icon and
   check `Offset ERP DB Backup` shows a green checkmark (fully synced), not
   a spinning/paused icon. This is what makes your local daily backups
   survive a PC loss.

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

### 2. Get the production database back
The cloud VM has the most recent *synced* copy (as of the last nightly
sync — could be up to 24h stale). Pull it back down:
```bash
ssh -i <path-to-restored-ssh-key> ubuntu@offseterp.duckdns.org \
  "docker compose -f ~/offset-erp/docker-compose.yml exec web cat /data/db.sqlite3" > db.sqlite3
```
If you have a **more recent** local backup recovered from OneDrive
(`Offset ERP DB Backup` folder), that one is fresher — use that instead as
`db.sqlite3` in the project root. Compare timestamps and pick the newer one.

Also recover `media/` the same way — either from OneDrive's backup ZIPs
(the app's `/backup/` includes media if that setting was enabled) or by
copying it back down from the VM's `offset-erp_media` Docker volume.

### 3. Restore the SSH key
Copy your backed-up `offset-erp-oracle.key` to
`%USERPROFILE%\.ssh\offset-erp-oracle.key` on the new PC. If you didn't
back it up and it's truly gone, you'll need to add a new key via **Oracle
Cloud Console → Compute → Instances → offset-erp-server → Console
Connection** (browser-based serial console, doesn't need SSH) to log in and
add a new public key to `~/.ssh/authorized_keys` manually.

### 4. Restore the Android keystore
Copy `android.keystore` and `keystore_password.txt` back into `android-twa/`
from your backup. Without these, `bubblewrap build` can still produce a
*new* signing key, but Google Play will reject it as a different
app — the original Play Store listing becomes permanently un-updatable.

### 5. Re-register the Windows Scheduled Tasks
These don't survive a PC replacement — recreate them from the XML files
already in `scripts/`:
```bash
schtasks /Create /F /TN "Offset ERP Cloud Sync" /XML "scripts\ERP_CloudSync_Task.xml"
schtasks /Create /F /TN "Offset ERP Auto Backup" /XML "scripts\ERP_AutoBackup_Task.xml"
schtasks /Create /F /TN "Offset ERP Server" /XML "scripts\ERP_Server_Startup_Task.xml"
```
(Check each XML's `<Command>` path matches where you cloned the repo on the
new PC — they may reference the old PC's drive letter/path.)

### 6. Get the Windows server running again
Follow `DEPLOYMENT.md` for the normal server startup process (this hasn't
changed — same as before any of the cloud work).

### 7. Verify
- Log in locally, confirm the restored data looks right (spot-check a few
  recent job cards against what you remember).
- Run `scripts\sync_db_to_cloud.bat` once to push the recovered data back
  up to the cloud copy, so both sides match again.
- Confirm `https://offseterp.duckdns.org` is still reachable (it should be
  — the cloud VM never went down, it's independent of this PC).

## The cloud VM itself never needs "recovery" from a PC crash

Worth remembering: **the live site at `https://offseterp.duckdns.org` keeps
running the whole time**, regardless of what happens to this PC — it's a
fully independent server. A PC crash affects your ability to *develop and
sync new changes*, not the live site's uptime.
