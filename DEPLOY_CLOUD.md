# Deploying to the Cloud (Free) — Web + Android

This is the companion to `DEPLOYMENT.md` (which covers the existing Windows
LAN server — untouched by any of this). This doc covers standing up a
**separate, public** deployment on a free cloud VM and making it installable
on Android as a PWA.

**Live deployment**: `offset-erp-server` on Oracle Cloud (Mumbai region),
**https://offseterp.duckdns.org** — VM.Standard.E2.1.Micro (1 OCPU/1GB + 2GB
swap), Ubuntu 24.04, public IP `130.210.51.29`. Set up 2026-08-04. (Originally
launched on a `sslip.io` hostname, then moved to this free DuckDNS name —
same VM/cert-issuing process either way, see §1a below for either option.)

Nothing here changes how the Windows server runs. Do this in parallel, verify
it, and only decide about cutover once it's proven stable.

**If this PC (the one managing the cloud deployment) is ever lost/crashes**,
see `DISASTER_RECOVERY.md` — the live site keeps running regardless, but the
SSH key and Android signing keystore should be backed up *before* that
happens since they can't be recovered otherwise.

Database: this stack uses **SQLite** (same engine/tuning as the Windows
server), not Postgres — a good fit for a closed system with ~15-20 users, and
lighter on RAM on a small free-tier VM. Postgres is still available any time
by setting `DATABASE_URL` (see `Offset_ERP/settings.py`), if this ever needs
to scale beyond SQLite's comfort zone.

## 1. Provision the VM

Sign up for Oracle Cloud (Always Free tier) and create a VM instance:
- Shape: try the Ampere A1 (4 OCPU/24GB) free shape first; if "out of
  capacity" in your region (common — retry later or try a smaller OCPU/memory
  request), fall back to the free AMD micro shape (1 OCPU/1GB — smaller but
  reliably available, still fine at this scale).
- OS: Ubuntu 22.04/24.04 LTS.
- Open ports 80 and 443 in the VM's security list/network security group (in
  the console: subnet → Default Security List → Add Ingress Rules, source
  `0.0.0.0/0`, TCP, ports 80 and 443).
- **Also open those ports in the VM's own OS-level firewall** — Oracle's
  Ubuntu images ship with `iptables` rules that only allow SSH by default,
  *in addition to* the cloud security list above. Both layers must allow the
  port, or you'll get "connection refused"/timeouts even with the security
  list correctly configured:
  ```bash
  sudo iptables -I INPUT 4 -p tcp --dport 80 -j ACCEPT
  sudo iptables -I INPUT 5 -p tcp --dport 443 -j ACCEPT
  sudo netfilter-persistent save   # or: sudo apt install -y iptables-persistent
  ```
  (Insert before the final REJECT rule — check position with
  `sudo iptables -L INPUT -n --line-numbers` first.)
- Install Docker + Docker Compose on the VM (`curl -fsSL https://get.docker.com | sh`).
- **Add a swap file** if using the 1GB AMD micro shape — 1GB is tight for
  Redis + Django/Daphne + Nginx together under real traffic:
  ```bash
  sudo fallocate -l 2G /swapfile && sudo chmod 600 /swapfile
  sudo mkswap /swapfile && sudo swapon /swapfile
  echo "/swapfile none swap sw 0 0" | sudo tee -a /etc/fstab
  ```

### 1a. Free HTTPS domain without owning one

Two free options, both give a real hostname Let's Encrypt will issue a
trusted cert for (no domain purchase needed):

- **DuckDNS** (what's actually in use — `offseterp.duckdns.org`): sign in at
  duckdns.org with a Google/GitHub/Reddit account, claim a subdomain, point
  it at the VM's public IP from the dashboard. Free forever, lets you pick
  a real name instead of an IP-derived one.
- **sslip.io**: zero signup, resolves `<ip-with-dashes>.sslip.io` straight to
  that IP automatically (e.g. `130.210.51.29` → `130-210-51-29.sslip.io`).
  Faster to stand up, but the hostname is tied to the IP and isn't
  memorable/brandable.

Either way, issuing the cert is the same:

```bash
sudo apt-get install -y certbot
docker compose stop nginx   # free port 80 for the certbot standalone check
sudo certbot certonly --standalone --non-interactive --agree-tos \
  -m you@example.com -d your-chosen-hostname
# copy into place for docker-compose.yml's nginx volume mount:
sudo mkdir -p certs
sudo cp /etc/letsencrypt/live/your-chosen-hostname/fullchain.pem certs/
sudo cp /etc/letsencrypt/live/your-chosen-hostname/privkey.pem certs/
sudo chown $USER:$USER certs/*.pem
docker compose up -d
```

Certbot auto-schedules renewal. If you later get a real domain, just repoint
DNS at the VM and re-run certbot for that domain instead — same process. Also
update `ALLOWED_HOSTS`/`CSRF_TRUSTED_ORIGINS` in `.env` to match.

## 2. Get the code and data onto the VM

```bash
git clone <your repo url> offset-erp
cd offset-erp
cp .env.example .env   # fill in real SECRET_KEY, domain
```

To bring over existing data, copy a **copy** of the SQLite file — never the
live `db.sqlite3` the Windows server is actively writing to:

```bash
# On the Windows machine (with the app stopped, or via the Backup & Restore
# dashboard's export, for a consistent snapshot):
copy db.sqlite3 db.sqlite3.snapshot

# scp/rsync db.sqlite3.snapshot to the VM, then, once the stack is up (step 3):
docker compose cp db.sqlite3.snapshot web:/data/db.sqlite3
docker compose restart web
```

Copy the `media/` folder over the same way (`scp`/`rsync`) into the `media`
Docker volume, or mount it directly.

## 3. Bring the stack up

```bash
docker compose up -d --build
```

This starts Redis, the Django/Daphne app, and Nginx. The app container's
entrypoint (`docker-entrypoint.sh`) runs `migrate` and the two `seed_*`
permission commands automatically on every start (they're idempotent — safe
to re-run, and take a while the first time — 169 migrations). On first boot
(before you copy real data in) this creates a fresh, empty `db.sqlite3` on
the `sqlitedata` volume — that's expected, and gets overwritten by the copy
step above.

`certs/fullchain.pem` and `certs/privkey.pem` (from the sslip.io/certbot
step above, or any other real cert) must exist **before** starting `nginx`,
or that container will fail to start. Set `ALLOWED_HOSTS` and
`CSRF_TRUSTED_ORIGINS` in `.env` to match the hostname the cert was issued
for. Required for WebRTC calling to work, same reasoning as the self-signed
cert on the Windows server.

Startup race note: Nginx starts about as fast as the `web` container, but
`web` takes ~30-40s longer to finish migrations/seeding before Daphne
actually starts listening. Expect a handful of `502`s from Nginx in that
window on a fresh deploy — not a bug, it resolves itself once Daphne is up.

## 4. Verify before treating it as real

- Log in, confirm dashboards/planning/production pages load.
- Send a chat message between two accounts, confirm it arrives live.
- Start a video call between two accounts, confirm camera/mic connect.
- Compare a few record counts against the Windows server's data to confirm
  the import was complete.

## 5. Android install (PWA)

Once the site is live over HTTPS, open it in Chrome on an Android phone —
the browser will offer "Add to Home Screen" / "Install app" (comes from
`manifest.json` + `service-worker.js`, already wired into every page via
`theme/templates/theme/base.html`). No Play Store step needed for this part.

### Optional: Play Store listing (TWA)

Since a Play Console account already exists:

```bash
npm install -g @bubblewrap/cli
bubblewrap init --manifest=https://yourdomain.com/static/manifest.json
bubblewrap build
```

This produces a signed `.aab` to upload to Play Console. It's a thin wrapper
around the same live site — no separate app logic to maintain.

## 5a. Deploying code updates from GitHub to the VM

The VM's `~/offset-erp` is a real git clone of
https://github.com/EngrMahmood/Offset_Printing_ERP (`main` branch) — `.env`
and `certs/` are gitignored so they're untouched by any of this. To push a
code change (not data — see §6 for that) from GitHub to the live server:

**One-click (from the Windows PC)**: double-click
`scripts\deploy_from_github.bat`. It SSHes in and runs the deploy remotely.

**Manual**:
```bash
ssh -i ~/.ssh/offset-erp-oracle.key ubuntu@offseterp.duckdns.org
bash ~/offset-erp/scripts/deploy_update.sh
```

That script (`git pull origin main` + `docker compose up -d --build web`) is
also on the VM already. It's safe to re-run — `docker-entrypoint.sh` runs
`migrate`/seed commands idempotently on every container start. Expect a
~30-40s window of `502`s from Nginx while the new container finishes
migrating (same startup-race note as §3) if there were pending migrations.

If you don't have the SSH key or terminal access handy, ask me to run it —
same one-liner either way.

## 6. Keeping the cloud copy in sync with production

The cloud deployment is a **separate copy** of the data — nothing syncs
automatically between it and the Windows LAN server unless you push it.

**Automatic (recommended)**: `scripts\ERP_CloudSync_Task.xml` is registered
as a Windows Scheduled Task ("Offset ERP Cloud Sync") on the production PC,
running `scripts\sync_db_to_cloud.bat silent` daily at 20:20 (15 min after
the existing local auto-backup task). Check `backups\sync_to_cloud.log` for
a history of runs/results. Re-register with:
```
schtasks /Create /F /TN "Offset ERP Cloud Sync" /XML "scripts\ERP_CloudSync_Task.xml"
```

**Manual, one click**: double-click `scripts\sync_db_to_cloud.bat` anytime.
It takes a live, non-disruptive snapshot of `db.sqlite3` (SQLite's online
backup API — safe even with the production server actively in use), packs
`media/`, uploads both to the VM, and loads them into the running
deployment. Needs `%USERPROFILE%\.ssh\offset-erp-oracle.key` present (the
VM's SSH private key) and `.venv` set up.

**Manual fallback, no script/SSH command-line needed (WinSCP)**: if the
`.bat` script or SSH access is ever unavailable, you can push a database
file by hand:
1. Install **WinSCP** (free, winscp.net).
2. New Session → File protocol: **SFTP** → Host name: `offseterp.duckdns.org`
   → User name: `ubuntu` → Advanced → SSH → Authentication → Private key
   file: browse to `%USERPROFILE%\.ssh\offset-erp-oracle.key` (convert to
   `.ppk` with WinSCP's bundled PuTTYgen if it won't accept the OpenSSH
   format directly — it'll prompt you to do this automatically) → Login.
3. Get a database file to upload — either a fresh snapshot (see the Python
   one-liner in `scripts\sync_db_to_cloud.bat` step [1/5]) or a ZIP
   downloaded from the cloud site's own `/backup/` dashboard (extract
   `db.sqlite3` from it first).
4. Drag that file onto the VM's home directory (`/home/ubuntu/`) in WinSCP's
   remote pane, named `sync_db.sqlite3`.
5. **One remaining step still needs a terminal** — a file sitting on the
   VM's disk doesn't reach the running app by itself. Open WinSCP's built-in
   terminal (Commands → Open Terminal, or use PuTTY separately with the same
   key) and run:
   ```bash
   cd ~/offset-erp
   docker compose cp ~/sync_db.sqlite3 web:/data/db.sqlite3
   docker compose restart web
   rm ~/sync_db.sqlite3
   ```

## 7. Cloud backups → Google Drive / OneDrive

There's no Google Drive or OneDrive desktop client for headless Ubuntu, so
the `/backup/` dashboard's OneDrive/Google Drive folder fields support two
kinds of value now:

- **A local synced folder path** (Windows only) — copied directly, unchanged
  from before.
- **An rclone remote** (Linux/cloud servers) — e.g. `gdrive:ERP_Backups/CloudVM`
  or `onedrive:ERP_Backups/CloudVM`. Auto-detected (anything matching
  `word:` where the word is 2+ letters, so a Windows drive letter like `C:`
  never matches this) and pushed via `rclone copy`.

**One-time setup per Drive/OneDrive account** — the OAuth consent has to
happen in a real browser as the account owner; this is a step only you can
do, not something that can be automated on your behalf:
1. Install rclone locally (Windows): `winget install --id Rclone.Rclone -e`
2. Run `rclone authorize "drive"` (Google Drive) or `rclone authorize "onedrive"`
   — opens your browser to the provider's consent screen. Sign in, approve.
3. It prints a JSON token blob — hand that over so the VM's `~/.config/rclone/rclone.conf`
   can get a `[gdrive]`/`[onedrive]` remote section built from it (OneDrive
   also needs `drive_id`/`drive_type`, fetched via one Graph API call using
   the same token — no password ever passes through the VM or through me).
4. In `/backup/` → Backup Settings, set the OneDrive/Google Drive fields to
   the rclone remote spec (e.g. `gdrive:ERP_Backups/CloudVM`). The backup
   app calls `rclone copy` automatically on every backup run — no cron job
   needed.

**Container access to rclone**: the `web` container needs `rclone` installed
(done in the `Dockerfile`) and the host's rclone config bind-mounted in
(`docker-compose.yml`, `RCLONE_CONFIG_HOST_PATH`, defaults to
`/home/ubuntu/.config/rclone/rclone.conf`). On a **fresh VM**, make sure that
file exists (even empty) *before* the first `docker compose up`, so Docker
mounts a file, not an auto-created directory:
```bash
mkdir -p ~/.config/rclone && touch ~/.config/rclone/rclone.conf
```

**Media routed separately**: `/backup/` → Backup Settings also has a "Media
Backup Destination" field — set it (e.g. to a Google Drive remote) to send
media to just that one destination instead of bundling it into the database
zip that goes to both OneDrive and Google Drive. Useful when one of the two
accounts is low on space.
