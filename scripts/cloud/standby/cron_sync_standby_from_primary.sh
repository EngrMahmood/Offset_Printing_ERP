#!/bin/bash
# Runs ON THE PRIMARY VM (offseterp.duckdns.org) via cron, twice daily, to
# refresh the standby (offseterpbackup.duckdns.org) with a live snapshot.
#
# This exists because the original standby-refresh path
# (scripts\cloud\standby\sync_standby_from_primary.bat) only runs when
# someone double-clicks it on the Windows dev PC, which isn't always on --
# the standby drifted 6 days stale before this was noticed. Running the
# refresh from the primary VM instead means it fires reliably since the VM
# is always live, with no dependency on the Windows PC being on or online.
#
# Needs ~/.ssh/offset-erp-oracle-standby on this VM (the standby's private
# key, copied over once during setup -- see DEPLOY_CLOUD.md). Never commit
# that key; it lives only on disk here and on the original Windows PC.
set -e

STANDBY_KEY="$HOME/.ssh/offset-erp-oracle-standby"
STANDBY_HOST="offseterpbackup.duckdns.org"
VM_USER="ubuntu"
REPO_DIR="$HOME/offset-erp"
LOG_FILE="$REPO_DIR/logs/standby_sync.log"

mkdir -p "$REPO_DIR/logs"

log() {
    echo "[$(date -u '+%Y-%m-%d %H:%M:%S UTC')] $1" >> "$LOG_FILE"
}

trap 'log "Standby sync FAILED - see output above this line in the log"' ERR

log "Starting standby sync (cron, from primary)"
cd "$REPO_DIR"

# Code first, then data — new code sometimes expects a migration the fresh
# data needs to run against. Without this step the standby only ever got
# data; its code silently drifted commits behind primary until someone
# happened to run update_standby_all.bat by hand (see git history for what
# that staleness already caused once: a bot email sent from the standby
# with pre-fix report logic because the standby's code was that far behind).
log "Updating standby code from GitHub"
ssh -i "$STANDBY_KEY" -o StrictHostKeyChecking=accept-new "$VM_USER@$STANDBY_HOST" "cd ~/offset-erp && git pull origin main && docker compose up -d --build web" >> "$LOG_FILE" 2>&1

# Uses SQLite's online backup API (via a python one-liner inside the web
# container) rather than a raw file copy -- the primary takes real live
# writes (WAL mode), and a raw copy can grab a half-written page mid-write.
docker compose exec -T web python -c "import sqlite3; s=sqlite3.connect('/data/db.sqlite3'); d=sqlite3.connect('/tmp/standby_pull_db.sqlite3'); s.backup(d); d.close(); s.close()" >> "$LOG_FILE" 2>&1
docker compose cp web:/tmp/standby_pull_db.sqlite3 ~/standby_pull_db.sqlite3 >> "$LOG_FILE" 2>&1
docker compose exec -T web rm -f /tmp/standby_pull_db.sqlite3 >> "$LOG_FILE" 2>&1
docker compose exec -T web tar -czf /app/standby_pull_media.tar.gz -C /app media >> "$LOG_FILE" 2>&1
docker compose cp web:/app/standby_pull_media.tar.gz ~/standby_pull_media.tar.gz >> "$LOG_FILE" 2>&1
docker compose exec -T web rm -f /app/standby_pull_media.tar.gz >> "$LOG_FILE" 2>&1

scp -i "$STANDBY_KEY" -o StrictHostKeyChecking=accept-new ~/standby_pull_db.sqlite3 "$VM_USER@$STANDBY_HOST:~/sync_db.sqlite3" >> "$LOG_FILE" 2>&1
scp -i "$STANDBY_KEY" -o StrictHostKeyChecking=accept-new ~/standby_pull_media.tar.gz "$VM_USER@$STANDBY_HOST:~/sync_media.tar.gz" >> "$LOG_FILE" 2>&1
rm -f ~/standby_pull_db.sqlite3 ~/standby_pull_media.tar.gz

scp -i "$STANDBY_KEY" -o StrictHostKeyChecking=accept-new "$REPO_DIR/scripts/cloud/remote_sync.sh" "$VM_USER@$STANDBY_HOST:~/remote_sync.sh" >> "$LOG_FILE" 2>&1
ssh -i "$STANDBY_KEY" -o StrictHostKeyChecking=accept-new "$VM_USER@$STANDBY_HOST" "bash ~/remote_sync.sh" >> "$LOG_FILE" 2>&1

log "Standby sync succeeded"
