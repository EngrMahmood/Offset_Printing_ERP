#!/bin/bash
# Run on the cloud VM by sync_db_to_cloud.bat — loads a freshly uploaded
# db.sqlite3 + media.tar.gz into the running Docker deployment.
set -e
cd ~/offset-erp
docker compose cp ~/sync_db.sqlite3 web:/data/db.sqlite3
docker run --rm -v offset-erp_media:/app/media -v ~/:/backup alpine \
  sh -c "rm -rf /app/media/* && tar -xzf /backup/sync_media.tar.gz -C /app"
docker compose restart web
rm -f ~/sync_db.sqlite3 ~/sync_media.tar.gz ~/remote_sync.sh
echo "remote sync complete"
