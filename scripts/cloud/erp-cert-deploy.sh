#!/bin/sh
# certbot deploy hook for the Offset ERP: copies the renewed certificate into the folder nginx reads
# (~/offset-erp/certs). certbot runs it as root right after a successful renewal; the post-hook then
# starts nginx again so it loads the new files.
set -e
DEST=/home/ubuntu/offset-erp/certs
[ -n "$RENEWED_LINEAGE" ] || { echo "RENEWED_LINEAGE not set" >&2; exit 1; }
cp "$RENEWED_LINEAGE/fullchain.pem" "$DEST/fullchain.pem"
cp "$RENEWED_LINEAGE/privkey.pem" "$DEST/privkey.pem"
chown ubuntu:ubuntu "$DEST/fullchain.pem" "$DEST/privkey.pem"
chmod 644 "$DEST/fullchain.pem"
chmod 600 "$DEST/privkey.pem"
echo "ERP certificate copied from $RENEWED_LINEAGE"
