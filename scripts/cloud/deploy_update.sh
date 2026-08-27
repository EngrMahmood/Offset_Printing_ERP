#!/bin/bash
# Run this ON THE ORACLE VM (not locally) to pull the latest code from
# GitHub and redeploy. Safe to re-run — migrate/seed are idempotent.
set -e
cd ~/offset-erp
echo "Pulling latest from GitHub..."
git pull origin main
echo "Rebuilding and restarting the web container..."
docker compose up -d --build web
echo "Invalidating cached report payloads (deploys can change report shape without bumping their cache key)..."
docker compose exec -T web python manage.py bump_report_cache
echo "Done. Tail logs with: docker compose logs web -f"
