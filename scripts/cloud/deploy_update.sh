#!/bin/bash
# Run this ON THE ORACLE VM (not locally) to pull the latest code from
# GitHub and redeploy. Safe to re-run — migrate/seed are idempotent.
set -e
cd ~/offset-erp
echo "Pulling latest from GitHub..."
git pull origin main
echo "Rebuilding and restarting the web container..."
docker compose up -d --build web
echo "Done. Tail logs with: docker compose logs web -f"
