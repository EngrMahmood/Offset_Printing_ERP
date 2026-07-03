#!/usr/bin/env bash
# Deploy Offset ERP to the development server (Linux/macOS).
# Backup the database first, then run:
#   ./deploy.sh
#
# Options are passed through to scripts/deploy_dev.py, for example:
#   ./deploy.sh --git-pull
#   ./deploy.sh --run-tests

set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$ROOT"

if [[ -n "${VIRTUAL_ENV:-}" && -x "${VIRTUAL_ENV}/bin/python" ]]; then
  PYTHON="${VIRTUAL_ENV}/bin/python"
else
  PYTHON="python3"
fi

exec "$PYTHON" scripts/deploy_dev.py --confirm-backup "$@"
