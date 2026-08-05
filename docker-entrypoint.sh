#!/bin/sh
set -e

python manage.py migrate --noinput
python manage.py seed_access_control || true
python manage.py seed_chat_permissions || true
python manage.py seed_viewer_role || true

# staticfiles is a named Docker volume mounted over /app/staticfiles, so the
# image's build-time `collectstatic` output never reaches it after the first
# deploy (the volume already exists, so Docker doesn't re-seed it from the
# image). Re-run collectstatic here, against the live volume, on every
# container start so static changes actually reach production.
python manage.py collectstatic --noinput

exec python -m daphne -b 0.0.0.0 -p 8000 Offset_ERP.asgi:application
