#!/bin/sh
set -e

python manage.py migrate --noinput
python manage.py seed_access_control || true
python manage.py seed_chat_permissions || true

exec python -m daphne -b 0.0.0.0 -p 8000 Offset_ERP.asgi:application
