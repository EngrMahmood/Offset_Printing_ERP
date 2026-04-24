import os
import sys
sys.path.insert(0, os.getcwd())
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'Offset_ERP.settings')
import django
try:
    django.setup()
except Exception as e:
    print('DJANGO SETUP ERROR:', e, file=sys.stderr)
    raise
from django.contrib.auth import get_user_model
User = get_user_model()
for u in User.objects.all():
    profile = getattr(u, 'profile', None)
    role = getattr(profile, 'role', 'NO_PROFILE')
    print(u.id, u.username, role)
