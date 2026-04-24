import os
import sys
sys.path.insert(0, os.getcwd())
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'Offset_ERP.settings')
import django
django.setup()
from django.contrib.auth import get_user_model
from core.models import UserProfile
User = get_user_model()

username = 'QC'
user = User.objects.filter(username__iexact=username).first()
if not user:
    # fallback to role-based lookup
    profile = UserProfile.objects.filter(role='qc').first()
    user = profile.user if profile else None

if not user:
    print('QC user not found')
    sys.exit(1)

profile = getattr(user, 'profile', None)
print('User:', user.id, user.username, 'is_staff=', user.is_staff, 'is_superuser=', user.is_superuser)
print('Profile role:', getattr(profile, 'role', None))

perm_methods = [
    'can_edit_jobcard',
    'can_edit_production',
    'can_approve_dispatch',
    'can_view_analytics',
    'can_manage_masters',
    'can_approve_qc',
    'can_manage_operators',
    'can_archive_records',
    'can_view_reports',
]

for pm in perm_methods:
    val = None
    try:
        val = getattr(profile, pm)()
    except Exception as e:
        val = f'ERROR: {e}'
    print(f'{pm}: {val}')

# Also print which URLs the user would be allowed based on permission checks in core.views
from django.urls import reverse
print('\nQuick navigation hints:')
print('Manage user roles URL:', reverse('manage_user_roles'))
print('Planning home URL:', reverse('planning:planning_home'))
print('SKU recipes URL:', reverse('planning:sku_recipes'))
