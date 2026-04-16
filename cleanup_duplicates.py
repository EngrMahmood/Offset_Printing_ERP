#!/usr/bin/env python
"""
Cleanup script to remove duplicate UserProfile entries
Run with: python manage.py shell < cleanup_duplicates.py
"""

from core.models import UserProfile
from django.contrib.auth import get_user_model
from django.db.models import Count

User = get_user_model()

# Find users with duplicate profiles
dup_users = User.objects.annotate(profile_count=Count('userprofile')).filter(profile_count__gt=1)
print(f"Found {dup_users.count()} users with duplicate profiles")

for user in dup_users:
    profiles = UserProfile.objects.filter(user=user).order_by('id')
    # Keep the first, delete the rest
    for p in profiles[1:]:
        p.delete()
        print(f"  Deleted duplicate for {user.username}")

print("✅ Cleanup complete")
