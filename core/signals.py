from django.db.models.signals import post_save
from django.dispatch import receiver
from django.contrib.auth import get_user_model
from django.db import IntegrityError

from .models import UserProfile

User = get_user_model()


@receiver(post_save, sender=User)
def create_user_profile(sender, instance, created, **kwargs):
    """Auto-create UserProfile when a new User is created"""
    if created:
        try:
            UserProfile.objects.get_or_create(user=instance)
        except IntegrityError:
            # Profile already exists, skip
            pass


@receiver(post_save, sender=User)
def ensure_user_profile_exists(sender, instance, **kwargs):
    """Ensure UserProfile exists for every user (handles edge cases)"""
    try:
        profile = instance.profile
    except UserProfile.DoesNotExist:
        try:
            UserProfile.objects.create(user=instance)
        except IntegrityError:
            # Race condition - another process created it
            pass

