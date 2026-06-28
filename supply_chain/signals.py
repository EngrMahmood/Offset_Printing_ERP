from django.db.models.signals import post_delete, post_save
from django.dispatch import receiver

from core.models import Production

from .jc_sync import sync_issuance_from_production


@receiver(post_save, sender=Production)
def sync_stock_on_production_save(sender, instance, **kwargs):
    sync_issuance_from_production(instance)


@receiver(post_delete, sender=Production)
def remove_stock_on_production_delete(sender, instance, **kwargs):
    if hasattr(instance, 'stock_issuance'):
        instance.stock_issuance.delete()
