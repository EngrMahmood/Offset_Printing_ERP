from django.db.models.signals import post_delete, post_save
from django.dispatch import receiver

from core.models import Production

from .jc_sync import sync_issuance_for_job_card_single, sync_issuance_from_production


@receiver(post_save, sender=Production)
def sync_stock_on_production_save(sender, instance, **kwargs):
    sync_issuance_from_production(instance)


@receiver(post_delete, sender=Production)
def remove_stock_on_production_delete(sender, instance, **kwargs):
    # Rows are keyed per job card (production=None), so re-sync the JC to
    # recompute or remove its single issuance row after a run is deleted.
    job_card = getattr(instance, 'job_card', None)
    if job_card:
        sync_issuance_for_job_card_single(job_card)
