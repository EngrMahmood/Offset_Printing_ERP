from django.db.models.signals import pre_save, post_save
from django.dispatch import receiver
import logging

from .models import SkuRecipe
from .views import _sync_new_jobs_for_approved_sku

logger = logging.getLogger(__name__)


@receiver(pre_save, sender=SkuRecipe)
def _sku_recipe_pre_save(sender, instance, **kwargs):
    try:
        if instance.pk:
            old = SkuRecipe.objects.filter(pk=instance.pk).values_list('master_data_status', flat=True).first()
        else:
            old = None
        instance._previous_master_data_status = old
    except Exception:
        instance._previous_master_data_status = None


@receiver(post_save, sender=SkuRecipe)
def _sku_recipe_post_save(sender, instance, created, **kwargs):
    prev = getattr(instance, '_previous_master_data_status', None)
    try:
        if (prev != 'approved') and (instance.master_data_status == 'approved'):
            # Master transitioned to approved — sync matching PO lines into Planning
            result = _sync_new_jobs_for_approved_sku(instance.sku)
            logger.info('SkuRecipe %s approved — sync result: %s', instance.sku, result)
    except Exception:
        logger.exception('Error syncing new jobs for approved SKU %s', getattr(instance, 'sku', None))
