from __future__ import annotations

from django.conf import settings
from django.db import transaction
from django.db.models.signals import post_save
from django.dispatch import receiver


@receiver(post_save, sender='planning.SkuRecipe')
def auto_generate_bom_on_sku_save(sender, instance, created, **kwargs):
    if not getattr(settings, 'BOM_AUTO_GENERATE', True):
        return

    def _run():
        from .models import SkuBom
        from .services import generate_bom_for_sku

        has_current = SkuBom.objects.filter(sku_recipe=instance, is_current=True).exists()
        if not has_current:
            generate_bom_for_sku(instance, user=None, mode='auto')
        # If a current BOM exists, leave it alone even if the spec has drifted —
        # SkuBom.needs_regeneration surfaces that as a badge instead of
        # silently superseding approved data.

    transaction.on_commit(_run)
