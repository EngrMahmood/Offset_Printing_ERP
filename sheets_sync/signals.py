import logging

from django.apps import apps
from django.db.models.signals import post_save, post_delete

logger = logging.getLogger(__name__)

_connected = False


def _on_save(sender, instance, created, **kwargs):
    try:
        from sheets_sync.services import enqueue_change
        enqueue_change(sender, instance, deleted=False)
    except Exception:
        logger.exception('sheets_sync: failed to enqueue change for %s', sender)


def _on_delete(sender, instance, **kwargs):
    try:
        from sheets_sync.services import enqueue_change
        enqueue_change(sender, instance, deleted=True)
    except Exception:
        logger.exception('sheets_sync: failed to enqueue delete for %s', sender)


def connect_all():
    """Connect the generic save/delete receivers for every registered model.

    Safe to call multiple times (e.g. under the dev autoreloader) — Django's
    dispatch_uid de-dupes repeat connect() calls for the same signal+sender.
    """
    global _connected
    from sheets_sync.registry import SYNCED_MODELS

    for entry in SYNCED_MODELS:
        app_label, model_name = entry['dotted_path'].split('.')
        try:
            model = apps.get_model(app_label, model_name)
        except LookupError:
            logger.warning('sheets_sync: model %s not found, skipping', entry['dotted_path'])
            continue

        post_save.connect(
            _on_save, sender=model, weak=False,
            dispatch_uid=f"sheets_sync_save_{entry['dotted_path']}",
        )
        post_delete.connect(
            _on_delete, sender=model, weak=False,
            dispatch_uid=f"sheets_sync_delete_{entry['dotted_path']}",
        )

    _connected = True
