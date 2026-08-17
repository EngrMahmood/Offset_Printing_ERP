import logging
from collections import namedtuple

logger = logging.getLogger(__name__)

ChangeEvent = namedtuple('ChangeEvent', ['tab_name', 'headers', 'object_pk', 'row', 'deleted'])


def _is_enabled():
    from sheets_sync.models import SheetsSyncSetting
    try:
        return SheetsSyncSetting.get_settings().enabled
    except Exception:
        # Table may not exist yet (e.g. mid-migration). Fail closed.
        return False


def enqueue_change(model_cls, instance, deleted=False):
    if not _is_enabled():
        return

    from sheets_sync.registry import get_registry_entry
    from sheets_sync.serializers import SERIALIZERS
    from sheets_sync.queue_worker import get_queue

    dotted_path = f"{model_cls._meta.app_label}.{model_cls.__name__}"
    entry = get_registry_entry(dotted_path)
    if not entry:
        return

    serializer = SERIALIZERS[entry['serializer']]
    row_dict = serializer(instance, deleted=deleted)
    row = [row_dict.get(header, '') for header in entry['headers']]

    event = ChangeEvent(
        tab_name=entry['tab_name'],
        headers=entry['headers'],
        object_pk=str(instance.pk),
        row=row,
        deleted=deleted,
    )

    queue_obj = get_queue()
    try:
        queue_obj.put_nowait(event)
    except Exception:
        logger.warning('sheets_sync: change queue is full, dropping event for %s', dotted_path)
