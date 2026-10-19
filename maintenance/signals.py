"""Tell the maintenance team when the parts a ticket is waiting for arrive.

A ticket sits in "Awaiting parts" until somebody notices the delivery. Without a
nudge that gap - parts in the store, machine still down - is invisible and counts as
downtime nobody owns. When Supply Chain records receipt of every part a ticket is
waiting for, the engineers and the assignee are notified once.
"""
import logging

from django.db.models.signals import post_save
from django.dispatch import receiver

logger = logging.getLogger(__name__)


def _all_received(record):
    lines = list(record.spare_parts.select_related('item_request__procurement'))
    linked = [line for line in lines if line.item_request_id]
    if not linked or len(linked) != len(lines):
        return False
    for line in linked:
        request = line.item_request
        procurement = getattr(request, 'procurement', None)
        received = bool(procurement and procurement.received_date) or request.status in ('RECEIVED', 'CLOSED')
        if not received:
            return False
    return True


def notify_parts_received_for(item_request):
    from .models import MaintenanceRecord
    from .services import ENGINEER_ROLES, _detail_link, log_activity
    from core.notifications import notify_roles, notify_users

    records = MaintenanceRecord.objects.filter(
        is_active=True, status='AWAITING_PARTS', spare_parts__item_request=item_request,
    ).distinct()
    for record in records:
        if record.activity_log.filter(action='Parts received').exists() and not record.activity_log.filter(
            action='Status changed', to_status='AWAITING_PARTS',
            created_at__gt=record.activity_log.filter(action='Parts received').latest('created_at').created_at,
        ).exists():
            continue                                        # already told for this wait
        if not _all_received(record):
            continue
        log_activity(record, None, 'Parts received', note='Supply Chain recorded receipt - repair can resume')
        title = f'Parts received: {record.record_no} - {record.machine}'
        message = 'The spare parts for this ticket have arrived. Resume the repair - the machine is waiting.'
        try:
            if record.assigned_to_id:
                notify_users([record.assigned_to], event_type='maintenance.parts_received', title=title, message=message,
                             link=_detail_link(record), entity_type='maintenancerecord', entity_id=record.pk)
            notify_roles(ENGINEER_ROLES, event_type='maintenance.parts_received', title=title, message=message,
                         link=_detail_link(record), entity_type='maintenancerecord', entity_id=record.pk)
        except Exception:
            logger.exception('Parts-received notification failed for record %s', record.pk)


def connect():
    from supply_chain.models import ItemProcurementTimeline, ItemRequest

    @receiver(post_save, sender=ItemProcurementTimeline, weak=False, dispatch_uid='maintenance_parts_received_timeline')
    def _timeline_saved(sender, instance, **kwargs):
        if instance.received_date:
            notify_parts_received_for(instance.request)

    @receiver(post_save, sender=ItemRequest, weak=False, dispatch_uid='maintenance_parts_received_request')
    def _request_saved(sender, instance, **kwargs):
        if instance.status in ('RECEIVED', 'CLOSED'):
            notify_parts_received_for(instance)
