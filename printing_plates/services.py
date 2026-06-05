from django.core.exceptions import ValidationError
from django.db import transaction
from django.utils import timezone

from core.models import ChangeLog
from .models import PlateRequest

PLATE_REQUEST_ACTION_MAP = {
    'send_plate': PlateRequest.STATUS_SENT,
    'receive_plate': PlateRequest.STATUS_RECEIVED,
    'mark_available': PlateRequest.STATUS_AVAILABLE,
    'archive': PlateRequest.STATUS_ARCHIVED,
}

ALLOWED_PLATE_REQUEST_TRANSITIONS = {
    (PlateRequest.STATUS_DRAFT, PlateRequest.STATUS_SENT),
    (PlateRequest.STATUS_SENT, PlateRequest.STATUS_RECEIVED),
    (PlateRequest.STATUS_RECEIVED, PlateRequest.STATUS_AVAILABLE),
    (PlateRequest.STATUS_AVAILABLE, PlateRequest.STATUS_ARCHIVED),
}


def transition_plate_request_status(plate_request, target_status, actor=None, reason=''):
    current_status = plate_request.status
    if (current_status, target_status) not in ALLOWED_PLATE_REQUEST_TRANSITIONS:
        raise ValidationError({'status': f'Transition not allowed from {current_status} to {target_status}.'})

    with transaction.atomic():
        before_status = current_status
        plate_request.status = target_status
        plate_request.save(update_fields=['status'])
        _log_plate_request_change(
            plate_request,
            actor=actor,
            action=target_status,
            reason=reason,
            before_status=before_status,
            after_status=target_status,
        )

    return plate_request


def execute_plate_request_action(plate_request, action, actor=None, reason=''):
    if action not in PLATE_REQUEST_ACTION_MAP:
        raise ValueError('Unknown plate request action.')

    target_status = PLATE_REQUEST_ACTION_MAP[action]

    if action == 'send_plate':
        if actor:
            plate_request.sent_by = actor
        plate_request.sent_at = plate_request.sent_at or timezone.now()
    elif action == 'receive_plate':
        if actor:
            plate_request.received_by = actor
        plate_request.received_at = plate_request.received_at or timezone.now()
    
    save_fields = ['status']
    if action == 'send_plate':
        save_fields.extend(['sent_by', 'sent_at'])
    elif action == 'receive_plate':
        save_fields.extend(['received_by', 'received_at'])

    before_status = plate_request.status
    plate_request.status = target_status

    with transaction.atomic():
        plate_request.save(update_fields=save_fields)
        _log_plate_request_change(
            plate_request,
            actor=actor,
            action=action,
            reason=reason,
            before_status=before_status,
            after_status=target_status,
        )

    return plate_request


def _log_plate_request_change(plate_request, actor=None, action='', reason='', before_status='', after_status=''):
    ChangeLog.objects.create(
        entity_type='plate_request',
        record_id=plate_request.pk,
        record_label=str(plate_request),
        action=action,
        changed_by=actor,
        change_reason=reason,
        field_changes={'status': {'from': before_status, 'to': after_status}},
    )
