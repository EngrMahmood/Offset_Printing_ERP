from django.core.exceptions import ValidationError
from django.db import transaction
from django.utils import timezone

from core.models import ChangeLog, JobCard
from planning.models import SkuRecipe
from .models import PlateRequest

PLANNING_TRIGGER_STAGES = {'new_plate_making', 'repeat_plate_making'}
PLATE_REQUEST_OPEN_STATUSES = {
    PlateRequest.STATUS_DRAFT,
    PlateRequest.STATUS_SENT,
    PlateRequest.STATUS_RECEIVED,
}
PLATE_REQUEST_COMPLETED_STATUSES = {
    PlateRequest.STATUS_AVAILABLE,
    PlateRequest.STATUS_ARCHIVED,
}
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


def create_or_get_plate_request_from_planning_job(planning_job, user):
    if planning_job.planning_stage not in PLANNING_TRIGGER_STAGES:
        return None

    existing_request = PlateRequest.objects.filter(
        planning_job=planning_job,
        status__in=PLATE_REQUEST_OPEN_STATUSES,
    ).order_by('-requested_at', '-created_at').first()

    if existing_request:
        return existing_request

    sku_recipe = None
    if getattr(planning_job, 'sku', None):
        sku_recipe = SkuRecipe.objects.filter(sku__iexact=planning_job.sku).first()

    job_card = None
    try:
        job_card = planning_job.job_card
    except (JobCard.DoesNotExist, AttributeError):
        job_card = None

    plate_request = PlateRequest.objects.create(
        planning_job=planning_job,
        job_card=job_card,
        sku_recipe=sku_recipe,
        machine=getattr(job_card, 'machine_name', None),
        department=getattr(job_card, 'department', None),
        status=PlateRequest.STATUS_DRAFT,
        requested_at=timezone.now(),
        requested_by=user,
    )

    _log_plate_request_change(
        plate_request,
        actor=user,
        action='create',
        reason='Automatically created from planning stage transition.',
        before_status='',
        after_status=plate_request.status,
    )

    return plate_request


def send_plate_request(plate_request, user, reason=''):
    return execute_plate_request_action(plate_request, 'send_plate', actor=user, reason=reason)


def receive_plate_request(plate_request, user, reason=''):
    return execute_plate_request_action(plate_request, 'receive_plate', actor=user, reason=reason)


def archive_plate_request(plate_request, user, reason=''):
    return execute_plate_request_action(plate_request, 'archive', actor=user, reason=reason)


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
    field_changes = {
        'status': {'from': before_status, 'to': after_status},
        'planning_job_id': plate_request.planning_job_id,
        'job_card_id': plate_request.job_card_id,
        'sku_recipe_id': plate_request.sku_recipe_id,
    }
    ChangeLog.objects.create(
        entity_type='plate_request',
        record_id=plate_request.pk,
        record_label=str(plate_request),
        action=action,
        changed_by=actor,
        change_reason=reason,
        field_changes=field_changes,
    )
