from django.core.exceptions import ValidationError
from django.db import transaction
from django.db.models import Q
from django.utils import timezone

from core.models import ChangeLog, JobCard, Vendor
from planning.models import SkuRecipe
from planning.services import ensure_sku_recipe_for_planning_job
from .models import PlateRequest


def get_plate_vendor_names():
    """Active master vendors plus any vendor names already used on plate requests."""
    canonical = {}
    for name in Vendor.objects.filter(is_active=True).order_by('name').values_list('name', flat=True):
        value = (name or '').strip()
        if value:
            canonical[value.lower()] = value
    for name in PlateRequest.objects.exclude(vendor='').values_list('vendor', flat=True).distinct():
        value = (name or '').strip()
        if value and value.lower() not in canonical:
            canonical[value.lower()] = value
    return sorted(canonical.values(), key=lambda item: item.lower())


def build_vendor_filter_options(queryset):
    """Return [{'name': ..., 'count': n}, ...] for sidebar/dashboard vendor filters."""
    options = []
    for name in get_plate_vendor_names():
        options.append({
            'name': name,
            'count': queryset.filter(vendor__iexact=name).count(),
        })
    return options

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

REPLACEMENT_SOURCES = {
    PlateRequest.SOURCE_REPLACEMENT,
    PlateRequest.SOURCE_PRODUCTION_PLATE_DAMAGE,
}

VALID_REPLACEMENT_REASONS = {choice[0] for choice in PlateRequest.REPLACEMENT_REASON_CHOICES}


def create_or_get_plate_request_from_planning_job(planning_job, user):
    if planning_job.planning_stage not in PLANNING_TRIGGER_STAGES:
        return None

    existing_request = PlateRequest.objects.filter(
        planning_job=planning_job,
        status__in=PLATE_REQUEST_OPEN_STATUSES,
    ).order_by('-requested_at', '-created_at').first()

    if existing_request:
        return existing_request

    sku_recipe = ensure_sku_recipe_for_planning_job(planning_job, actor=user)
    if sku_recipe:
        from planning.services import sync_planning_job_fields_to_sku_recipe
        sync_planning_job_fields_to_sku_recipe(planning_job, sku_recipe)

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
        source=PlateRequest.SOURCE_PLANNING,
        requested_at=timezone.now(),
        requested_by=user,
        remarks=getattr(planning_job, 'remarks', '') or '',
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


def job_card_plate_requests_qs(job_card):
    planning_job = getattr(job_card, 'planning_job', None)
    filters = Q(job_card=job_card)
    if planning_job:
        filters |= Q(planning_job=planning_job)
    return PlateRequest.objects.filter(filters)


def get_issued_plate_request(job_card):
    """Latest plate set that was issued (available) or previously issued (archived)."""
    return (
        job_card_plate_requests_qs(job_card)
        .filter(status__in={PlateRequest.STATUS_AVAILABLE, PlateRequest.STATUS_ARCHIVED})
        .order_by('-updated_at', '-created_at')
        .first()
    )


def plates_were_issued_to_production(job_card):
    """
    True when plates are (or were) on the floor.
    Supports legacy jobs that never had an available PlateRequest row.
    """
    if get_issued_plate_request(job_card):
        return True

    if job_card_plate_requests_qs(job_card).filter(status=PlateRequest.STATUS_AVAILABLE).exists():
        return True

    planning_job = getattr(job_card, 'planning_job', None)
    if planning_job:
        stage = (planning_job.planning_stage or '').strip()
        if stage in {'plate_received', 'in_production', 'planning_done'}:
            return True
        if (planning_job.plate_set_no or '').strip():
            return True

    if (job_card.plate_set_no or '').strip():
        return True

    if job_card.productions.filter(is_active=True, entry_type='printing').exists():
        return True

    return False


def get_open_replacement_requests(job_card):
    return (
        job_card_plate_requests_qs(job_card)
        .filter(
            status__in=PLATE_REQUEST_OPEN_STATUSES,
        )
        .filter(
            Q(source__in=REPLACEMENT_SOURCES) | ~Q(replacement_reason='')
        )
        .order_by('-requested_at', '-created_at')
    )


def job_is_waiting_for_plates(job_card):
    return get_open_replacement_requests(job_card).exists()


def get_plate_remake_count(job_card):
    return (
        job_card_plate_requests_qs(job_card)
        .filter(Q(source__in=REPLACEMENT_SOURCES) | ~Q(replacement_reason=''))
        .count()
    )


def get_job_card_awc_no(job_card):
    planning_job = getattr(job_card, 'planning_job', None)
    if planning_job:
        awc = (planning_job.awc_no_display or '').strip()
        if awc:
            return awc
    issued = get_issued_plate_request(job_card)
    if issued and (issued.awc_no or '').strip():
        return issued.awc_no.strip()
    latest = job_card_plate_requests_qs(job_card).exclude(awc_no='').order_by('-updated_at').first()
    if latest:
        return (latest.awc_no or '').strip()
    return ''


def get_job_card_plate_set_no(job_card):
    if (job_card.plate_set_no or '').strip():
        return job_card.plate_set_no.strip()
    planning_job = getattr(job_card, 'planning_job', None)
    if planning_job and (planning_job.plate_set_no or '').strip():
        return planning_job.plate_set_no.strip()
    issued = get_issued_plate_request(job_card)
    if issued:
        return (issued.set_no or issued.new_set_no or '').strip()
    return ''


def validate_plate_remake_request(job_card, reason, notes='', damaged_colors=''):
    reason = (reason or '').strip()
    notes = (notes or '').strip()
    damaged_colors = (damaged_colors or '').strip()

    if not reason or reason not in VALID_REPLACEMENT_REASONS:
        raise ValidationError('Select a valid plate replacement reason.')

    if not damaged_colors:
        raise ValidationError('Select at least one damaged colour.')

    if reason == PlateRequest.REASON_OTHER and not notes:
        raise ValidationError('Notes are required when reason is Other.')

    if getattr(job_card, 'is_print_job', True) is False:
        raise ValidationError('Plate requests are only for print jobs.')

    issued = get_issued_plate_request(job_card)
    if plates_were_issued_to_production(job_card):
        return issued

    open_non_replacement = (
        job_card_plate_requests_qs(job_card)
        .filter(status__in=PLATE_REQUEST_OPEN_STATUSES)
        .exclude(Q(source__in=REPLACEMENT_SOURCES) | ~Q(replacement_reason=''))
        .exists()
    )
    if open_non_replacement:
        raise ValidationError(
            'Plates are not issued to production yet. Graphics is still processing the current plate request.'
        )
    raise ValidationError(
        'Plates have not been issued to production yet. Wait until graphics issues plates, then request a replacement.'
    )


def request_plate_remake(job_card, actor=None, reason='', damaged_colors='', notes='', source=None):
    """
    Open a plate replacement request from Released Jobs (or on behalf of production).
    Keeps prior plate sets visible (archived, not deleted). Multiple open requests allowed.
    """
    from workflow.services import _append_unique_note_line

    reason = (reason or '').strip()
    damaged_colors = (damaged_colors or '').strip()
    notes = (notes or '').strip()
    source = source or PlateRequest.SOURCE_REPLACEMENT

    if not damaged_colors:
        raise ValidationError('Select at least one damaged colour.')

    issued = validate_plate_remake_request(
        job_card,
        reason,
        notes=notes,
        damaged_colors=damaged_colors,
    )
    planning_job = getattr(job_card, 'planning_job', None)

    with transaction.atomic():
        # Keep prior issued sets visible in history as archived (not deleted).
        prior_qs = job_card_plate_requests_qs(job_card).filter(status=PlateRequest.STATUS_AVAILABLE)
        prior_qs.update(status=PlateRequest.STATUS_ARCHIVED)

        if planning_job:
            planning_job.planning_stage = 'repeat_plate_making'
            planning_job.planning_stage_changed_at = timezone.now()
            planning_job.planning_stage_changed_by = actor
            update_fields = [
                'planning_stage',
                'planning_stage_changed_at',
                'planning_stage_changed_by',
                'updated_at',
            ]
            label = dict(PlateRequest.REPLACEMENT_REASON_CHOICES).get(reason, reason)
            planning_job.remarks = _append_unique_note_line(
                planning_job.remarks,
                f'Plate replacement requested ({label}): {notes or "—"}',
            )
            update_fields.append('remarks')
            planning_job.save(update_fields=update_fields)

        sku_recipe = None
        if planning_job:
            sku_recipe = ensure_sku_recipe_for_planning_job(planning_job, actor=actor)
        elif (job_card.SKU or '').strip():
            sku_recipe = SkuRecipe.objects.filter(sku__iexact=job_card.SKU.strip()).first()

        remarks_parts = []
        label = dict(PlateRequest.REPLACEMENT_REASON_CHOICES).get(reason, reason)
        remarks_parts.append(f'Reason: {label}')
        remarks_parts.append(f'Damaged colours: {damaged_colors}')
        if notes:
            remarks_parts.append(notes)

        prior_set_no = ''
        prior_awc = ''
        if issued:
            prior_set_no = (issued.set_no or issued.new_set_no or '').strip()
            prior_awc = (issued.awc_no or '').strip()
        prior_set_no = prior_set_no or get_job_card_plate_set_no(job_card)
        prior_awc = prior_awc or get_job_card_awc_no(job_card)

        plate_request = PlateRequest.objects.create(
            planning_job=planning_job,
            job_card=job_card,
            sku_recipe=sku_recipe,
            machine=getattr(job_card, 'machine_name', None),
            department=getattr(job_card, 'department', None),
            status=PlateRequest.STATUS_DRAFT,
            source=source,
            replacement_reason=reason,
            damaged_colors=damaged_colors,
            plate_color=damaged_colors,
            replaces_request=issued,
            set_no=prior_set_no,
            awc_no=prior_awc,
            remarks='\n'.join(remarks_parts),
            requested_at=timezone.now(),
            requested_by=actor,
        )

        _log_plate_request_change(
            plate_request,
            actor=actor,
            action='create_replacement',
            reason=notes or label,
            before_status='',
            after_status=plate_request.status,
        )

    try:
        from core.notifications import notify_plate_replacement
        notify_plate_replacement(plate_request, actor=actor)
    except Exception:
        pass

    return plate_request


def request_replacement_plates_for_production(job_card, actor=None, reason=''):
    """Backward-compatible wrapper used by older entry points."""
    notes = (reason or '').strip() or 'Plate damaged during production.'
    return request_plate_remake(
        job_card,
        actor=actor,
        reason=PlateRequest.REASON_DAMAGED_DURING_RUN,
        notes=notes,
        source=PlateRequest.SOURCE_REPLACEMENT,
    )


def send_plate_request(plate_request, user, reason=''):
    return execute_plate_request_action(plate_request, 'send_plate', actor=user, reason=reason)


def receive_plate_request(plate_request, user, reason=''):
    return execute_plate_request_action(plate_request, 'receive_plate', actor=user, reason=reason)


def archive_plate_request(plate_request, user, reason=''):
    return execute_plate_request_action(plate_request, 'archive', actor=user, reason=reason)


def cancel_plate_request(plate_request, actor=None, reason=''):
    """
    Cancel an open plate request when plates are not required.
    Archives the request and clears waiting-for-plate on the planning job when possible.
    """
    from workflow.services import _append_unique_note_line

    reason = (reason or '').strip()
    if not reason:
        raise ValidationError('A cancel reason is required.')

    if plate_request.status not in PLATE_REQUEST_OPEN_STATUSES:
        raise ValidationError('Only open plate requests (draft, sent, or received) can be cancelled.')

    with transaction.atomic():
        before_status = plate_request.status
        note = f'Cancelled — plates not required: {reason}'
        plate_request.status = PlateRequest.STATUS_ARCHIVED
        plate_request.remarks = _append_unique_note_line(plate_request.remarks, note)
        plate_request.progress = 'Cancelled'
        plate_request.save(update_fields=['status', 'remarks', 'progress', 'updated_at'])

        planning_job = plate_request.planning_job
        if planning_job and planning_job.planning_stage in PLANNING_TRIGGER_STAGES:
            # If any plates were previously issued, job can return to plate_received.
            has_issued = PlateRequest.objects.filter(
                planning_job=planning_job,
                status__in={PlateRequest.STATUS_AVAILABLE, PlateRequest.STATUS_ARCHIVED},
            ).exclude(pk=plate_request.pk).exists()
            planning_job.planning_stage = 'plate_received' if has_issued else 'jc_ready'
            planning_job.planning_stage_changed_at = timezone.now()
            planning_job.planning_stage_changed_by = actor
            planning_job.remarks = _append_unique_note_line(planning_job.remarks, note)
            planning_job.save(update_fields=[
                'planning_stage',
                'planning_stage_changed_at',
                'planning_stage_changed_by',
                'remarks',
                'updated_at',
            ])

        _log_plate_request_change(
            plate_request,
            actor=actor,
            action='cancel_request',
            reason=reason,
            before_status=before_status,
            after_status=plate_request.status,
        )

    try:
        from core.notifications import notify_plate_cancelled
        notify_plate_cancelled(plate_request, actor=actor)
    except Exception:
        pass

    return plate_request


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
        'source': plate_request.source,
        'replacement_reason': plate_request.replacement_reason,
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
