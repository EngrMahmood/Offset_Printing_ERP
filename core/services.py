from datetime import datetime, date, timedelta
import re
from django.db.models import Sum
from django.utils import timezone
from django.conf import settings
from django.contrib import messages
from .models import JobCard, Production, Dispatch, ChangeLog, EditOverrideRequest
from .constants import AUDIT_CONFIG

def add_unique_message(request, level, text):
    """Avoid stacking the same flash message multiple times in a single session flow."""
    storage = getattr(request, '_messages', None)
    if storage is not None:
        existing_messages = []
        existing_messages.extend(getattr(storage, '_loaded_messages', []))
        existing_messages.extend(getattr(storage, '_queued_messages', []))
        if any(getattr(message, 'level', None) == level and str(message) == text for message in existing_messages):
            return

    messages.add_message(request, level, text)


def format_audit_value(value):
    if value in (None, ''):
        return '-'
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if hasattr(value, '_meta'):
        return str(value)
    return str(value)


def normalize_colour_notation(value):
    """Convert compact notation like 1+1 into readable front/back text."""
    raw = (value or '').strip()
    if not raw:
        return None

    match = re.fullmatch(r'(\d+)\s*\+\s*(\d+)', raw)
    if not match:
        return raw

    front = int(match.group(1))
    back = int(match.group(2))
    front_label = 'color' if front == 1 else 'colors'
    back_label = 'color' if back == 1 else 'colors'
    return f"{front} {front_label} front and {back} {back_label} back"


def extract_total_colors(value):
    """Extract total color units from supported color formats."""
    raw = (value or '').strip().lower()
    if not raw:
        return 0

    simple_plus = re.fullmatch(r'(\d+)\s*\+\s*(\d+)', raw)
    if simple_plus:
        return int(simple_plus.group(1)) + int(simple_plus.group(2))

    normalized_text = re.fullmatch(r'(\d+)\s*colors?\s*front\s*and\s*(\d+)\s*colors?\s*back', raw)
    if normalized_text:
        return int(normalized_text.group(1)) + int(normalized_text.group(2))

    digits_only = re.fullmatch(r'(\d+)', raw)
    if digits_only:
        return int(digits_only.group(1))

    return 0


def compute_planned_minutes(total_impressions_required, machine, colour_value):
    """Return (run_minutes, setup_minutes, total_minutes) from machine+impressions+colors."""
    impressions = int(total_impressions_required or 0)
    if impressions <= 0 or not machine:
        return (None, None, None)

    speed = float(machine.standard_impressions_per_hour or 0)
    if speed <= 0:
        return (None, None, None)

    run_minutes = (impressions / speed) * 60
    total_colors = extract_total_colors(colour_value)
    setup_per_color = float(machine.standard_setup_minutes_per_color or 0)
    setup_minutes = total_colors * setup_per_color
    total_minutes = run_minutes + setup_minutes
    return (round(run_minutes, 2), round(setup_minutes, 2), round(total_minutes, 2))


def get_remaining_planned_minutes(job_card, exclude_production_id=None):
    """Remaining planned minutes for a job card after already allocated production entries."""
    total_planned = float(job_card.estimated_total_time_minutes or 0)
    if total_planned <= 0:
        return 0

    allocated_qs = job_card.productions.filter(is_active=True)
    if exclude_production_id:
        allocated_qs = allocated_qs.exclude(pk=exclude_production_id)
    allocated = float(allocated_qs.aggregate(total=Sum('planned_time'))['total'] or 0)
    return max(total_planned - allocated, 0)


def build_audit_snapshot(entity_type, instance):
    config = AUDIT_CONFIG[entity_type]
    return {
        field_name: format_audit_value(getattr(instance, field_name))
        for field_name in config['fields']
    }


def build_change_summary(entity_type, before_snapshot, after_snapshot):
    config = AUDIT_CONFIG[entity_type]
    summary = {}

    for field_name in config['fields']:
        before_value = before_snapshot.get(field_name, '-')
        after_value = after_snapshot.get(field_name, '-')
        if before_value == after_value:
            continue
        summary[field_name] = {
            'label': config['labels'].get(field_name, field_name.replace('_', ' ').title()),
            'from': before_value,
            'to': after_value,
        }

    return summary


def log_change(entity_type, instance, before_snapshot, changed_by, action, reason):
    config = AUDIT_CONFIG[entity_type]
    after_snapshot = build_audit_snapshot(entity_type, instance)

    if action == 'create':
        summary = {
            field_name: {
                'label': config['labels'].get(field_name, field_name.replace('_', ' ').title()),
                'from': '-',
                'to': after_value,
            }
            for field_name, after_value in after_snapshot.items()
            if after_value != '-'
        }
    elif action == 'delete':
        summary = {
            'record_state': {
                'label': 'Record State',
                'from': 'Active',
                'to': 'Archived',
            }
        }
    elif action == 'restore':
        summary = {
            'record_state': {
                'label': 'Record State',
                'from': 'Archived',
                'to': 'Active',
            }
        }
    else:
        summary = build_change_summary(entity_type, before_snapshot, after_snapshot)

    if action == 'update' and not summary:
        return False

    ChangeLog.objects.create(
        entity_type=entity_type,
        record_id=instance.pk,
        record_label=str(instance),
        action=action,
        changed_by=changed_by,
        change_reason=reason,
        field_changes=summary,
    )
    return True


def user_has_entity_permission(user, entity_type):
    config = AUDIT_CONFIG.get(entity_type)
    if not config:
        return False
    if user.is_staff:
        return True
    profile = getattr(user, 'profile', None)
    if not profile:
        return False
    return getattr(profile, config['permission'])()


def user_can_archive_records(user):
    if user.is_staff:
        return True
    profile = getattr(user, 'profile', None)
    if not profile:
        return False
    return profile.can_archive_records()


def user_can_bypass_edit_lock(user):
    return user_can_archive_records(user)


def get_record_edit_lock_days():
    try:
        return max(int(getattr(settings, 'ERP_RECORD_EDIT_LOCK_DAYS', 0)), 0)
    except (TypeError, ValueError):
        return 0


def get_record_edit_lock_cutoff():
    lock_days = get_record_edit_lock_days()
    if lock_days <= 0:
        return None
    return timezone.localdate() - timedelta(days=lock_days)


def record_is_time_locked(entity_type, record):
    date_field_map = {
        'job_card': 'po_date',
        'production': 'date',
        'dispatch': 'dispatch_date',
    }
    record_date_field = date_field_map.get(entity_type)
    cutoff = get_record_edit_lock_cutoff()
    if not record_date_field or cutoff is None:
        return False
    record_date = getattr(record, record_date_field, None)
    return bool(record_date and record_date < cutoff)


def get_valid_override(user, entity_type, record):
    """Return an approved, unexpired EditOverrideRequest for this user/record, or None."""
    return EditOverrideRequest.objects.filter(
        entity_type=entity_type,
        record_id=record.pk,
        requested_by=user,
        status='approved',
        expires_at__gt=timezone.now(),
    ).first()


def ensure_edit_lock_allowed(request, entity_type, record):
    if not record_is_time_locked(entity_type, record) or user_can_bypass_edit_lock(request.user):
        return True

    if get_valid_override(request.user, entity_type, record):
        return True

    lock_days = get_record_edit_lock_days()
    entity_label = AUDIT_CONFIG[entity_type]['model']._meta.verbose_name.title()
    add_unique_message(
        request,
        messages.ERROR,
        f'{entity_label} older than {lock_days} days is locked. Submit an override request from the records list.'
    )
    return False


def get_active_record_or_404(model, pk):
    return get_object_or_404(model, pk=pk, is_active=True)


def get_inactive_record_or_404(model, pk):
    return get_object_or_404(model, pk=pk, is_active=False)


def get_accessible_entities(user):
    entities = []
    for entity_type in ('job_card', 'production', 'dispatch'):
        if user_has_entity_permission(user, entity_type):
            entities.append(entity_type)
    return entities


def validate_delete_allowed(entity_type, record):
    if entity_type == 'job_card':
        if record.productions.filter(is_active=True).exists() or record.dispatch_set.filter(is_active=True).exists():
            raise ValueError('Job card cannot be archived while active production or dispatch records exist.')
        return

    if entity_type == 'production':
        remaining_production = sum(
            item.pcs_produced
            for item in record.job_card.productions.filter(is_active=True).exclude(pk=record.pk)
        )
        total_dispatch = record.job_card.dispatch_set.filter(is_active=True).aggregate(total=Sum('dispatch_qty'))['total'] or 0
        if total_dispatch > remaining_production:
            raise ValueError('Production record cannot be archived because active dispatch would exceed remaining production.')
        return


def validate_restore_allowed(entity_type, record):
    if entity_type == 'job_card':
        return

    if entity_type == 'production':
        if not record.job_card.is_active:
            raise ValueError('Restore the parent job card before restoring this production record.')
        return

    if entity_type == 'dispatch':
        if not record.job_card.is_active:
            raise ValueError('Restore the parent job card before restoring this dispatch record.')
        return


def archive_record(entity_type, record, user, reason):
    before_snapshot = build_audit_snapshot(entity_type, record)
    record.is_active = False
    record.save(update_fields=['is_active'])
    log_change(entity_type, record, before_snapshot, user, 'delete', reason)


def restore_record_state(entity_type, record, user, reason):
    before_snapshot = build_audit_snapshot(entity_type, record)
    record.is_active = True
    record.save(update_fields=['is_active'])
    log_change(entity_type, record, before_snapshot, user, 'restore', reason)


def run_bulk_archive(request, entity_type, record_ids):
    """Archive multiple active records with the same validation/audit pipeline as single archive."""
    config = AUDIT_CONFIG.get(entity_type)
    if not config:
        return (0, ['Unsupported entity type'])

    archived_count = 0
    failures = []
    unique_ids = []
    seen = set()
    for raw_id in record_ids:
        try:
            rid = int(raw_id)
        except (TypeError, ValueError):
            continue
        if rid in seen:
            continue
        seen.add(rid)
        unique_ids.append(rid)

    for rid in unique_ids:
        record = config['model'].objects.filter(pk=rid, is_active=True).first()
        if record is None:
            failures.append(f'#{rid}: record not found or already archived')
            continue
        try:
            validate_delete_allowed(entity_type, record)
            archive_record(entity_type, record, request.user, 'Bulk archive by admin')
            archived_count += 1
        except Exception as exc:
            failures.append(f'#{rid}: {str(exc)}')

    return (archived_count, failures)


def run_bulk_permanent_delete(request, entity_type, record_ids):
    """Permanently delete multiple active records (admin only)."""
    config = AUDIT_CONFIG.get(entity_type)
    if not config:
        return (0, ['Unsupported entity type'])

    deleted_count = 0
    failures = []
    unique_ids = []
    seen = set()
    for raw_id in record_ids:
        try:
            rid = int(raw_id)
        except (TypeError, ValueError):
            continue
        if rid in seen:
            continue
        seen.add(rid)
        unique_ids.append(rid)

    for rid in unique_ids:
        record = config['model'].objects.filter(pk=rid, is_active=True).first()
        if record is None:
            failures.append(f'#{rid}: record not found or already removed')
            continue

        try:
            validate_delete_allowed(entity_type, record)
            before_snapshot = build_audit_snapshot(entity_type, record)
            log_change(entity_type, record, before_snapshot, request.user, 'delete', 'Permanent delete by admin (bulk)')
            record.delete()
            deleted_count += 1
        except Exception as exc:
            message = str(exc)
            message = message.replace('archived', 'deleted').replace('Archive', 'Delete')
            if entity_type == 'job_card' and 'production or dispatch records exist' in message:
                message += ' Delete related production/dispatch records first, then retry job card delete.'
            failures.append(f'#{rid}: {message}')

    return (deleted_count, failures)


