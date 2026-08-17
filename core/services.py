from datetime import datetime, date, timedelta
import re
from django.db.models import Q, Sum
from django.shortcuts import get_object_or_404
from django.utils import timezone
from django.conf import settings
from django.contrib import messages
from .models import Department, DeliveryLocation, ProductType, JobCard, Production, Dispatch, ChangeLog, EditOverrideRequest
from .constants import AUDIT_CONFIG

# SQLite's compiled-in bound-parameter ceiling (SQLITE_MAX_VARIABLE_NUMBER)
# is commonly 999 on older builds and 32766 on modern ones — a single
# `field__in=values` filter with more values than that raises "too many SQL
# variables". 900 stays safely under either, with headroom for whatever
# other params share the same query.
SQL_IN_CHUNK_SIZE = 900


def chunked(values, size=SQL_IN_CHUNK_SIZE):
    """Yield `values` (any iterable) in fixed-size list chunks.

    Use this to split a large `field__in=values` filter into several safe
    queries instead of one that can exceed SQLite's bound-parameter limit —
    see SQL_IN_CHUNK_SIZE. Unlike a `Q(...) |= Q(...)` OR-chain per value,
    this has no expression-tree-depth ceiling either.
    """
    values = list(values)
    for i in range(0, len(values), size):
        yield values[i:i + size]


def find_completed_job_card_matches(query, limit=5):
    """
    Job cards matching a search query that are already Completed.

    Entry-form job card search boxes only ever return jobs still eligible
    for new entries, so a completed job simply looks "not found" — which
    reads as broken/missing rather than "already finished". Callers use
    this to add a clear "already Completed" note instead of a bare empty
    result when that's the actual reason nothing matched.
    """
    query = (query or '').strip()
    if not query:
        return []

    qs = JobCard.objects.filter(is_active=True, status='completed').filter(
        Q(job_card_no__icontains=query)
        | Q(SKU__icontains=query)
        | Q(PO_No__icontains=query)
        | Q(destination__icontains=query)
    ).order_by('-updated_at')[:limit]

    return [
        {'job_card_no': jc.job_card_no, 'sku': jc.SKU or '-'}
        for jc in qs
    ]


def compute_job_card_wastage_metrics(job_card):
    """
    Per-job wastage figures, matching the Wastage Report (reports/services.py
    build_wastage_report_context) definition exactly: total wastage is not
    just printing/sorting waste — it also includes the dispatch gap (planned
    qty not yet dispatched), and status is Tentative until the job is
    Completed/Closed (waste can still change until then).
    """
    if job_card is None:
        return None

    ups = job_card.ups or 1
    plan_qty_pcs = int(job_card.total_sheets_planned * ups)

    dispatch_qty_pcs = sum(d.dispatch_qty for d in job_card.dispatch_set.filter(is_active=True))
    printing_waste_sheets = sum(
        p.waste_sheets for p in job_card.productions.filter(is_active=True, entry_type='printing')
    )
    printing_waste_pcs = printing_waste_sheets * ups
    sorting_waste_pcs = sum(
        p.sorting_waste_qty for p in job_card.productions.filter(is_active=True, entry_type='packing')
    )
    dispatch_gap_pcs = max(plan_qty_pcs - dispatch_qty_pcs, 0)

    is_completed = job_card.status in ('completed', 'closed') or job_card.job_status == 'Completed'
    wastage_status = 'Finalized' if is_completed else 'Tentative'

    total_wastage_pcs = printing_waste_pcs + sorting_waste_pcs + dispatch_gap_pcs
    total_wastage_pct = round((total_wastage_pcs / plan_qty_pcs * 100), 2) if plan_qty_pcs > 0 else 0.0

    # The dispatch-gap component is exactly the part of total wastage that
    # process wastage (printing + sorting) did NOT explain — it's what a
    # wrong/understated waste entry during production shows up as once the
    # real, final dispatch count comes in. Flag it once the job is finalized
    # and the gap exceeds the job's own planned production tolerance, so a
    # supervisor can go back and check that job's entries rather than the
    # gap just sitting unexplained in the report.
    dispatch_gap_pct = round((dispatch_gap_pcs / plan_qty_pcs * 100), 2) if plan_qty_pcs > 0 else 0.0
    tolerance_pct = float(job_card.production_tolerance_percent or 5)
    needs_reconciliation_review = wastage_status == 'Finalized' and dispatch_gap_pct > tolerance_pct

    return {
        'plan_qty_pcs': plan_qty_pcs,
        'dispatch_qty_pcs': dispatch_qty_pcs,
        'printing_waste_sheets': printing_waste_sheets,
        'printing_waste_pcs': printing_waste_pcs,
        'sorting_waste_pcs': sorting_waste_pcs,
        'dispatch_gap_pcs': dispatch_gap_pcs,
        'wastage_status': wastage_status,
        'total_wastage_pcs': total_wastage_pcs,
        'total_wastage_pct': total_wastage_pct,
        'needs_reconciliation_review': needs_reconciliation_review,
    }


def collect_planning_department_names():
    """Return distinct non-empty department names used in planning jobs and PO documents."""
    from planning.models import PlanningJob, PoDocument

    names = set()
    for department in PlanningJob.objects.exclude(department='').values_list('department', flat=True).distinct():
        cleaned = (department or '').strip()
        if cleaned:
            names.add(cleaned)

    for payload in PoDocument.objects.exclude(extracted_payload__isnull=True).values_list('extracted_payload', flat=True):
        if not isinstance(payload, dict):
            continue
        cleaned = (payload.get('department') or '').strip()
        if cleaned:
            names.add(cleaned)

    return names


def sync_departments_from_planning():
    """Create Department master records for planning department names that are not in master data yet."""
    created = 0
    for name in sorted(collect_planning_department_names()):
        if Department.objects.filter(name__iexact=name).exists():
            continue
        Department.objects.create(name=name)
        created += 1
    return created


def collect_planning_delivery_location_names():
    """Return distinct non-empty delivery location names used in planning jobs and PO documents."""
    from planning.models import PlanningJob, PoDocument

    names = set()
    for destination in PlanningJob.objects.exclude(destination='').values_list('destination', flat=True).distinct():
        cleaned = (destination or '').strip()
        if cleaned:
            names.add(cleaned)

    for payload in PoDocument.objects.exclude(extracted_payload__isnull=True).values_list('extracted_payload', flat=True):
        if not isinstance(payload, dict):
            continue
        cleaned = (payload.get('delivery_location') or '').strip()
        if cleaned:
            names.add(cleaned)

    return names


def sync_delivery_locations_from_planning():
    """Create DeliveryLocation master records for planning delivery names not yet in master data."""
    created = 0
    for name in sorted(collect_planning_delivery_location_names()):
        if DeliveryLocation.objects.filter(name__iexact=name).exists():
            continue
        DeliveryLocation.objects.create(name=name)
        created += 1
    return created


def collect_sku_recipe_product_type_names():
    """Return distinct non-empty product type names used in SKU master recipes."""
    from planning.models import SkuRecipe

    names = set()
    for product_type in SkuRecipe.objects.exclude(product_type='').values_list('product_type', flat=True).distinct():
        cleaned = (product_type or '').strip()
        if cleaned:
            names.add(cleaned)
    return names


def sync_product_types_from_sku_recipes():
    """Create ProductType master records for SKU recipe product types not yet in master data."""
    created = 0
    for name in sorted(collect_sku_recipe_product_type_names()):
        if ProductType.objects.filter(name__iexact=name).exists():
            continue
        ProductType.objects.create(name=name)
        created += 1
    return created


def collect_planning_material_names():
    """Return distinct non-empty material names used in planning jobs and SKU recipes."""
    from planning.models import PlanningJob, SkuRecipe

    names = set()
    for material in PlanningJob.objects.exclude(material='').values_list('material', flat=True).distinct():
        cleaned = (material or '').strip()
        if cleaned:
            names.add(cleaned)

    for material in SkuRecipe.objects.exclude(material='').values_list('material', flat=True).distinct():
        cleaned = (material or '').strip()
        if cleaned:
            names.add(cleaned)

    return names


def sync_materials_from_planning():
    """Create Material master records for planning material names not yet in master data."""
    from .models import Material
    created = 0
    for name in sorted(collect_planning_material_names()):
        if Material.objects.filter(name__iexact=name).exists():
            continue
        Material.objects.create(name=name)
        created += 1
    return created


def count_active_planning_jobs_for_product_type(product_type_name):
    """Count active planning jobs whose SKU master recipe uses the given product type."""
    from django.db.models import Exists, OuterRef
    from planning.models import PlanningJob, SkuRecipe

    return PlanningJob.objects.filter(is_active=True).filter(
        Exists(
            SkuRecipe.objects.filter(
                product_type__iexact=product_type_name,
                sku__iexact=OuterRef('sku'),
            )
        )
    ).count()

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
    if entity_type in ('production', 'dispatch'):
        from production.wip_service import evaluate_and_update_job_wip_status
        evaluate_and_update_job_wip_status(record.job_card, user=user)


def restore_record_state(entity_type, record, user, reason):
    before_snapshot = build_audit_snapshot(entity_type, record)
    record.is_active = True
    record.save(update_fields=['is_active'])
    log_change(entity_type, record, before_snapshot, user, 'restore', reason)
    if entity_type in ('production', 'dispatch'):
        from production.wip_service import evaluate_and_update_job_wip_status
        evaluate_and_update_job_wip_status(record.job_card, user=user)


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
            job_card = record.job_card if entity_type in ('production', 'dispatch') else None
            record.delete()
            if job_card:
                from production.wip_service import evaluate_and_update_job_wip_status
                evaluate_and_update_job_wip_status(job_card, user=request.user)
            deleted_count += 1
        except Exception as exc:
            message = str(exc)
            message = message.replace('archived', 'deleted').replace('Archive', 'Delete')
            if entity_type == 'job_card' and 'production or dispatch records exist' in message:
                message += ' Delete related production/dispatch records first, then retry job card delete.'
            failures.append(f'#{rid}: {message}')

    return (deleted_count, failures)


