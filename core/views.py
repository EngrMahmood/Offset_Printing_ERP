from django.shortcuts import render, get_object_or_404, redirect
from django.conf import settings
from django.http import HttpResponse, JsonResponse, Http404
from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.utils import timezone
from django.views.decorators.http import require_POST
from django.db.models import Q, Sum, Min, Max
from django.db import transaction
from django.db.models.deletion import ProtectedError
from django.db.models.functions import Coalesce
from django.core.paginator import Paginator
from django.urls import reverse
from functools import wraps
import csv
import json
import re
from datetime import datetime, date, timedelta

from .bulk_upload import process_jobcard_upload, get_template_headers, get_template_example
from .jc_numbering import allocate_next_jc_number
from .models import JobCard, Production, ProductionDowntime, Machine, Operator, Department, Material, Dispatch, UserProfile, ChangeLog, EditOverrideRequest, ShiftConfig, MachineWorkSchedule

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False


# ═══════════════════════════════════
# PERMISSION DECORATORS (RBAC)
# ═══════════════════════════════════

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

def require_role(*allowed_roles):
    """Decorator to check if user has required role"""
    def decorator(view_func):
        @wraps(view_func)
        def wrapped_view(request, *args, **kwargs):
            if not request.user.is_authenticated:
                return redirect('login')
            try:
                profile = request.user.profile
                if profile.role not in allowed_roles and not request.user.is_staff:
                    add_unique_message(request, messages.ERROR, '❌ You do not have permission to access this page.')
                    return redirect('home')
            except UserProfile.DoesNotExist:
                add_unique_message(request, messages.ERROR, '⚠️ Your user profile is not configured. Contact admin.')
                return redirect('login')
            return view_func(request, *args, **kwargs)
        return wrapped_view
    return decorator


def permission_required(permission_method):
    """Decorator to check specific permission method on UserProfile"""
    def decorator(view_func):
        @wraps(view_func)
        def wrapped_view(request, *args, **kwargs):
            if not request.user.is_authenticated:
                return redirect('login')
            try:
                profile = request.user.profile
                if not getattr(profile, permission_method)():
                    add_unique_message(request, messages.ERROR, '❌ You do not have permission to access this feature.')
                    return redirect('home')
            except (UserProfile.DoesNotExist, AttributeError):
                add_unique_message(request, messages.ERROR, '⚠️ Permission check failed. Contact admin.')
                return redirect('login')
            return view_func(request, *args, **kwargs)
        return wrapped_view
    return decorator


AUDIT_CONFIG = {
    'job_card': {
        'model': JobCard,
        'permission': 'can_edit_jobcard',
        'list_view': 'job_card_records',
        'fields': [
            'job_card_no', 'SKU', 'PO_No', 'po_date', 'month', 'material', 'colour', 'application',
            'order_qty', 'total_impressions_required', 'ups', 'wastage', 'print_sheet_size',
            'purchase_sheet_size', 'purchase_sheet_ups', 'machine_name', 'department', 'destination',
            'die_cutting', 'status', 'remarks',
        ],
        'labels': {
            'job_card_no': 'Job Card No',
            'SKU': 'SKU',
            'PO_No': 'PO Number',
            'po_date': 'PO Date',
            'month': 'Month',
            'material': 'Material',
            'colour': 'Colours',
            'application': 'Application',
            'order_qty': 'Order Qty',
            'total_impressions_required': 'Total Impressions Required',
            'ups': 'UPS',
            'wastage': 'Wastage',
            'print_sheet_size': 'Print Sheet Size',
            'purchase_sheet_size': 'Purchase Sheet Size',
            'purchase_sheet_ups': 'Purchase Sheet UPS',
            'machine_name': 'Machine',
            'department': 'Department',
            'destination': 'Destination',
            'die_cutting': 'Die Cutting',
            'status': 'Status',
            'remarks': 'Remarks',
        },
    },
    'production': {
        'model': Production,
        'permission': 'can_edit_production',
        'list_view': 'production_records',
        'fields': [
            'job_card', 'machine', 'operator', 'shift', 'date', 'impressions', 'output_sheets',
            'waste_sheets', 'planned_time', 'run_time', 'setup_time', 'downtime', 'downtime_category',
            'downtime_breakdown_text',
            'waste_reason',
        ],
        'labels': {
            'job_card': 'Job Card',
            'machine': 'Machine',
            'operator': 'Operator',
            'shift': 'Shift',
            'date': 'Production Date',
            'impressions': 'Impressions',
            'output_sheets': 'Output Sheets',
            'waste_sheets': 'Waste Sheets',
            'planned_time': 'Planned Time',
            'run_time': 'Run Time',
            'setup_time': 'Setup Time',
            'downtime': 'Downtime',
            'downtime_category': 'Downtime Category',
            'downtime_breakdown_text': 'Downtime Breakdown',
            'waste_reason': 'Waste Reason',
        },
    },
    'dispatch': {
        'model': Dispatch,
        'permission': 'can_approve_dispatch',
        'list_view': 'dispatch_records',
        'fields': ['job_card', 'dc_no', 'dispatch_date', 'dispatch_qty'],
        'labels': {
            'job_card': 'Job Card',
            'dc_no': 'DC No',
            'dispatch_date': 'Dispatch Date',
            'dispatch_qty': 'Dispatch Qty',
        },
    },
}


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


@login_required
def home(request):
    return render(request, 'home.html')


def erp_version(request):
    return JsonResponse({
        'erp_software_version': getattr(settings, 'ERP_SOFTWARE_VERSION', '0.0.0'),
        'erp_software_release_date': getattr(settings, 'ERP_SOFTWARE_RELEASE_DATE', ''),
        'server_time': timezone.now().isoformat(),
    })


@login_required
@permission_required('can_edit_jobcard')
def bulk_upload_jobcards(request):
    context = {}

    if request.method == "POST":
        file = request.FILES.get('file')

        if not file:
            context = {
                'success_count': 0,
                'error_count': 1,
                'errors': [{'row': 0, 'errors': 'Please choose a file to upload.'}],
            }
            return render(request, "upload.html", context)

        result = process_jobcard_upload(file, uploaded_by=request.user)
        context = result

        return render(request, "upload.html", context)

    return render(request, "upload.html", context)


@login_required
@require_POST
def quick_add_master(request):
    """Create master dropdown values for planner workflow without admin dependency."""
    master_type = (request.POST.get('type') or '').strip().lower()
    name = (request.POST.get('name') or '').strip()

    if master_type not in {'material', 'machine', 'department', 'operator'}:
        return JsonResponse({'ok': False, 'error': 'Invalid master type.'}, status=400)

    if not name:
        return JsonResponse({'ok': False, 'error': 'Name is required.'}, status=400)

    if master_type == 'operator':
        employee_code = (request.POST.get('employee_code') or '').strip() or None

        existing = None
        if employee_code:
            existing = Operator.objects.filter(employee_code__iexact=employee_code).first()
        if existing is None:
            existing = Operator.objects.filter(
                name__iexact=name,
                employee_code__iexact=employee_code or ''
            ).first() if employee_code else Operator.objects.filter(name__iexact=name, employee_code__isnull=True).first()

        if existing:
            display_name = f"{existing.name} ({existing.employee_code})" if existing.employee_code else existing.name
            return JsonResponse({
                'ok': True,
                'created': False,
                'id': existing.id,
                'name': existing.name,
                'display_name': display_name,
                'employee_code': existing.employee_code,
                'type': master_type,
                'message': 'Already exists. Selected existing value.'
            })

        obj = Operator.objects.create(name=name, employee_code=employee_code)
        display_name = f"{obj.name} ({obj.employee_code})" if obj.employee_code else obj.name
        return JsonResponse({
            'ok': True,
            'created': True,
            'id': obj.id,
            'name': obj.name,
            'display_name': display_name,
            'employee_code': obj.employee_code,
            'type': master_type,
            'message': 'Created successfully.'
        })

    if master_type == 'machine':
        standard_speed_raw = (request.POST.get('standard_impressions_per_hour') or '').strip()
        setup_per_color_raw = (request.POST.get('standard_setup_minutes_per_color') or '').strip()
        standard_speed = 4000
        setup_per_color = 15
        if standard_speed_raw:
            try:
                standard_speed = float(standard_speed_raw)
            except ValueError:
                return JsonResponse({'ok': False, 'error': 'Ideal speed must be a number.'}, status=400)
            if standard_speed <= 0:
                return JsonResponse({'ok': False, 'error': 'Ideal speed must be greater than 0.'}, status=400)

        if setup_per_color_raw:
            try:
                setup_per_color = float(setup_per_color_raw)
            except ValueError:
                return JsonResponse({'ok': False, 'error': 'Setup minutes per color must be a number.'}, status=400)
            if setup_per_color < 0:
                return JsonResponse({'ok': False, 'error': 'Setup minutes per color cannot be negative.'}, status=400)

        existing = Machine.objects.filter(name__iexact=name).first()
        if existing:
            return JsonResponse({
                'ok': True,
                'created': False,
                'id': existing.id,
                'name': existing.name,
                'standard_impressions_per_hour': existing.standard_impressions_per_hour,
                'standard_setup_minutes_per_color': existing.standard_setup_minutes_per_color,
                'type': master_type,
                'message': 'Already exists. Selected existing value.'
            })

        obj = Machine.objects.create(
            name=name,
            standard_impressions_per_hour=standard_speed,
            standard_setup_minutes_per_color=setup_per_color,
        )
        return JsonResponse({
            'ok': True,
            'created': True,
            'id': obj.id,
            'name': obj.name,
            'standard_impressions_per_hour': obj.standard_impressions_per_hour,
            'standard_setup_minutes_per_color': obj.standard_setup_minutes_per_color,
            'type': master_type,
            'message': 'Created successfully.'
        })

    model_map = {
        'material': Material,
        'department': Department,
    }
    model = model_map[master_type]

    existing = model.objects.filter(name__iexact=name).first()
    if existing:
        return JsonResponse({
            'ok': True,
            'created': False,
            'id': existing.id,
            'name': existing.name,
            'type': master_type,
            'message': 'Already exists. Selected existing value.'
        })

    obj = model.objects.create(name=name)
    return JsonResponse({
        'ok': True,
        'created': True,
        'id': obj.id,
        'name': obj.name,
        'type': master_type,
        'message': 'Created successfully.'
    })


@login_required
@permission_required('can_edit_jobcard')
def job_card_entry(request):
    """Manual job card entry form"""
    view_id = (request.GET.get('view') or '').strip()
    is_view_mode = bool(view_id)
    edit_id = '' if is_view_mode else (request.POST.get('edit_id') or request.GET.get('edit') or '').strip()
    edit_record = get_active_record_or_404(JobCard, view_id) if is_view_mode else (get_active_record_or_404(JobCard, edit_id) if edit_id else None)

    if edit_record and not is_view_mode and not ensure_edit_lock_allowed(request, 'job_card', edit_record):
        return redirect('job_card_records')

    if request.method == "POST" and not is_view_mode:
        try:
            change_reason = (request.POST.get('change_reason') or '').strip()
            sku = (request.POST.get('sku') or '').strip()
            order_qty = int(request.POST.get('order_qty') or 0)
            production_tolerance_percent = float(request.POST.get('production_tolerance_percent') or 5)

            po_date_raw = (request.POST.get('po_date') or '').strip()
            if not po_date_raw:
                raise ValueError("PO Date is required")
            po_date = datetime.strptime(po_date_raw, "%Y-%m-%d").date()
            month_value = po_date.strftime("%B")

            # Enforce system-generated immutable JC numbering.
            if edit_record:
                job_card_no = edit_record.job_card_no
            else:
                job_card_no = allocate_next_jc_number(po_date)

            if edit_record and not change_reason:
                raise ValueError("Change reason is required when editing a job card")
            if not job_card_no:
                raise ValueError("Job card number is required")
            if not sku:
                raise ValueError("SKU is required")
            if order_qty <= 0:
                raise ValueError("Order quantity must be greater than 0")
            if production_tolerance_percent < 0:
                raise ValueError("Production tolerance cannot be negative")

            duplicate_query = JobCard.objects.filter(job_card_no=job_card_no)
            if edit_record:
                duplicate_query = duplicate_query.exclude(pk=edit_record.pk)
            if duplicate_query.exists():
                raise ValueError(f"Job card number {job_card_no} already exists")

            material = Material.objects.filter(id=request.POST.get('material')).first() if request.POST.get('material') else None
            machine = Machine.objects.filter(id=request.POST.get('machine_name')).first() if request.POST.get('machine_name') else None
            department = Department.objects.filter(id=request.POST.get('department')).first() if request.POST.get('department') else None
            normalized_colour = normalize_colour_notation(request.POST.get('colour'))
            total_impressions_required = int(request.POST.get('total_impressions_required') or 0) or None
            estimated_run_minutes, estimated_setup_minutes, estimated_total_minutes = compute_planned_minutes(
                total_impressions_required,
                machine,
                normalized_colour,
            )

            payload = {
                'job_card_no': job_card_no,
                'month': month_value,
                'po_date': po_date,
                'PO_No': (request.POST.get('po_no') or '').strip() or None,
                'SKU': sku,
                'material': material,
                'colour': normalized_colour,
                'application': (request.POST.get('application') or '').strip() or None,
                'order_qty': order_qty,
                'production_tolerance_percent': production_tolerance_percent,
                'total_impressions_required': total_impressions_required,
                'estimated_run_time_minutes': estimated_run_minutes,
                'estimated_setup_time_minutes': estimated_setup_minutes,
                'estimated_total_time_minutes': estimated_total_minutes,
                'ups': int(request.POST.get('ups') or 0) or None,
                'print_sheet_size': (request.POST.get('print_sheet_size') or '').strip() or None,
                'wastage': int(request.POST.get('wastage') or 0),
                'purchase_sheet_size': (request.POST.get('purchase_sheet_size') or '').strip() or None,
                'purchase_sheet_ups': int(request.POST.get('purchase_sheet_ups') or 0) or None,
                'remarks': (request.POST.get('remarks') or '').strip() or None,
                'destination': (request.POST.get('destination') or '').strip() or None,
                'machine_name': machine,
                'department': department,
                'die_cutting': (request.POST.get('die_cutting') or '').strip() or None,
                'status': (request.POST.get('status') or 'Open').strip() or 'Open',
                'is_print_job': request.POST.get('is_print_job') == 'true',
            }

            if edit_record:
                before_snapshot = build_audit_snapshot('job_card', edit_record)
                for field_name, value in payload.items():
                    setattr(edit_record, field_name, value)
                edit_record.save()

                if log_change('job_card', edit_record, before_snapshot, request.user, 'update', change_reason):
                    messages.success(request, f"Job card {edit_record.job_card_no} updated successfully")
                else:
                    messages.success(request, f"No changes detected for job card {edit_record.job_card_no}")
                return redirect('job_card_records')

            job_card = JobCard.objects.create(**payload)
            job_card.created_by = request.user
            job_card.save(update_fields=['created_by'])
            log_change('job_card', job_card, {}, request.user, 'create', 'Initial entry created')

            messages.success(request, f"Job card {job_card_no} created successfully")
            return redirect('job_card_entry')

        except Exception as e:
            messages.error(request, f"Error creating job card: {str(e)}")

    machine_options = Machine.objects.filter(is_active=True).order_by('name', 'id')

    context = {
        'today': edit_record.po_date if edit_record and edit_record.po_date else timezone.now().date(),
        'materials': Material.objects.all().order_by('name'),
        'machines': machine_options,
        'departments': Department.objects.all().order_by('name'),
        'edit_record': edit_record,
        'is_view_mode': is_view_mode,
        'machine_meta_json': json.dumps({
            str(m.id): {
                'speed': float(m.standard_impressions_per_hour or 0),
                'setup_per_color': float(m.standard_setup_minutes_per_color or 0),
            }
            for m in machine_options
        }),
    }
    return render(request, 'job_card_entry.html', context)


@login_required
@permission_required('can_edit_jobcard')
def job_card_records(request):
    """Job card records list page"""
    if request.method == 'POST':
        action = (request.POST.get('action') or '').strip()
        if action == 'close_short_close':
            if not request.user.profile.can_approve_dispatch():
                add_unique_message(request, messages.ERROR, '❌ Only manager/dispatch roles can close short-close gap.')
                return redirect('job_card_records')

            job_card_id = (request.POST.get('job_card_id') or '').strip()
            reason = (request.POST.get('short_close_reason') or '').strip()

            if not job_card_id:
                add_unique_message(request, messages.ERROR, 'Job card is required for short-close action.')
                return redirect('job_card_records')
            if not reason:
                add_unique_message(request, messages.ERROR, 'Reason is required to close pending short-close.')
                return redirect('job_card_records')

            row = get_active_record_or_404(JobCard, job_card_id)
            pending_qty = row.short_close_qty
            if pending_qty <= 0:
                add_unique_message(request, messages.INFO, 'No pending short-close quantity to close.')
                return redirect('job_card_records')

            before_snapshot = build_audit_snapshot('job_card', row)
            row.short_close_closed_qty = (row.short_close_closed_qty or 0) + pending_qty
            row.short_close_wastage_qty = (row.short_close_wastage_qty or 0) + pending_qty
            row.short_close_closed_by = request.user
            row.short_close_closed_at = timezone.now()
            row.short_close_close_reason = reason
            row.save(update_fields=[
                'short_close_closed_qty',
                'short_close_wastage_qty',
                'short_close_closed_by',
                'short_close_closed_at',
                'short_close_close_reason',
            ])
            log_change(
                'job_card',
                row,
                before_snapshot,
                request.user,
                'update',
                f'Short-close closed: {pending_qty} pcs moved to wastage. Reason: {reason}'
            )
            add_unique_message(request, messages.SUCCESS, f'{pending_qty} pcs short-close moved to wastage for {row.job_card_no}.')
            return redirect('job_card_records')
        elif action == 'bulk_delete':
            if request.user.profile.role != 'admin':
                add_unique_message(request, messages.ERROR, '❌ Only admin can run bulk delete.')
                return redirect('job_card_records')
            if not user_can_archive_records(request.user):
                add_unique_message(request, messages.ERROR, '❌ You do not have permission to delete records.')
                return redirect('job_card_records')

            selected_ids = request.POST.getlist('selected_ids')
            deleted_count, failures = run_bulk_permanent_delete(request, 'job_card', selected_ids)
            if deleted_count:
                add_unique_message(request, messages.SUCCESS, f'Deleted {deleted_count} job card record(s) permanently.')
            if failures:
                add_unique_message(request, messages.ERROR, f'Bulk delete completed with issues: {"; ".join(failures[:5])}')
            return redirect('job_card_records')

    query = (request.GET.get('q') or '').strip()
    status = (request.GET.get('status') or '').strip()
    date_from_raw = (request.GET.get('date_from') or '').strip()
    date_to_raw = (request.GET.get('date_to') or '').strip()
    entry_date_from_raw = (request.GET.get('entry_date_from') or '').strip()
    entry_date_to_raw = (request.GET.get('entry_date_to') or '').strip()
    sort = (request.GET.get('sort') or 'created_at').strip()
    direction = (request.GET.get('dir') or 'desc').strip().lower()
    per_page = request.GET.get('per_page') or '50'
    try:
        per_page = int(per_page)
    except (TypeError, ValueError):
        per_page = 50
    if per_page not in (50, 100):
        per_page = 50

    jobcards = JobCard.objects.filter(is_active=True).select_related('material', 'machine_name', 'department', 'created_by').annotate(
        total_dispatch_agg=Coalesce(Sum('dispatch__dispatch_qty', filter=Q(dispatch__is_active=True)), 0),
        total_production_agg=Coalesce(Sum('productions__output_sheets', filter=Q(productions__is_active=True)), 0),
    ).order_by('-created_at')

    if query:
        jobcards = jobcards.filter(
            Q(job_card_no__icontains=query) |
            Q(SKU__icontains=query) |
            Q(PO_No__icontains=query)
        )

    if status:
        jobcards = jobcards.filter(status=status)

    date_from = None
    date_to = None
    if date_from_raw:
        try:
            date_from = datetime.strptime(date_from_raw, '%Y-%m-%d').date()
            jobcards = jobcards.filter(po_date__gte=date_from)
        except ValueError:
            add_unique_message(request, messages.ERROR, 'Invalid From date format. Use YYYY-MM-DD.')
            date_from_raw = ''
    if date_to_raw:
        try:
            date_to = datetime.strptime(date_to_raw, '%Y-%m-%d').date()
            jobcards = jobcards.filter(po_date__lte=date_to)
        except ValueError:
            add_unique_message(request, messages.ERROR, 'Invalid To date format. Use YYYY-MM-DD.')
            date_to_raw = ''

    entry_date_from = None
    entry_date_to = None
    if entry_date_from_raw:
        try:
            entry_date_from = datetime.strptime(entry_date_from_raw, '%Y-%m-%d').date()
            jobcards = jobcards.filter(created_at__date__gte=entry_date_from)
        except ValueError:
            add_unique_message(request, messages.ERROR, 'Invalid Entry Date From format. Use YYYY-MM-DD.')
            entry_date_from_raw = ''
    if entry_date_to_raw:
        try:
            entry_date_to = datetime.strptime(entry_date_to_raw, '%Y-%m-%d').date()
            jobcards = jobcards.filter(created_at__date__lte=entry_date_to)
        except ValueError:
            add_unique_message(request, messages.ERROR, 'Invalid Entry Date To format. Use YYYY-MM-DD.')
            entry_date_to_raw = ''

    if date_from and date_to and date_from > date_to:
        add_unique_message(request, messages.ERROR, 'From date cannot be later than To date.')
        jobcards = jobcards.none()

    sortable_fields = {
        'job_card_no': 'job_card_no',
        'sku': 'SKU',
        'po_no': 'PO_No',
        'po_date': 'po_date',
        'order_qty': 'order_qty',
        'dispatch': 'total_dispatch_agg',
        'status': 'status',
        'added_by': 'created_by__username',
        'created_at': 'created_at',
    }
    order_field = sortable_fields.get(sort, 'created_at')
    if direction not in ('asc', 'desc'):
        direction = 'desc'
    ordering = order_field if direction == 'asc' else f'-{order_field}'
    jobcards = jobcards.order_by(ordering)

    total_count = jobcards.count()
    paginator = Paginator(jobcards, per_page)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)

    jobcards = list(page_obj.object_list)
    for row in jobcards:
        total_dispatch = int(row.total_dispatch_agg or 0)
        total_production = int(row.total_production_agg or 0)
        row.total_dispatch_display = total_dispatch
        row.dispatch_completion_percent_display = round((total_dispatch / row.order_qty) * 100, 2) if row.order_qty else 0
        row.balance_qty_display = (row.order_qty or 0) - total_dispatch

        if row.order_qty == 0:
            row.job_status_display = 'Open'
        elif row.dispatch_completion_percent_display >= 95:
            row.job_status_display = 'Completed'
        elif total_production > 0:
            row.job_status_display = 'In Progress'
        else:
            row.job_status_display = 'Open'

        if row.job_status_display == 'Completed' and total_dispatch < (row.order_qty or 0):
            gap = (row.order_qty or 0) - total_dispatch
            row.short_close_qty_display = max(gap - int(row.short_close_closed_qty or 0), 0)
        else:
            row.short_close_qty_display = 0

    cutoff = get_record_edit_lock_cutoff()
    pending_ids: set = set()
    approved_ids: set = set()
    if cutoff and not user_can_bypass_edit_lock(request.user):
        user_overrides = EditOverrideRequest.objects.filter(
            entity_type='job_card',
            requested_by=request.user,
        ).values('record_id', 'status', 'expires_at')
        for ov in user_overrides:
            if ov['status'] == 'pending':
                pending_ids.add(ov['record_id'])
            elif ov['status'] == 'approved' and ov['expires_at'] and ov['expires_at'] > timezone.now():
                approved_ids.add(ov['record_id'])

    context = {
        'jobcards': jobcards,
        'page_obj': page_obj,
        'total_count': total_count,
        'q': query,
        'status': status,
        'date_from': date_from_raw,
        'date_to': date_to_raw,
        'entry_date_from': entry_date_from_raw,
        'entry_date_to': entry_date_to_raw,
        'sort': sort,
        'dir': direction,
        'per_page': per_page,
        'edit_lock_days': get_record_edit_lock_days(),
        'edit_lock_cutoff': cutoff,
        'can_bypass_edit_lock': user_can_bypass_edit_lock(request.user),
        'pending_override_ids': pending_ids,
        'approved_override_ids': approved_ids,
    }
    return render(request, 'job_card_records.html', context)


@login_required
@permission_required('can_edit_production')
def production_entry(request):
    """Production data entry form for operators"""
    view_id = (request.GET.get('view') or '').strip()
    is_view_mode = bool(view_id)
    edit_id = '' if is_view_mode else (request.POST.get('edit_id') or request.GET.get('edit') or '').strip()
    edit_record = get_active_record_or_404(Production, view_id) if is_view_mode else (get_active_record_or_404(Production, edit_id) if edit_id else None)
    if edit_record and not is_view_mode and not ensure_edit_lock_allowed(request, 'production', edit_record):
        return redirect('production_records')

    if request.method == "POST" and not is_view_mode:
        job_card_id = request.POST.get('job_card')
        machine_id = request.POST.get('machine')
        machine_override = request.POST.get('machine_override') == 'on'
        operator_id = request.POST.get('operator')
        shift = request.POST.get('shift')
        date = request.POST.get('date')
        impressions = request.POST.get('impressions')
        output_sheets = request.POST.get('output_sheets')
        waste_sheets = request.POST.get('waste_sheets')
        intermediate_pass = request.POST.get('intermediate_pass') == 'on'
        downtime_minutes = request.POST.get('downtime_minutes')
        planned_time = request.POST.get('planned_time')
        run_time = request.POST.get('run_time')
        setup_time = request.POST.get('setup_time')
        downtime_category = request.POST.get('downtime_category')
        downtime_categories = request.POST.getlist('downtime_category[]')
        downtime_minutes_rows = request.POST.getlist('downtime_minutes_detail[]')
        downtime_notes = request.POST.getlist('downtime_note[]')
        waste_reason = request.POST.get('waste_reason')
        remarks = request.POST.get('remarks')
        overrun_reason_select = (request.POST.get('overrun_reason_select') or '').strip()
        overrun_reason_other = (request.POST.get('overrun_reason_other') or '').strip()

        try:
            change_reason = (request.POST.get('change_reason') or '').strip()
            output_sheets_val = int(output_sheets) if output_sheets else 0
            waste_sheets_val = int(waste_sheets) if waste_sheets else 0
            legacy_downtime_val = float(downtime_minutes) if downtime_minutes else 0
            planned_time_val = float(planned_time) if planned_time else 0
            run_time_val = float(run_time) if run_time else 0
            setup_time_val = float(setup_time) if setup_time else 0
            impressions_val = int(impressions) if impressions else 0

            downtime_entries = []
            max_rows = max(len(downtime_categories), len(downtime_minutes_rows), len(downtime_notes))
            for idx in range(max_rows):
                category = (downtime_categories[idx] if idx < len(downtime_categories) else '').strip()
                minute_raw = (downtime_minutes_rows[idx] if idx < len(downtime_minutes_rows) else '').strip()
                note = (downtime_notes[idx] if idx < len(downtime_notes) else '').strip()

                if not category and not minute_raw and not note:
                    continue

                if not category:
                    raise ValueError("Downtime category is required for each downtime row")

                try:
                    minutes = float(minute_raw or 0)
                except ValueError:
                    raise ValueError("Downtime minutes must be numeric")

                if minutes <= 0:
                    raise ValueError("Downtime minutes must be greater than 0 for each downtime row")

                downtime_entries.append({
                    'category': category,
                    'minutes': minutes,
                    'note': note or None,
                })

            downtime_val = float(sum(item['minutes'] for item in downtime_entries)) if downtime_entries else legacy_downtime_val
            primary_downtime_category = downtime_entries[0]['category'] if downtime_entries else downtime_category

            if intermediate_pass:
                output_sheets_val = 0

            if edit_record and not change_reason:
                raise ValueError("Change reason is required when editing production data")
            if planned_time_val < 0:
                raise ValueError("Planned time cannot be negative")
            if run_time_val <= 0:
                raise ValueError("Run time must be greater than 0")
            if downtime_val > 0 and not primary_downtime_category:
                raise ValueError("Downtime category is required when downtime is greater than 0")
            if waste_sheets_val > 0 and not waste_reason:
                raise ValueError("Waste reason is required when waste sheets are greater than 0")
            if intermediate_pass and impressions_val <= 0:
                raise ValueError("Impressions must be greater than 0 for intermediate pass entry")
            if intermediate_pass and waste_sheets_val < 0:
                raise ValueError("Waste sheets cannot be negative")

            job_card = get_active_record_or_404(JobCard, job_card_id)
            if machine_override and machine_id:
                machine = get_object_or_404(Machine, pk=machine_id)
            elif job_card.machine_name_id:
                machine = job_card.machine_name
            elif machine_id:
                machine = get_object_or_404(Machine, pk=machine_id)
            elif edit_record and edit_record.machine_id:
                machine = edit_record.machine
            else:
                raise ValueError("No machine mapped on selected Job Card. Please set machine in Job Card (or choose fallback machine).")
            operator = get_object_or_404(Operator, pk=operator_id)

            remaining_planned = get_remaining_planned_minutes(
                job_card,
                exclude_production_id=edit_record.pk if edit_record else None,
            )

            if job_card.estimated_total_time_minutes and job_card.estimated_total_time_minutes > 0:
                if edit_record:
                    # Keep current allocation while editing the existing row.
                    planned_time_val = float(edit_record.planned_time or 0)
                else:
                    # New rows consume remaining planned time. If nothing remains, allow manual overrun allocation.
                    if remaining_planned > 0:
                        planned_time_val = remaining_planned
                    elif planned_time_val <= 0:
                        planned_time_val = 0

            overrun_reason_map = {
                'extra_setup': 'Extra Setup / Make Ready',
                'machine_slowdown': 'Machine Slowdown',
                'operator_learning': 'Operator Learning / New Team',
                'material_issue': 'Material-related Delay',
                'quality_rework': 'Quality Rework',
                'job_complexity': 'Higher Job Complexity',
                'other': 'Other',
            }
            overrun_reason = ''
            if overrun_reason_select == 'other':
                overrun_reason = overrun_reason_other
            elif overrun_reason_select in overrun_reason_map:
                overrun_reason = overrun_reason_map[overrun_reason_select]

            if not edit_record and planned_time_val > remaining_planned:
                if not overrun_reason_select:
                    raise ValueError(
                        "Overrun allocation reason is required when planned minutes exceed remaining planned time."
                    )
                if overrun_reason_select == 'other' and not overrun_reason_other:
                    raise ValueError("Please specify overrun reason details for 'Other'.")

            payload = {
                'job_card': job_card,
                'machine': machine,
                'operator': operator,
                'shift': shift,
                'date': date,
                'impressions': impressions_val,
                'output_sheets': output_sheets_val,
                'waste_sheets': waste_sheets_val,
                'planned_time': planned_time_val,
                'run_time': run_time_val,
                'setup_time': setup_time_val,
                'downtime': downtime_val,
                'downtime_category': primary_downtime_category,
                'waste_reason': waste_reason,
            }

            if edit_record:
                before_snapshot = build_audit_snapshot('production', edit_record)
                with transaction.atomic():
                    for field_name, value in payload.items():
                        setattr(edit_record, field_name, value)
                    edit_record.save()
                    edit_record.downtime_entries.all().delete()
                    if downtime_entries:
                        ProductionDowntime.objects.bulk_create([
                            ProductionDowntime(
                                production=edit_record,
                                category=item['category'],
                                minutes=item['minutes'],
                                note=item['note'],
                            )
                            for item in downtime_entries
                        ])

                if log_change('production', edit_record, before_snapshot, request.user, 'update', change_reason):
                    messages.success(request, f'Production record updated for Job Card {job_card.job_card_no}')
                else:
                    messages.success(request, f'No changes detected for Job Card {job_card.job_card_no}')
                return redirect('production_records')

            with transaction.atomic():
                record = Production.objects.create(**payload)
                if downtime_entries:
                    ProductionDowntime.objects.bulk_create([
                        ProductionDowntime(
                            production=record,
                            category=item['category'],
                            minutes=item['minutes'],
                            note=item['note'],
                        )
                        for item in downtime_entries
                    ])
                record.created_by = request.user
                record.save(update_fields=['created_by'])
            create_reason = overrun_reason or 'Initial entry created'
            log_change('production', record, {}, request.user, 'create', create_reason)

            messages.success(request, f'Production data saved successfully for Job Card {job_card.job_card_no}')
            return redirect('production_entry')

        except Exception as e:
            messages.error(request, f'Error saving production data: {str(e)}')

    # Get data for form dropdowns
    job_cards = JobCard.objects.filter(is_active=True, status__in=['Open', 'In Progress']).order_by('-created_at')
    if edit_record:
        job_cards = JobCard.objects.filter(is_active=True).filter(Q(status__in=['Open', 'In Progress']) | Q(pk=edit_record.job_card_id)).distinct().order_by('-created_at')
    machines = Machine.objects.filter(is_active=True)
    operators = Operator.objects.all()

    context = {
        'job_cards': job_cards,
        'machines': machines,
        'operators': operators,
        'job_card_plan_json': json.dumps({
            str(j.id): {
                'planned_total': float(j.estimated_total_time_minutes or 0),
                'planned_setup': float(j.estimated_setup_time_minutes or 0),
                'planned_run': float(j.estimated_run_time_minutes or 0),
                'remaining_planned': float(get_remaining_planned_minutes(j, exclude_production_id=edit_record.pk if edit_record else None)),
            }
            for j in job_cards
        }),
        'job_card_machine_json': json.dumps({
            str(j.id): {
                'machine_id': j.machine_name_id,
                'machine_name': j.machine_name.name if j.machine_name else '',
            }
            for j in job_cards
        }),
        'today': edit_record.date if edit_record else timezone.now().date(),
        'edit_record': edit_record,
        'edit_downtime_rows_json': json.dumps([
            {
                'category': row.category,
                'minutes': float(row.minutes or 0),
                'note': row.note or '',
            }
            for row in (edit_record.downtime_entries.all() if edit_record else [])
        ]),
        'edit_lock_days': get_record_edit_lock_days(),
        'edit_lock_applies': bool(edit_record and record_is_time_locked('production', edit_record)),
        'is_view_mode': is_view_mode,
    }

    return render(request, 'production_entry.html', context)


@login_required
@permission_required('can_edit_production')
def production_records(request):
    """Production records list page"""
    if request.method == 'POST':
        action = (request.POST.get('action') or '').strip()
        if action == 'bulk_delete':
            if request.user.profile.role != 'admin':
                add_unique_message(request, messages.ERROR, '❌ Only admin can run bulk delete.')
                return redirect('production_records')
            if not user_can_archive_records(request.user):
                add_unique_message(request, messages.ERROR, '❌ You do not have permission to delete records.')
                return redirect('production_records')

            selected_ids = request.POST.getlist('selected_ids')
            deleted_count, failures = run_bulk_permanent_delete(request, 'production', selected_ids)
            if deleted_count:
                add_unique_message(request, messages.SUCCESS, f'Deleted {deleted_count} production record(s) permanently.')
            if failures:
                add_unique_message(request, messages.ERROR, f'Bulk delete completed with issues: {"; ".join(failures[:5])}')
            return redirect('production_records')

    query = (request.GET.get('q') or '').strip()
    shift = (request.GET.get('shift') or '').strip()
    date_from_raw = (request.GET.get('date_from') or '').strip()
    date_to_raw = (request.GET.get('date_to') or '').strip()
    sort = (request.GET.get('sort') or 'date').strip()
    direction = (request.GET.get('dir') or 'desc').strip().lower()
    per_page = request.GET.get('per_page') or '50'
    try:
        per_page = int(per_page)
    except (TypeError, ValueError):
        per_page = 50
    if per_page not in (50, 100):
        per_page = 50

    records = Production.objects.filter(is_active=True, job_card__is_active=True).select_related('job_card', 'machine', 'operator', 'created_by').order_by('-date', '-id')

    if query:
        records = records.filter(
            Q(job_card__job_card_no__icontains=query) |
            Q(machine__name__icontains=query) |
            Q(operator__name__icontains=query)
        )

    if shift:
        records = records.filter(shift=shift)

    date_from = None
    date_to = None
    if date_from_raw:
        try:
            date_from = datetime.strptime(date_from_raw, '%Y-%m-%d').date()
            records = records.filter(date__gte=date_from)
        except ValueError:
            add_unique_message(request, messages.ERROR, 'Invalid From date format. Use YYYY-MM-DD.')
            date_from_raw = ''
    if date_to_raw:
        try:
            date_to = datetime.strptime(date_to_raw, '%Y-%m-%d').date()
            records = records.filter(date__lte=date_to)
        except ValueError:
            add_unique_message(request, messages.ERROR, 'Invalid To date format. Use YYYY-MM-DD.')
            date_to_raw = ''

    if date_from and date_to and date_from > date_to:
        add_unique_message(request, messages.ERROR, 'From date cannot be later than To date.')
        records = records.none()

    sortable_fields = {
        'date': 'date',
        'job_card': 'job_card__job_card_no',
        'machine': 'machine__name',
        'operator': 'operator__name',
        'shift': 'shift',
        'impressions': 'impressions',
        'output': 'output_sheets',
        'waste': 'waste_sheets',
        'planned': 'planned_time',
        'added_by': 'created_by__username',
    }
    order_field = sortable_fields.get(sort, 'date')
    if direction not in ('asc', 'desc'):
        direction = 'desc'
    ordering = order_field if direction == 'asc' else f'-{order_field}'
    records = records.order_by(ordering)

    total_count = records.count()
    paginator = Paginator(records, per_page)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)

    records = list(page_obj.object_list)
    job_card_ids = list({row.job_card_id for row in records})
    consumption_map = {
        item['job_card_id']: int((item['total_output'] or 0) + (item['total_waste'] or 0))
        for item in Production.objects.filter(is_active=True, job_card_id__in=job_card_ids).values('job_card_id').annotate(
            total_output=Sum('output_sheets'),
            total_waste=Sum('waste_sheets'),
        )
    }
    for row in records:
        consumed = consumption_map.get(row.job_card_id, 0)
        row.job_card_extra_sheets_used = max(consumed - row.job_card.total_sheets_planned, 0)
        row.job_card_tolerance_sheets = row.job_card.tolerance_sheets

    cutoff = get_record_edit_lock_cutoff()
    pending_ids: set = set()
    approved_ids: set = set()
    if cutoff and not user_can_bypass_edit_lock(request.user):
        user_overrides = EditOverrideRequest.objects.filter(
            entity_type='job_card',
            requested_by=request.user,
        ).values('record_id', 'status', 'expires_at')
        for ov in user_overrides:
            if ov['status'] == 'pending':
                pending_ids.add(ov['record_id'])
            elif ov['status'] == 'approved' and ov['expires_at'] and ov['expires_at'] > timezone.now():
                approved_ids.add(ov['record_id'])

    context = {
        'edit_lock_days': get_record_edit_lock_days(),
        'edit_lock_cutoff': cutoff,
        'can_bypass_edit_lock': user_can_bypass_edit_lock(request.user),
        'pending_override_ids': pending_ids,
        'approved_override_ids': approved_ids,
        'records': records,
        'page_obj': page_obj,
        'total_count': total_count,
        'q': query,
        'shift': shift,
        'date_from': date_from_raw,
        'date_to': date_to_raw,
        'sort': sort,
        'dir': direction,
        'per_page': per_page,
        'edit_lock_days': get_record_edit_lock_days(),
        'edit_lock_cutoff': cutoff,
        'can_bypass_edit_lock': user_can_bypass_edit_lock(request.user),
        'pending_override_ids': pending_ids,
        'approved_override_ids': approved_ids,
    }
    return render(request, 'production_records.html', context)


@login_required
@permission_required('can_approve_dispatch')
def dispatch_entry(request):
    """Dispatch entry form"""
    view_id = (request.GET.get('view') or '').strip()
    is_view_mode = bool(view_id)
    edit_id = '' if is_view_mode else (request.POST.get('edit_id') or request.GET.get('edit') or '').strip()
    edit_record = get_active_record_or_404(Dispatch, view_id) if is_view_mode else (get_active_record_or_404(Dispatch, edit_id) if edit_id else None)
    if edit_record and not is_view_mode and not ensure_edit_lock_allowed(request, 'dispatch', edit_record):
        return redirect('dispatch_records')

    if request.method == 'POST' and not is_view_mode:
        try:
            change_reason = (request.POST.get('change_reason') or '').strip()
            job_card_id = request.POST.get('job_card')
            dc_no = (request.POST.get('dc_no') or '').strip() or None
            dispatch_date_raw = request.POST.get('dispatch_date')
            dispatch_qty = int(request.POST.get('dispatch_qty') or 0)

            if edit_record and not change_reason:
                raise ValueError('Change reason is required when editing dispatch')
            if not job_card_id:
                raise ValueError('Job card is required')
            if dispatch_qty <= 0:
                raise ValueError('Dispatch quantity must be greater than 0')

            job_card = get_active_record_or_404(JobCard, job_card_id)
            dispatch_date = datetime.strptime(dispatch_date_raw, "%Y-%m-%d").date() if dispatch_date_raw else timezone.now().date()

            payload = {
                'job_card': job_card,
                'dc_no': dc_no,
                'dispatch_date': dispatch_date,
                'dispatch_qty': dispatch_qty,
            }

            if edit_record:
                before_snapshot = build_audit_snapshot('dispatch', edit_record)
                for field_name, value in payload.items():
                    setattr(edit_record, field_name, value)
                edit_record.save()

                if log_change('dispatch', edit_record, before_snapshot, request.user, 'update', change_reason):
                    messages.success(request, f'Dispatch updated for {job_card.job_card_no}')
                else:
                    messages.success(request, f'No changes detected for {job_card.job_card_no}')
                return redirect('dispatch_records')

            record = Dispatch.objects.create(**payload)
            record.created_by = request.user
            record.save(update_fields=['created_by'])
            log_change('dispatch', record, {}, request.user, 'create', 'Initial entry created')

            messages.success(request, f'Dispatch saved for {job_card.job_card_no}')
            return redirect('dispatch_entry')
        except Exception as e:
            messages.error(request, f'Error saving dispatch: {str(e)}')

    context = {
        'job_cards': JobCard.objects.filter(is_active=True).order_by('-created_at')[:200],
        'today': edit_record.dispatch_date if edit_record else timezone.now().date(),
        'edit_record': edit_record,
        'edit_lock_days': get_record_edit_lock_days(),
        'edit_lock_applies': bool(edit_record and record_is_time_locked('dispatch', edit_record)),
        'is_view_mode': is_view_mode,
    }
    return render(request, 'dispatch_entry.html', context)


@login_required
@permission_required('can_approve_dispatch')
def dispatch_records(request):
    """Dispatch records list page"""
    if request.method == 'POST':
        action = (request.POST.get('action') or '').strip()
        if action == 'bulk_delete':
            if request.user.profile.role != 'admin':
                add_unique_message(request, messages.ERROR, '❌ Only admin can run bulk delete.')
                return redirect('dispatch_records')
            if not user_can_archive_records(request.user):
                add_unique_message(request, messages.ERROR, '❌ You do not have permission to delete records.')
                return redirect('dispatch_records')

            selected_ids = request.POST.getlist('selected_ids')
            deleted_count, failures = run_bulk_permanent_delete(request, 'dispatch', selected_ids)
            if deleted_count:
                add_unique_message(request, messages.SUCCESS, f'Deleted {deleted_count} dispatch record(s) permanently.')
            if failures:
                add_unique_message(request, messages.ERROR, f'Bulk delete completed with issues: {"; ".join(failures[:5])}')
            return redirect('dispatch_records')

    query = (request.GET.get('q') or '').strip()
    date_from_raw = (request.GET.get('date_from') or '').strip()
    date_to_raw = (request.GET.get('date_to') or '').strip()
    sort = (request.GET.get('sort') or 'dispatch_date').strip()
    direction = (request.GET.get('dir') or 'desc').strip().lower()
    per_page = request.GET.get('per_page') or '50'
    try:
        per_page = int(per_page)
    except (TypeError, ValueError):
        per_page = 50
    if per_page not in (50, 100):
        per_page = 50

    records = Dispatch.objects.filter(is_active=True, job_card__is_active=True).select_related('job_card', 'created_by').order_by('-dispatch_date', '-id')
    if query:
        records = records.filter(
            Q(job_card__job_card_no__icontains=query) |
            Q(dc_no__icontains=query)
        )

    date_from = None
    date_to = None
    if date_from_raw:
        try:
            date_from = datetime.strptime(date_from_raw, '%Y-%m-%d').date()
            records = records.filter(dispatch_date__gte=date_from)
        except ValueError:
            add_unique_message(request, messages.ERROR, 'Invalid From date format. Use YYYY-MM-DD.')
            date_from_raw = ''
    if date_to_raw:
        try:
            date_to = datetime.strptime(date_to_raw, '%Y-%m-%d').date()
            records = records.filter(dispatch_date__lte=date_to)
        except ValueError:
            add_unique_message(request, messages.ERROR, 'Invalid To date format. Use YYYY-MM-DD.')
            date_to_raw = ''

    if date_from and date_to and date_from > date_to:
        add_unique_message(request, messages.ERROR, 'From date cannot be later than To date.')
        records = records.none()

    sortable_fields = {
        'date': 'dispatch_date',
        'job_card': 'job_card__job_card_no',
        'dc_no': 'dc_no',
        'qty': 'dispatch_qty',
        'added_by': 'created_by__username',
    }
    order_field = sortable_fields.get(sort, 'dispatch_date')
    if direction not in ('asc', 'desc'):
        direction = 'desc'
    ordering = order_field if direction == 'asc' else f'-{order_field}'
    records = records.order_by(ordering)

    total_count = records.count()
    paginator = Paginator(records, per_page)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)
    records = list(page_obj.object_list)

    cutoff = get_record_edit_lock_cutoff()
    pending_ids: set = set()
    approved_ids: set = set()
    if cutoff and not user_can_bypass_edit_lock(request.user):
        user_overrides = EditOverrideRequest.objects.filter(
            entity_type='dispatch',
            requested_by=request.user,
        ).values('record_id', 'status', 'expires_at')
        for ov in user_overrides:
            if ov['status'] == 'pending':
                pending_ids.add(ov['record_id'])
            elif ov['status'] == 'approved' and ov['expires_at'] and ov['expires_at'] > timezone.now():
                approved_ids.add(ov['record_id'])

    context = {
        'records': records,
        'page_obj': page_obj,
        'total_count': total_count,
        'q': query,
        'date_from': date_from_raw,
        'date_to': date_to_raw,
        'sort': sort,
        'dir': direction,
        'per_page': per_page,
        'edit_lock_days': get_record_edit_lock_days(),
        'edit_lock_cutoff': cutoff,
        'can_bypass_edit_lock': user_can_bypass_edit_lock(request.user),
        'pending_override_ids': pending_ids,
        'approved_override_ids': approved_ids,
    }
    return render(request, 'dispatch_records.html', context)


@login_required
def change_history(request, entity_type, record_id):
    config = AUDIT_CONFIG.get(entity_type)
    if not config:
        messages.error(request, 'Unsupported history request.')
        return redirect('home')

    if not user_has_entity_permission(request.user, entity_type):
        add_unique_message(request, messages.ERROR, '❌ You do not have permission to access this feature.')
        return redirect('home')

    record = get_object_or_404(config['model'], pk=record_id)
    history_entries = ChangeLog.objects.filter(entity_type=entity_type, record_id=record_id).select_related('changed_by')

    context = {
        'entity_type': entity_type,
        'entity_label': config['model']._meta.verbose_name.title(),
        'record': record,
        'history_entries': history_entries,
        'back_view_name': config['list_view'],
    }
    return render(request, 'change_history.html', context)


@login_required
def archived_records(request):
    if not user_can_archive_records(request.user):
        add_unique_message(request, messages.ERROR, '❌ You do not have permission to access this feature.')
        return redirect('home')

    accessible_entities = get_accessible_entities(request.user)

    requested_entity = (request.GET.get('entity') or '').strip().lower()
    if requested_entity == 'all':
        entity_type = 'all'
    elif requested_entity in accessible_entities:
        entity_type = requested_entity
    else:
        entity_type = 'all'

    query = (request.GET.get('q') or '').strip()

    def build_archived_queryset(entity):
        config = AUDIT_CONFIG[entity]
        records = config['model'].objects.filter(is_active=False)

        if entity == 'job_card':
            records = records.select_related('material', 'machine_name', 'department').order_by('-created_at')
            if query:
                records = records.filter(
                    Q(job_card_no__icontains=query) |
                    Q(SKU__icontains=query) |
                    Q(PO_No__icontains=query)
                )
        elif entity == 'production':
            records = records.select_related('job_card', 'machine', 'operator').order_by('-date', '-id')
            if query:
                records = records.filter(
                    Q(job_card__job_card_no__icontains=query) |
                    Q(machine__name__icontains=query) |
                    Q(operator__name__icontains=query)
                )
        else:
            records = records.select_related('job_card').order_by('-dispatch_date', '-id')
            if query:
                records = records.filter(
                    Q(job_card__job_card_no__icontains=query) |
                    Q(dc_no__icontains=query)
                )

        return records

    records = None
    records_by_entity = None
    if entity_type == 'all':
        records_by_entity = {entity: build_archived_queryset(entity) for entity in accessible_entities}
    else:
        records = build_archived_queryset(entity_type)

    context = {
        'accessible_entities': accessible_entities,
        'entity_type': entity_type,
        'query': query,
        'records': records,
        'records_by_entity': records_by_entity,
        'all_mode': entity_type == 'all',
    }
    return render(request, 'archived_records.html', context)


@login_required
def delete_record(request, entity_type, record_id):
    config = AUDIT_CONFIG.get(entity_type)
    if not config:
        raise Http404('Unsupported record type')

    if not user_can_archive_records(request.user):
        add_unique_message(request, messages.ERROR, '❌ You do not have permission to access this feature.')
        return redirect('home')

    record = get_active_record_or_404(config['model'], record_id)

    if request.method == 'POST':
        reason = (request.POST.get('delete_reason') or '').strip()
        if not reason:
            messages.error(request, 'Delete reason is required.')
        else:
            try:
                validate_delete_allowed(entity_type, record)
                archive_record(entity_type, record, request.user, reason)
                messages.success(request, f'{config["model"]._meta.verbose_name.title()} archived successfully.')
                return redirect(config['list_view'])
            except Exception as exc:
                messages.error(request, str(exc))

    context = {
        'entity_type': entity_type,
        'entity_label': config['model']._meta.verbose_name.title(),
        'record': record,
        'back_view_name': config['list_view'],
    }
    return render(request, 'confirm_delete.html', context)


@login_required
def restore_record(request, entity_type, record_id):
    config = AUDIT_CONFIG.get(entity_type)
    if not config:
        raise Http404('Unsupported record type')

    if not user_can_archive_records(request.user):
        add_unique_message(request, messages.ERROR, '❌ You do not have permission to access this feature.')
        return redirect('home')

    record = get_inactive_record_or_404(config['model'], record_id)

    if request.method == 'POST':
        reason = (request.POST.get('restore_reason') or '').strip()
        if not reason:
            messages.error(request, 'Restore reason is required.')
        else:
            try:
                validate_restore_allowed(entity_type, record)
                restore_record_state(entity_type, record, request.user, reason)
                messages.success(request, f'{config["model"]._meta.verbose_name.title()} restored successfully.')
                return redirect(f"{reverse('archived_records')}?entity={entity_type}")
            except Exception as exc:
                messages.error(request, str(exc))

    context = {
        'entity_type': entity_type,
        'entity_label': config['model']._meta.verbose_name.title(),
        'record': record,
    }
    return render(request, 'confirm_restore.html', context)


OVERRIDE_EDIT_WINDOW_HOURS = 2


@login_required
def request_edit_override(request, entity_type, record_id):
    """Operational user submits a reason-based request to edit a locked record."""
    config = AUDIT_CONFIG.get(entity_type)
    if config is None or entity_type not in ('job_card', 'production', 'dispatch'):
        messages.error(request, 'Override requests are not supported for this record type.')
        return redirect('home')

    if not user_has_entity_permission(request.user, entity_type):
        add_unique_message(request, messages.ERROR, '❌ You do not have permission to access this feature.')
        return redirect('home')

    record = get_active_record_or_404(config['model'], record_id)
    entry_view_name = config['list_view'].replace('_records', '_entry')

    if not record_is_time_locked(entity_type, record):
        messages.info(request, 'This record is not locked — you can edit it directly.')
        return redirect(f"{reverse(entry_view_name)}?edit={record_id}")

    if user_can_bypass_edit_lock(request.user):
        return redirect(f"{reverse(entry_view_name)}?edit={record_id}")

    existing = EditOverrideRequest.objects.filter(
        entity_type=entity_type,
        record_id=record_id,
        requested_by=request.user,
        status='pending',
    ).first()
    if existing:
        messages.info(request, 'You already have a pending override request for this record.')
        return redirect(config['list_view'])

    if request.method == 'POST':
        reason = (request.POST.get('reason') or '').strip()
        if not reason:
            messages.error(request, 'A reason is required for the override request.')
        else:
            EditOverrideRequest.objects.create(
                entity_type=entity_type,
                record_id=record_id,
                record_label=str(record),
                requested_by=request.user,
                reason=reason,
            )
            messages.success(request, 'Override request submitted. You will be able to edit once a manager approves it.')
            return redirect(config['list_view'])

    context = {
        'entity_type': entity_type,
        'entity_label': config['model']._meta.verbose_name.title(),
        'record': record,
        'back_view_name': config['list_view'],
        'override_hours': OVERRIDE_EDIT_WINDOW_HOURS,
    }
    return render(request, 'request_edit_override.html', context)


@login_required
def override_requests(request):
    """Manager/admin inbox of all override requests."""
    if not user_can_archive_records(request.user):
        add_unique_message(request, messages.ERROR, '❌ You do not have permission to access this feature.')
        return redirect('home')

    status_filter = (request.GET.get('status') or 'pending').strip()
    qs_all = EditOverrideRequest.objects.select_related('requested_by', 'reviewed_by').all()
    if status_filter in ('pending', 'approved', 'rejected'):
        qs = qs_all.filter(status=status_filter)
    else:
        qs = qs_all
        status_filter = 'all'

    context = {
        'override_list': qs,
        'status_filter': status_filter,
        'pending_count': qs_all.filter(status='pending').count(),
    }
    return render(request, 'override_requests.html', context)


@login_required
def review_override_request(request, override_id):
    """Manager/admin approves or rejects an override request."""
    if not user_can_archive_records(request.user):
        add_unique_message(request, messages.ERROR, '❌ You do not have permission to access this feature.')
        return redirect('home')

    override = get_object_or_404(EditOverrideRequest, pk=override_id, status='pending')

    if request.method == 'POST':
        action = (request.POST.get('action') or '').strip()
        note = (request.POST.get('review_note') or '').strip()

        if action not in ('approve', 'reject'):
            messages.error(request, 'Invalid action.')
        else:
            override.reviewed_by = request.user
            override.review_note = note
            override.reviewed_at = timezone.now()
            if action == 'approve':
                override.status = 'approved'
                override.expires_at = timezone.now() + timedelta(hours=OVERRIDE_EDIT_WINDOW_HOURS)
                messages.success(
                    request,
                    f'Override approved. {override.requested_by.get_full_name() or override.requested_by.username}'
                    f' can now edit the record for {OVERRIDE_EDIT_WINDOW_HOURS} hour(s).'
                )
            else:
                override.status = 'rejected'
                messages.success(request, 'Override request rejected.')
            override.save()
            return redirect('override_requests')

    context = {
        'override': override,
        'override_hours': OVERRIDE_EDIT_WINDOW_HOURS,
    }
    return render(request, 'review_override_request.html', context)


@login_required
@permission_required('can_view_analytics')
def production_dashboard(request):
    """Real-time production dashboard with OEE metrics"""
    from django.db.models import Count
    from datetime import timedelta

    # Flexible date filtering: explicit start/end date overrides quick day presets.
    today = timezone.now().date()
    start_date_input = (request.GET.get('start_date') or '').strip()
    end_date_input = (request.GET.get('end_date') or '').strip()

    start_date = None
    end_date = None
    if start_date_input and end_date_input:
        try:
            start_date = datetime.strptime(start_date_input, '%Y-%m-%d').date()
            end_date = datetime.strptime(end_date_input, '%Y-%m-%d').date()
            if start_date > end_date:
                start_date, end_date = end_date, start_date
        except ValueError:
            start_date = None
            end_date = None

    if start_date is None or end_date is None:
        try:
            days = int(request.GET.get('days', 7))
        except (TypeError, ValueError):
            days = 7
        if days < 1:
            days = 1
        end_date = today
        start_date = end_date - timedelta(days=days - 1)
    else:
        days = max((end_date - start_date).days + 1, 1)

    period_productions = Production.objects.filter(is_active=True, job_card__is_active=True, date__gte=start_date, date__lte=end_date)
    period_dispatches = Dispatch.objects.filter(is_active=True, job_card__is_active=True, dispatch_date__gte=start_date, dispatch_date__lte=end_date)

    # Calculate period metrics
    total_impressions = period_productions.aggregate(total=Sum('impressions'))['total'] or 0
    total_downtime = period_productions.aggregate(total=Sum('downtime'))['total'] or 0
    total_output = period_productions.aggregate(total=Sum('output_sheets'))['total'] or 0
    total_waste = period_productions.aggregate(total=Sum('waste_sheets'))['total'] or 0
    total_planned_minutes = period_productions.aggregate(total=Sum('planned_time'))['total'] or 0
    total_run_minutes = period_productions.aggregate(total=Sum('run_time'))['total'] or 0
    total_setup_minutes = period_productions.aggregate(total=Sum('setup_time'))['total'] or 0
    total_actual_minutes = float(total_run_minutes or 0) + float(total_setup_minutes or 0) + float(total_downtime or 0)
    planned_variance_minutes = total_actual_minutes - float(total_planned_minutes or 0)
    overrun_setup_minutes = float(total_setup_minutes or 0)
    overrun_downtime_minutes = float(sum(
        p.unplanned_downtime_minutes for p in period_productions
    ))
    overrun_run_perf_minutes = max(float(planned_variance_minutes) - overrun_setup_minutes - overrun_downtime_minutes, 0)

    # Dispatch metrics for same period
    total_dispatch_qty = period_dispatches.aggregate(total=Sum('dispatch_qty'))['total'] or 0
    dispatch_count = period_dispatches.count()
    dispatched_job_cards_count = period_dispatches.values('job_card').distinct().count()
    avg_dispatch_qty = (total_dispatch_qty / dispatch_count) if dispatch_count else 0
    dispatch_fulfillment_pct = (total_dispatch_qty / total_output * 100) if total_output > 0 else 0

    # Short-close decision tracking KPIs for manager visibility
    short_close_closed_period_qs = JobCard.objects.filter(
        is_active=True,
        short_close_closed_at__date__gte=start_date,
        short_close_closed_at__date__lte=end_date,
        short_close_wastage_qty__gt=0,
    )
    short_close_closed_period_qty = short_close_closed_period_qs.aggregate(
        total=Sum('short_close_wastage_qty')
    )['total'] or 0
    short_close_closed_period_jobs = short_close_closed_period_qs.count()

    pending_short_close_total = 0
    pending_short_close_jobs = 0
    pending_short_close_rows = JobCard.objects.filter(is_active=True).annotate(
        agg_dispatch=Coalesce(Sum('dispatch__dispatch_qty', filter=Q(dispatch__is_active=True)), 0)
    ).values('order_qty', 'agg_dispatch', 'short_close_closed_qty')
    for row in pending_short_close_rows:
        order_qty = int(row['order_qty'] or 0)
        dispatch_qty = int(row['agg_dispatch'] or 0)
        closed_qty = int(row['short_close_closed_qty'] or 0)
        if order_qty <= 0 or dispatch_qty >= order_qty:
            continue

        completion_pct = (dispatch_qty / order_qty) * 100
        if completion_pct < 95:
            continue

        pending_qty = max((order_qty - dispatch_qty) - closed_qty, 0)
        if pending_qty > 0:
            pending_short_close_total += pending_qty
            pending_short_close_jobs += 1

    # Calculate OEE from aggregated period data.
    # Use actual planned_time from entries (correct), not a fixed per-entry assumption.
    available_time_minutes = total_planned_minutes
    UNPLANNED_CATEGORIES = {'breakdown', 'operator', 'other'}
    from django.db.models import Case, When, FloatField as DjFloatField
    unplanned_downtime = sum(
        p.unplanned_downtime_minutes for p in period_productions
    )
    actual_run_time = available_time_minutes - unplanned_downtime
    availability = (actual_run_time / available_time_minutes * 100) if available_time_minutes > 0 else 0

    # Convert run-time minutes to expected impressions using each machine's standard speed.
    ideal_impressions = sum(
        ((p.run_time or 0) / 60) * ((p.machine.standard_impressions_per_hour if p.machine else 0) or 4000)
        for p in period_productions.select_related('machine')
    )

    performance = (total_impressions / ideal_impressions * 100) if ideal_impressions > 0 else 0
    performance = min(performance, 100)

    quality = ((total_output / (total_output + total_waste)) * 100) if (total_output + total_waste) > 0 else 0

    oee_value = (availability * performance * quality) / 10000

    # Schedule-based utilization (manager-defined weekly shift config)
    shift_configs = list(ShiftConfig.objects.all().order_by('effective_from', 'id'))
    machine_schedules = list(MachineWorkSchedule.objects.all().order_by('effective_from', 'id'))

    def resolve_shift_minutes(target_date, day_of_week, shift_code):
        matched = None
        for cfg in shift_configs:
            if cfg.day_of_week != day_of_week or cfg.shift != shift_code:
                continue
            if cfg.effective_from and cfg.effective_to:
                if cfg.effective_from <= target_date <= cfg.effective_to:
                    matched = cfg
            elif cfg.effective_from is None and cfg.effective_to is None and matched is None:
                matched = cfg
        if matched:
            return float(matched.net_hours or 0) * 60
        return default_shift_minutes

    def resolve_machine_working(machine_id, target_date, day_of_week, shift_code):
        matched = None
        for cfg in machine_schedules:
            if cfg.machine_id != machine_id or cfg.day_of_week != day_of_week or cfg.shift != shift_code:
                continue
            if cfg.effective_from and cfg.effective_to:
                if cfg.effective_from <= target_date <= cfg.effective_to:
                    matched = cfg
            elif cfg.effective_from is None and cfg.effective_to is None and matched is None:
                matched = cfg
        if matched is not None:
            return bool(matched.is_working)
        return True

    default_shift_minutes = 11 * 60

    scheduled_minutes_total = 0.0
    actual_used_minutes_total = 0.0
    seen_machine_sessions = set()
    for p in period_productions.select_related('machine'):
        day_of_week = p.date.weekday() if p.date else None
        if day_of_week is None or not p.machine_id:
            continue

        is_working = resolve_machine_working(p.machine_id, p.date, day_of_week, p.shift)
        if not is_working:
            continue

        session_key = (p.machine_id, p.date, p.shift)
        if session_key not in seen_machine_sessions:
            scheduled_minutes_total += resolve_shift_minutes(p.date, day_of_week, p.shift)
            seen_machine_sessions.add(session_key)

        actual_used_minutes_total += (p.run_time or 0) + (p.setup_time or 0) + (p.downtime or 0)

    schedule_utilization_pct = (actual_used_minutes_total / scheduled_minutes_total * 100) if scheduled_minutes_total > 0 else 0

    # Recent productions (last 10)
    recent_productions = period_productions.select_related('job_card', 'machine', 'operator').order_by('-date', '-id')[:10]
    recent_dispatches = period_dispatches.select_related('job_card').order_by('-dispatch_date', '-id')[:10]

    # Production summary by date
    production_by_date = period_productions\
        .values('date')\
        .annotate(
            total_impressions=Sum('impressions'),
            total_output=Sum('output_sheets'),
            total_waste=Sum('waste_sheets'),
            total_downtime=Sum('downtime'),
            count=Count('id')
        )\
        .order_by('-date')

    # Machine utilization
    machine_utilization = period_productions\
        .values('machine__name')\
        .annotate(
            total_impressions=Sum('impressions'),
            total_downtime=Sum('downtime'),
            production_count=Count('id')
        )\
        .order_by('-total_impressions')[:5]

    # Downtime analysis (multi-row details with fallback for legacy records)
    downtime_totals = {}
    downtime_counts = {}
    for production in period_productions.prefetch_related('downtime_entries'):
        entries = list(production.downtime_entries.all())
        if entries:
            for row in entries:
                key = row.category
                downtime_totals[key] = downtime_totals.get(key, 0.0) + float(row.minutes or 0)
                downtime_counts[key] = downtime_counts.get(key, 0) + 1
            continue

        if production.downtime_category and float(production.downtime or 0) > 0:
            key = production.downtime_category
            downtime_totals[key] = downtime_totals.get(key, 0.0) + float(production.downtime or 0)
            downtime_counts[key] = downtime_counts.get(key, 0) + 1

    downtime_by_category_data = [
        {
            'downtime_category': category,
            'downtime_label': dict(Production.DOWNTIME_CHOICES).get(category, category),
            'total_minutes': round(float(minutes or 0), 2),
            'count': int(downtime_counts.get(category, 0)),
        }
        for category, minutes in sorted(downtime_totals.items(), key=lambda item: item[1], reverse=True)[:5]
    ]

    # Dispatch trends and top dispatched job cards
    dispatch_by_date = period_dispatches\
        .values('dispatch_date')\
        .annotate(
            total_dispatch=Sum('dispatch_qty'),
            count=Count('id')
        )\
        .order_by('-dispatch_date')

    top_dispatched_jobcards = period_dispatches\
        .values('job_card__job_card_no')\
        .annotate(total_dispatch=Sum('dispatch_qty'))\
        .order_by('-total_dispatch')[:5]

    # ── Machine efficiency ranking ──────────────────────────────────────────────
    machine_efficiency_raw = period_productions \
        .values('machine__name', 'machine__standard_impressions_per_hour') \
        .annotate(
            m_output=Sum('output_sheets'),
            m_waste=Sum('waste_sheets'),
            m_impressions=Sum('impressions'),
            m_run_time=Sum('run_time'),
            m_planned_time=Sum('planned_time'),
            m_sessions=Count('id'),
        ).order_by('-m_impressions')

    machine_efficiency_data = []
    for m in machine_efficiency_raw:
        planned = m['m_planned_time'] or 0
        run = m['m_run_time'] or 0
        impressions = m['m_impressions'] or 0
        output = m['m_output'] or 0
        waste = m['m_waste'] or 0
        std_speed = m['machine__standard_impressions_per_hour'] or 4000
        avail = (run / planned * 100) if planned > 0 else 0
        ideal = (run / 60) * std_speed
        perf = min((impressions / ideal * 100) if ideal > 0 else 0, 100)
        qual = (output / (output + waste) * 100) if (output + waste) > 0 else 100
        oee = round((avail * perf * qual) / 10000, 1)
        machine_efficiency_data.append({
            'name': m['machine__name'] or 'Unknown',
            'oee': oee,
            'availability': round(avail, 1),
            'performance': round(perf, 1),
            'quality': round(qual, 1),
            'total_output': int(output),
            'total_waste': int(waste),
            'sessions': m['m_sessions'],
        })
    machine_efficiency_data.sort(key=lambda x: x['oee'], reverse=True)

    # ── Operator efficiency ranking ─────────────────────────────────────────────
    operator_efficiency_raw = period_productions \
        .values('operator__name', 'machine__standard_impressions_per_hour') \
        .annotate(
            o_output=Sum('output_sheets'),
            o_waste=Sum('waste_sheets'),
            o_impressions=Sum('impressions'),
            o_run_time=Sum('run_time'),
            o_planned_time=Sum('planned_time'),
            o_sessions=Count('id'),
        ).order_by('-o_impressions')

    operator_efficiency_data = []
    for o in operator_efficiency_raw:
        planned = o['o_planned_time'] or 0
        run = o['o_run_time'] or 0
        impressions = o['o_impressions'] or 0
        output = o['o_output'] or 0
        waste = o['o_waste'] or 0
        std_speed = o['machine__standard_impressions_per_hour'] or 4000
        avail = (run / planned * 100) if planned > 0 else 0
        ideal = (run / 60) * std_speed
        perf = min((impressions / ideal * 100) if ideal > 0 else 0, 100)
        qual = (output / (output + waste) * 100) if (output + waste) > 0 else 100
        oee = round((avail * perf * qual) / 10000, 1)
        operator_efficiency_data.append({
            'name': o['operator__name'] or 'Unknown',
            'oee': oee,
            'availability': round(avail, 1),
            'performance': round(perf, 1),
            'quality': round(qual, 1),
            'total_output': int(output),
            'sessions': o['o_sessions'],
        })
    operator_efficiency_data.sort(key=lambda x: x['oee'], reverse=True)

    # ── Waste Pareto (80/20) ────────────────────────────────────────────────────
    waste_by_reason_raw = period_productions \
        .exclude(waste_reason__isnull=True).exclude(waste_reason='') \
        .values('waste_reason') \
        .annotate(total_waste=Sum('waste_sheets')) \
        .order_by('-total_waste')

    total_waste_pareto = sum(w['total_waste'] or 0 for w in waste_by_reason_raw)
    cumulative = 0
    waste_pareto_data = []
    for w in waste_by_reason_raw:
        wval = w['total_waste'] or 0
        cumulative += wval
        waste_pareto_data.append({
            'reason': dict(Production.WASTE_CHOICES).get(w['waste_reason'], w['waste_reason']),
            'total_waste': int(wval),
            'pct': round(wval / total_waste_pareto * 100, 1) if total_waste_pareto > 0 else 0,
            'cumulative_pct': round(cumulative / total_waste_pareto * 100, 1) if total_waste_pareto > 0 else 0,
        })

    # ── Shift comparison (A vs B) ───────────────────────────────────────────────
    shift_raw = period_productions \
        .values('shift') \
        .annotate(
            s_output=Sum('output_sheets'),
            s_waste=Sum('waste_sheets'),
            s_impressions=Sum('impressions'),
            s_run_time=Sum('run_time'),
            s_planned_time=Sum('planned_time'),
            s_downtime=Sum('downtime'),
            s_sessions=Count('id'),
        )

    shift_comparison_data = {}
    for s in shift_raw:
        planned = s['s_planned_time'] or 0
        run = s['s_run_time'] or 0
        impressions = s['s_impressions'] or 0
        output = s['s_output'] or 0
        waste = s['s_waste'] or 0
        avail = (run / planned * 100) if planned > 0 else 0
        ideal = (run / 60) * 4000
        perf = min((impressions / ideal * 100) if ideal > 0 else 0, 100)
        qual = (output / (output + waste) * 100) if (output + waste) > 0 else 100
        oee = round((avail * perf * qual) / 10000, 1)
        shift_comparison_data[s['shift']] = {
            'shift': s['shift'],
            'oee': oee,
            'availability': round(avail, 1),
            'performance': round(perf, 1),
            'quality': round(qual, 1),
            'total_output': int(output),
            'total_waste': int(waste),
            'downtime': float(s['s_downtime'] or 0),
            'sessions': s['s_sessions'],
        }

    shift_a = shift_comparison_data.get('A', {'shift': 'A', 'oee': 0, 'availability': 0, 'performance': 0, 'quality': 0, 'total_output': 0, 'total_waste': 0, 'downtime': 0, 'sessions': 0})
    shift_b = shift_comparison_data.get('B', {'shift': 'B', 'oee': 0, 'availability': 0, 'performance': 0, 'quality': 0, 'total_output': 0, 'total_waste': 0, 'downtime': 0, 'sessions': 0})

    # ── Job cycle time analysis ─────────────────────────────────────────────────
    job_cycle_raw = JobCard.objects \
        .filter(is_active=True, productions__is_active=True) \
        .annotate(
            first_prod=Min('productions__date', filter=Q(productions__is_active=True)),
            last_prod=Max('productions__date', filter=Q(productions__is_active=True)),
            prod_output=Sum('productions__output_sheets', filter=Q(productions__is_active=True)),
        ) \
        .values('job_card_no', 'order_qty', 'status', 'first_prod', 'last_prod', 'prod_output') \
        .order_by('-last_prod')[:15]

    job_cycle_data = []
    for j in job_cycle_raw:
        first = j['first_prod']
        last = j['last_prod']
        cycle_days = (last - first).days + 1 if first and last else 0
        output = j['prod_output'] or 0
        job_cycle_data.append({
            'job_card_no': j['job_card_no'],
            'order_qty': j['order_qty'],
            'output': int(output),
            'status': j['status'],
            'first_prod': first.isoformat() if first else '-',
            'last_prod': last.isoformat() if last else '-',
            'cycle_days': cycle_days,
            'completion_pct': round(output / j['order_qty'] * 100, 1) if j['order_qty'] > 0 else 0,
        })

    # ── Predictive delay detection ──────────────────────────────────────────────
    from datetime import date as date_type
    today = timezone.now().date()
    at_risk_jobs = []
    open_jobs = JobCard.objects.filter(is_active=True, status__in=['Open', 'In Progress']) \
        .annotate(
            first_prod=Min('productions__date', filter=Q(productions__is_active=True)),
            last_prod=Max('productions__date', filter=Q(productions__is_active=True)),
            prod_output=Sum('productions__output_sheets', filter=Q(productions__is_active=True)),
        ).values('job_card_no', 'order_qty', 'status', 'first_prod', 'last_prod', 'prod_output')

    for j in open_jobs:
        output = j['prod_output'] or 0
        order_qty = j['order_qty'] or 0
        remaining = order_qty - output
        first = j['first_prod']
        if not first or remaining <= 0:
            continue
        days_active = (today - first).days or 1
        daily_rate = output / days_active
        if daily_rate <= 0:
            est_days_left = None
            risk = 'No Production'
        else:
            est_days_left = round(remaining / daily_rate)
            if est_days_left > 14:
                risk = 'High Risk'
            elif est_days_left > 7:
                risk = 'Medium Risk'
            else:
                risk = 'On Track'

        if risk in ('High Risk', 'Medium Risk', 'No Production'):
            at_risk_jobs.append({
                'job_card_no': j['job_card_no'],
                'order_qty': order_qty,
                'output': int(output),
                'remaining': int(remaining),
                'daily_rate': round(daily_rate, 0),
                'est_days_left': est_days_left,
                'risk': risk,
                'completion_pct': round(output / order_qty * 100, 1) if order_qty > 0 else 0,
            })
    at_risk_jobs.sort(key=lambda x: (0 if x['risk'] == 'No Production' else 1 if x['risk'] == 'High Risk' else 2))

    # ── Planned but not started jobs ────────────────────────────────────────────
    pending_start_qs = JobCard.objects.filter(is_active=True, status__in=['Open', 'In Progress']) \
        .annotate(prod_entries=Count('productions', filter=Q(productions__is_active=True))) \
        .filter(prod_entries=0)

    pending_start_count = pending_start_qs.count()
    pending_start_in_period_count = pending_start_qs.filter(
        created_at__date__gte=start_date,
        created_at__date__lte=end_date,
    ).count()

    pending_start_jobs = pending_start_qs.values(
        'job_card_no', 'SKU', 'order_qty', 'total_impressions_required', 'created_at'
    ).order_by('-created_at')[:15]

    pending_start_jobs_data = [
        {
            'job_card_no': item['job_card_no'],
            'sku': item['SKU'],
            'order_qty': int(item['order_qty'] or 0),
            'total_impressions_required': int(item['total_impressions_required'] or 0),
            'created_at': item['created_at'].date().isoformat() if item['created_at'] else '-',
        }
        for item in pending_start_jobs
    ]

    # Convert query data into JSON-serializable lists for charts.
    production_by_date_data = [
        {
            'date_only': item['date'].isoformat() if item['date'] else None,
            'total_impressions': float(item['total_impressions'] or 0),
            'total_output': float(item['total_output'] or 0),
            'total_waste': float(item['total_waste'] or 0),
            'total_downtime': float(item['total_downtime'] or 0),
            'count': int(item['count'] or 0),
        }
        for item in production_by_date
    ]

    machine_utilization_data = [
        {
            'machine_name': item['machine__name'] or 'Unknown',
            'total_impressions': float(item['total_impressions'] or 0),
            'total_downtime': float(item['total_downtime'] or 0),
            'production_count': int(item['production_count'] or 0),
        }
        for item in machine_utilization
    ]

    dispatch_by_date_data = [
        {
            'date_only': item['dispatch_date'].isoformat() if item['dispatch_date'] else None,
            'total_dispatch': float(item['total_dispatch'] or 0),
            'count': int(item['count'] or 0),
        }
        for item in dispatch_by_date
    ]

    top_dispatched_jobcards_data = [
        {
            'job_card_no': item['job_card__job_card_no'] or 'Unknown',
            'total_dispatch': float(item['total_dispatch'] or 0),
        }
        for item in top_dispatched_jobcards
    ]

    context = {
        'today': end_date,
        'days': days,
        'period_label': (
            f'{start_date.isoformat()} to {end_date.isoformat()}'
            if start_date_input and end_date_input
            else ('Today' if days == 1 else f'Last {days} Days')
        ),
        'start_date': start_date.isoformat() if start_date else '',
        'end_date': end_date.isoformat() if end_date else '',
        'oee_value': round(oee_value, 2),
        'availability_value': round(availability, 2),
        'performance_value': round(performance, 2),
        'quality_value': round(quality, 2),
        'total_impressions': total_impressions,
        'total_output': total_output,
        'total_waste': total_waste,
        'total_downtime': total_downtime,
        'total_planned_minutes': round(float(total_planned_minutes or 0), 2),
        'total_actual_minutes': round(float(total_actual_minutes or 0), 2),
        'planned_variance_minutes': round(float(planned_variance_minutes or 0), 2),
        'overrun_setup_minutes': round(float(overrun_setup_minutes or 0), 2),
        'overrun_downtime_minutes': round(float(overrun_downtime_minutes or 0), 2),
        'overrun_run_perf_minutes': round(float(overrun_run_perf_minutes or 0), 2),
        'schedule_utilization_pct': round(float(schedule_utilization_pct or 0), 2),
        'scheduled_minutes_total': round(float(scheduled_minutes_total or 0), 2),
        'actual_used_minutes_total': round(float(actual_used_minutes_total or 0), 2),
        'total_dispatch_qty': total_dispatch_qty,
        'dispatch_count': dispatch_count,
        'dispatched_job_cards_count': dispatched_job_cards_count,
        'avg_dispatch_qty': round(avg_dispatch_qty, 2),
        'dispatch_fulfillment_pct': round(dispatch_fulfillment_pct, 2),
        'short_close_closed_period_qty': int(short_close_closed_period_qty or 0),
        'short_close_closed_period_jobs': int(short_close_closed_period_jobs or 0),
        'pending_short_close_total': int(pending_short_close_total or 0),
        'pending_short_close_jobs': int(pending_short_close_jobs or 0),
        'recent_productions': recent_productions,
        'recent_dispatches': recent_dispatches,
        'production_by_date_json': json.dumps(production_by_date_data),
        'machine_utilization_json': json.dumps(machine_utilization_data),
        'downtime_by_category_json': json.dumps(downtime_by_category_data),
        'dispatch_by_date_json': json.dumps(dispatch_by_date_data),
        'top_dispatched_jobcards_json': json.dumps(top_dispatched_jobcards_data),
        'machine_efficiency': machine_efficiency_data,
        'operator_efficiency': operator_efficiency_data,
        'waste_pareto': waste_pareto_data,
        'waste_pareto_json': json.dumps(waste_pareto_data),
        'shift_a': shift_a,
        'shift_b': shift_b,
        'shift_comparison_json': json.dumps([shift_a, shift_b]),
        'job_cycle_data': job_cycle_data,
        'at_risk_jobs': at_risk_jobs,
        'pending_start_count': pending_start_count,
        'pending_start_in_period_count': pending_start_in_period_count,
        'pending_start_jobs': pending_start_jobs_data,
    }

    return render(request, 'production_dashboard.html', context)


def build_erp_readme_text():
    return """Offset Printing ERP - Easy User Guide

Last Updated: 2026-04-17

=============================
1) JOB CARD ENTRY (Planning)
=============================
Purpose:
- This is the planning sheet for one customer order.

Important fields (simple meaning):
- Job Card No: unique ID of the order.
- SKU: product code/name.
- PO Number / PO Date: customer purchase order details.
- Material: paper/board type.
- Colours: print colors (example 4 or 1+1).
- Order Qty (pcs): final quantity customer needs.
- UPS: how many pieces fit on one sheet.
- Wastage (sheets): planned extra sheets for setup/loss.
- Machine: planned machine for this job.
- Department: process area.
- Production Tolerance %: allowed extra production beyond plan.

Auto calculations:
- Required Sheets = Order Qty / UPS
- Total Planned Sheets = Required Sheets + Wastage
- Tolerance Sheets = Total Planned Sheets * Tolerance %
- Allowed Sheets = Planned Sheets + Tolerance Sheets

Time planning (auto):
- Run Time (min) = (Total Impressions Required / Machine Speed per hour) * 60
- Setup Time (min) = Total Colors * Machine Setup Minutes per Color
- Total Planned Time = Run Time + Setup Time

================================
2) PRODUCTION ENTRY (Execution)
================================
Purpose:
- Operator/supervisor logs actual production done in each shift.

Important fields:
- Job Card: which planned job is running.
- Machine: auto from Job Card (override allowed if needed).
- Operator, Shift, Date: who ran, when, which shift.
- Impressions: total machine impressions done.
- Output Sheets: good sheets produced.
- Waste Sheets + Waste Reason: scrap and reason.
- Planned Time: auto remaining planned minutes.
- Run Time, Setup Time, Downtime: actual consumed minutes.
- Downtime Category: reason bucket for downtime.

Validations:
- Output + Waste cannot exceed Allowed Sheets.
- Total impressions are cumulative across repeated production entries for the same job card.
- Total impressions are validated against Total Impressions Required plus Production Tolerance %.
- If planned allocation exceeds remaining planned minutes, overrun reason is mandatory.
- Overrun Minutes = (Run + Setup + Downtime) - Planned Time

================================
3) DISPATCH ENTRY (Delivery)
================================
Purpose:
- Record quantity sent to customer.

Important fields:
- Job Card
- Dispatch Date
- Dispatch Qty (pcs)
- DC No (optional)

Completion logic:
- Job is treated as Completed at 95%+ dispatch ratio.
- Remaining below 100% is Short Close (not auto waste).
- Manager/dispatch can close pending short-close with reason.
- Closed short-close is moved to Closed as Wastage.

================================
4) SHIFT & MACHINE SCHEDULE
================================
Purpose:
- Define realistic available capacity by shift and machine.

Shift hours fields:
- Effective From / Effective To: date range for this schedule version.
- Day + Shift A/B net hours: productive hours after breaks.

Machine work schedule fields:
- Check box ON = machine runs in that day+shift.
- Check box OFF = machine is not planned to run.

This is used in dashboard:
- Schedule Utilization % = Actual used minutes / Scheduled available minutes * 100

=============================
5) DASHBOARD (What it means)
=============================
Top KPIs:
- OEE: overall productivity quality score.
- Availability: uptime after unplanned downtime impact.
- Performance: speed efficiency vs machine standard speed.
- Quality: good output ratio.

Planning KPIs:
- Planned Time vs Actual Time
- Planned Variance
- Overrun split (setup, downtime, run-performance gap)

Dispatch and closure KPIs:
- Dispatch qty and fulfillment
- Pending Short Close
- Short Close Closed as Wastage

=============================
6) IF DROPDOWN VALUE IS WRONG
=============================
Example: wrong machine, operator, material, or department name added by mistake.

Use "Master Corrections" page:
1. Open the correction page from home/nav.
2. Rename wrong text to correct name.
3. For machine/operator, you can deactivate so it no longer appears in dropdowns.
4. Existing historical records remain safe and auditable.

=============================
7) QUICK RULES FOR USERS
=============================
- Always select correct Job Card first.
- Do not use waste to hide dispatch short-close.
- Total impressions are tracked across production entries and must stay within allowance.
- Give a reason when overriding machine or closing short-close.
- Keep shift schedule dates current for accurate utilization.

=============================
8) ARCHIVED RECORDS
=============================
- Use the Archived Records page to view deleted job cards, production, and dispatch entries.
- You can filter by entity type or use the All view to see every archived record type together.
- Restored records come back active with audit history intact.

=============================
9) MAINTENANCE NOTE
=============================
- Update this guide whenever fields, formulas, rules, or dashboard KPIs change.
"""


@login_required
def erp_readme(request):
    context = {
        'generated_on': timezone.now(),
    }
    return render(request, 'erp_readme.html', context)


@login_required
def download_erp_readme(request):
    content = build_erp_readme_text()
    response = HttpResponse(content, content_type='text/plain; charset=utf-8')
    response['Content-Disposition'] = 'attachment; filename="offset_erp_calculation_guide.txt"'
    return response


@login_required
@permission_required('can_manage_masters')
def machine_master_tools(request):
    """Manager/admin screen to correct dropdown master values across ERP."""

    model_map = {
        'machine': Machine,
        'operator': Operator,
        'material': Material,
        'department': Department,
    }

    if request.method == 'POST':
        action = (request.POST.get('action') or '').strip()
        entity_type = (request.POST.get('entity_type') or '').strip().lower()
        machine_id = (request.POST.get('machine_id') or request.POST.get('record_id') or '').strip()
        model = model_map.get(entity_type)
        record = get_object_or_404(model, pk=machine_id) if (model and machine_id) else None
        is_admin_user = bool(getattr(request.user, 'profile', None) and request.user.profile.role == 'admin')

        if action in {'rename_machine', 'edit_master', 'delete_master'} and not is_admin_user:
            messages.error(request, 'Only admin can edit or delete master values.')
            return redirect('machine_master_tools')

        if action in {'rename_machine', 'edit_master'} and record:
            new_name = (request.POST.get('new_name') or '').strip()
            if not new_name:
                messages.error(request, 'Name is required.')
            else:
                duplicate = model.objects.exclude(pk=record.pk).filter(name__iexact=new_name).first()
                if duplicate:
                    messages.error(request, f'Name already exists as #{duplicate.id} ({duplicate.name}).')
                else:
                    changed_fields = []
                    old_name = record.name

                    if record.name != new_name:
                        record.name = new_name
                        changed_fields.append('name')

                    if entity_type == 'operator':
                        new_employee_code = (request.POST.get('employee_code') or '').strip() or None
                        if record.employee_code != new_employee_code:
                            record.employee_code = new_employee_code
                            changed_fields.append('employee_code')

                    if entity_type == 'machine':
                        speed_raw = (request.POST.get('standard_impressions_per_hour') or '').strip()
                        setup_raw = (request.POST.get('standard_setup_minutes_per_color') or '').strip()

                        try:
                            new_speed = float(speed_raw) if speed_raw else float(record.standard_impressions_per_hour or 0)
                            new_setup = float(setup_raw) if setup_raw else float(record.standard_setup_minutes_per_color or 0)
                        except ValueError:
                            messages.error(request, 'Machine speed and setup minutes per color must be numeric values.')
                            return redirect('machine_master_tools')

                        if new_speed <= 0 or new_setup <= 0:
                            messages.error(request, 'Machine speed and setup minutes per color must be greater than 0.')
                            return redirect('machine_master_tools')

                        if float(record.standard_impressions_per_hour) != float(new_speed):
                            record.standard_impressions_per_hour = new_speed
                            changed_fields.append('standard_impressions_per_hour')
                        if float(record.standard_setup_minutes_per_color) != float(new_setup):
                            record.standard_setup_minutes_per_color = new_setup
                            changed_fields.append('standard_setup_minutes_per_color')

                    if changed_fields:
                        record.save(update_fields=changed_fields)
                        if old_name != record.name:
                            messages.success(request, f'{entity_type.title()} updated: {old_name} -> {record.name}')
                        else:
                            messages.success(request, f'{entity_type.title()} details updated successfully.')
                    else:
                        messages.success(request, f'No changes detected for {entity_type.title()} {record.name}.')

        elif action == 'toggle_machine' and record and hasattr(record, 'is_active'):
            record.is_active = not record.is_active
            record.save(update_fields=['is_active'])
            state = 'active' if record.is_active else 'inactive'
            messages.success(request, f'{entity_type.title()} {record.name} marked {state}.')

        elif action == 'delete_master' and record:
            record_name = record.name

            if entity_type == 'machine':
                linked_jobcards = JobCard.objects.filter(machine_name=record).count()
                linked_productions = Production.objects.filter(machine=record).count()
                if linked_jobcards or linked_productions:
                    messages.error(
                        request,
                        f'Cannot delete Machine {record_name}. Linked records found (Job Cards: {linked_jobcards}, Production: {linked_productions}).'
                    )
                    return redirect('machine_master_tools')

            elif entity_type == 'operator':
                linked_productions = Production.objects.filter(operator=record).count()
                if linked_productions:
                    messages.error(
                        request,
                        f'Cannot delete Operator {record_name}. Linked production records: {linked_productions}.'
                    )
                    return redirect('machine_master_tools')

            elif entity_type == 'material':
                linked_jobcards = JobCard.objects.filter(material=record).count()
                if linked_jobcards:
                    messages.error(
                        request,
                        f'Cannot delete Material {record_name}. Linked job cards: {linked_jobcards}.'
                    )
                    return redirect('machine_master_tools')

            elif entity_type == 'department':
                linked_jobcards = JobCard.objects.filter(department=record).count()
                if linked_jobcards:
                    messages.error(
                        request,
                        f'Cannot delete Department {record_name}. Linked job cards: {linked_jobcards}.'
                    )
                    return redirect('machine_master_tools')

            try:
                record.delete()
                messages.success(request, f'{entity_type.title()} {record_name} deleted successfully.')
            except ProtectedError:
                messages.error(request, f'Cannot delete {entity_type.title()} {record_name} because it is referenced by other records.')

        return redirect('machine_master_tools')

    machine_rows = []
    for item in Machine.objects.all().order_by('name', 'id'):
        machine_rows.append({
            'record': item,
            'job_card_count': JobCard.objects.filter(machine_name=item, is_active=True).count(),
            'production_count': Production.objects.filter(machine=item, is_active=True).count(),
        })

    operator_rows = []
    for item in Operator.objects.all().order_by('name', 'id'):
        operator_rows.append({
            'record': item,
            'production_count': Production.objects.filter(operator=item, is_active=True).count(),
        })

    material_rows = []
    for item in Material.objects.all().order_by('name', 'id'):
        material_rows.append({
            'record': item,
            'job_card_count': JobCard.objects.filter(material=item, is_active=True).count(),
        })

    department_rows = []
    for item in Department.objects.all().order_by('name', 'id'):
        department_rows.append({
            'record': item,
            'job_card_count': JobCard.objects.filter(department=item, is_active=True).count(),
        })

    context = {
        'machine_rows': machine_rows,
        'operator_rows': operator_rows,
        'material_rows': material_rows,
        'department_rows': department_rows,
        'is_admin_user': bool(getattr(request.user, 'profile', None) and request.user.profile.role == 'admin'),
    }
    return render(request, 'machine_master_tools.html', context)


@login_required
@require_role('admin')
def manage_user_roles(request):
    """Admin interface for managing user roles"""
    from django.contrib.auth import get_user_model
    
    User = get_user_model()
    users = User.objects.select_related('profile').all()
    
    if request.method == 'POST':
        user_id = request.POST.get('user_id')
        new_role = request.POST.get('role')
        
        try:
            user = User.objects.get(pk=user_id)
            profile = user.profile
            profile.role = new_role
            profile.save()
            messages.success(request, f'✅ {user.username} role updated to {profile.get_role_display()}')
        except (User.DoesNotExist, UserProfile.DoesNotExist) as e:
            messages.error(request, f'❌ Error updating role: {str(e)}')
        return redirect('manage_user_roles')
    
    context = {
        'users': users,
        'role_choices': UserProfile.ROLE_CHOICES,
    }
    return render(request, 'manage_user_roles.html', context)


@login_required
def download_template(request):
    """Download template in CSV or Excel format"""
    file_format = request.GET.get('format', 'csv').lower()
    
    headers = get_template_headers()
    example = get_template_example()
    
    if file_format == 'excel' and EXCEL_AVAILABLE:
        try:
            # Generate Excel file
            workbook = openpyxl.Workbook()
            worksheet = workbook.active
            worksheet.title = "Job Cards"
            
            # Add headers with styling
            header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
            header_font = Font(bold=True, color="FFFFFF")
            
            for col_num, header in enumerate(headers, 1):
                cell = worksheet.cell(row=1, column=col_num, value=header)
                cell.fill = header_fill
                cell.font = header_font
            
            # Add example row
            for col_num, value in enumerate(example, 1):
                worksheet.cell(row=2, column=col_num, value=value)
            
            # Auto-adjust column widths
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
            
            # Send file
            response = HttpResponse(
                content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
            response['Content-Disposition'] = 'attachment; filename="jobcard_template.xlsx"'
            workbook.save(response)
            return response
        except Exception as e:
            # Fallback to CSV if Excel generation fails
            print(f"Excel generation failed: {e}")
            import traceback
            traceback.print_exc()
    
    # Generate CSV file (default/fallback)
    response = HttpResponse(content_type='text/csv')
    response['Content-Disposition'] = 'attachment; filename="jobcard_template.csv"'
    
    writer = csv.writer(response)
    writer.writerow(headers)
    writer.writerow(example)
    
    return response


@login_required
@require_role('admin', 'manager')
def shift_config(request):
    """Manage weekly shift net hours and machine work schedules."""
    days = ShiftConfig.DAY_CHOICES
    shifts = ['A', 'B']
    machines = Machine.objects.filter(
        Q(is_active=True) |
        Q(id__in=JobCard.objects.exclude(machine_name__isnull=True).values_list('machine_name_id', flat=True)) |
        Q(production__isnull=False)
    ).distinct().order_by('name', 'id')

    active_date_raw = (request.GET.get('effective_date') or '').strip()
    if active_date_raw:
        try:
            active_date = datetime.strptime(active_date_raw, '%Y-%m-%d').date()
        except ValueError:
            active_date = timezone.now().date()
    else:
        active_date = timezone.now().date()

    def parse_effective_range(post_data):
        start_raw = (post_data.get('effective_from') or '').strip()
        end_raw = (post_data.get('effective_to') or '').strip()

        if not start_raw and not end_raw:
            return (None, None, None)
        if not start_raw or not end_raw:
            return (None, None, 'Both Effective From and Effective To are required.')

        try:
            start_date = datetime.strptime(start_raw, '%Y-%m-%d').date()
            end_date = datetime.strptime(end_raw, '%Y-%m-%d').date()
        except ValueError:
            return (None, None, 'Invalid effective date format.')

        if start_date > end_date:
            return (None, None, 'Effective From cannot be after Effective To.')

        return (start_date, end_date, None)

    def range_filter_for(target_date):
        return (
            Q(effective_from__isnull=True, effective_to__isnull=True) |
            Q(effective_from__lte=target_date, effective_to__gte=target_date)
        )

    if request.method == 'POST':
        action = request.POST.get('action')
        effective_from, effective_to, range_error = parse_effective_range(request.POST)

        if range_error:
            messages.error(request, range_error)
            return redirect('shift_config')

        if action == 'save_hours':
            for day_val, _ in days:
                for shift_val in shifts:
                    key = f'hours_{day_val}_{shift_val}'
                    raw = (request.POST.get(key) or '').strip()
                    if not raw:
                        continue
                    try:
                        net_h = float(raw)
                    except ValueError:
                        continue
                    ShiftConfig.objects.update_or_create(
                        day_of_week=day_val,
                        shift=shift_val,
                        effective_from=effective_from,
                        effective_to=effective_to,
                        defaults={'net_hours': net_h},
                    )
            messages.success(request, 'Shift hours saved.')

        elif action == 'save_schedule':
            for machine in machines:
                for day_val, _ in days:
                    for shift_val in shifts:
                        key = f'work_{machine.id}_{day_val}_{shift_val}'
                        is_working = request.POST.get(key) == 'on'
                        MachineWorkSchedule.objects.update_or_create(
                            machine=machine,
                            day_of_week=day_val,
                            shift=shift_val,
                            effective_from=effective_from,
                            effective_to=effective_to,
                            defaults={'is_working': is_working},
                        )
            messages.success(request, 'Machine schedule saved.')

        return redirect('shift_config')

    raw_hours = {}
    for sc in ShiftConfig.objects.filter(range_filter_for(active_date)).order_by('day_of_week', 'shift', 'effective_from', 'id'):
        raw_hours[(sc.day_of_week, sc.shift)] = sc.net_hours
    shift_hours_rows = []
    for day_val, day_name in days:
        shift_hours_rows.append({
            'day_val': day_val,
            'day_name': day_name,
            'A': raw_hours.get((day_val, 'A'), ''),
            'B': raw_hours.get((day_val, 'B'), ''),
        })

    working_map = {}
    for ms in MachineWorkSchedule.objects.filter(range_filter_for(active_date)).order_by('effective_from', 'id'):
        working_map[f'{ms.machine_id}_{ms.day_of_week}_{ms.shift}'] = bool(ms.is_working)

    machine_rows = []
    for machine in machines:
        cells = []
        for day_val, _ in days:
            for shift_val in shifts:
                key = f'{machine.id}_{day_val}_{shift_val}'
                cells.append({
                    'name': f'work_{machine.id}_{day_val}_{shift_val}',
                    'is_working': working_map.get(key, True),
                })
        machine_rows.append({'machine': machine, 'cells': cells})

    context = {
        'days': days,
        'shifts': shifts,
        'machines': machines,
        'shift_hours_rows': shift_hours_rows,
        'machine_rows': machine_rows,
        'active_effective_date': active_date.isoformat(),
        'active_effective_from': '',
        'active_effective_to': '',
    }
    return render(request, 'shift_config.html', context)