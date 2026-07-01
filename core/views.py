from django.shortcuts import render, get_object_or_404, redirect
from django.conf import settings
from django.http import HttpResponse, JsonResponse, Http404
from decimal import Decimal, InvalidOperation
from django.contrib import messages
from django.contrib.auth import get_user_model
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
from .models import (
    ChangeLog,
    Department,
    Dispatch,
    EditOverrideRequest,
    JOB_CARD_DISPATCHABLE_STATUSES,
    JOB_CARD_PRODUCTION_START_STATUSES,
    JobCard,
    Machine,
    MachineWorkSchedule,
    Material,
    Operator,
    Production,
    ProductionDowntime,
    ShiftConfig,
    Supervisor,
    UserProfile,
    Vendor,
)
from .jobcard_service import normalize_job_card_status
from workflow.services import start_production

from .constants import AUDIT_CONFIG
from .services import (
    add_unique_message,
    format_audit_value, normalize_colour_notation, extract_total_colors,
    compute_planned_minutes, get_remaining_planned_minutes, build_audit_snapshot,
    build_change_summary, log_change, user_has_entity_permission,
    user_can_archive_records, user_can_bypass_edit_lock, get_record_edit_lock_days,
    get_record_edit_lock_cutoff, record_is_time_locked, get_valid_override,
    ensure_edit_lock_allowed, get_accessible_entities, get_active_record_or_404,
    validate_delete_allowed,
    validate_restore_allowed, archive_record, restore_record_state,
    run_bulk_archive, run_bulk_permanent_delete
)


def _parse_optional_decimal(raw_value):
    if raw_value is None:
        return None
    raw_text = str(raw_value).strip().replace(',', '')
    if not raw_text:
        return None
    try:
        return Decimal(raw_text)
    except InvalidOperation:
        return None

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False


# ═══════════════════════════════════
# PERMISSION DECORATORS (RBAC)
# ═══════════════════════════════════


def require_role(*allowed_roles):
    """Decorator to check if user has required role"""
    def decorator(view_func):
        @wraps(view_func)
        def wrapped_view(request, *args, **kwargs):
            if not request.user.is_authenticated:
                return redirect('login')
            try:
                profile = request.user.profile
                user_role = (profile.role or '').strip().lower()
                if user_role not in [role.lower() for role in allowed_roles] and not request.user.is_staff:
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

    if master_type not in {'material', 'machine', 'department', 'operator', 'vendor'}:
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
        'vendor': Vendor,
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
    """Redirect legacy production job card entry to planning."""
    return redirect('planning:job_cards')


@login_required
@permission_required('can_edit_jobcard')
def job_card_records(request):
    """Legacy job card records redirect."""
    return redirect('planning:job_cards')


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
            if not (job_card.PO_No or '').strip():
                raise ValueError('PO number is required before dispatch. Link PO to this job first.')
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

    dispatch_jobs = JobCard.objects.filter(is_active=True, status__in=JOB_CARD_DISPATCHABLE_STATUSES)
    if edit_record:
        dispatch_jobs = JobCard.objects.filter(is_active=True).filter(
            Q(status__in=JOB_CARD_DISPATCHABLE_STATUSES) | Q(pk=edit_record.job_card_id)
        ).distinct()

    context = {
        'job_cards': dispatch_jobs.order_by('-created_at')[:200],
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

    User = get_user_model()
    model_map = {
        'machine': Machine,
        'operator': Operator,
        'material': Material,
        'department': Department,
        'supervisor': Supervisor,
        'vendor': Vendor,
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
                    return redirect('machine_master_tools')
                old_name = record.name

                changed_fields = []

                if record.name != new_name:
                    record.name = new_name
                    changed_fields.append('name')

                if entity_type in {'operator', 'supervisor'}:
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
            display_name = record.name
            messages.success(request, f'{entity_type.title()} {display_name} marked {state}.')

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

            elif entity_type == 'supervisor':
                linked_productions = Production.objects.filter(supervisor=record).count()
                if linked_productions:
                    messages.error(
                        request,
                        f'Cannot delete Supervisor {record_name}. Linked production records: {linked_productions}.'
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

    supervisor_rows = []
    for item in Supervisor.objects.all().order_by('name', 'id'):
        supervisor_rows.append({
            'record': item,
            'production_count': Production.objects.filter(supervisor=item, is_active=True).count(),
        })

    vendor_rows = []
    for item in Vendor.objects.all().order_by('name', 'id'):
        from printing_plates.models import PlateRequest
        vendor_rows.append({
            'record': item,
            'plate_request_count': PlateRequest.objects.filter(vendor=item.name).count(),
        })

    context = {
        'machine_rows': machine_rows,
        'operator_rows': operator_rows,
        'supervisor_rows': supervisor_rows,
        'material_rows': material_rows,
        'department_rows': department_rows,
        'vendor_rows': vendor_rows,
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