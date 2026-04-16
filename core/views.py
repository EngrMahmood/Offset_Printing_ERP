from django.shortcuts import render, get_object_or_404, redirect
from django.conf import settings
from django.http import HttpResponse, JsonResponse, Http404
from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.utils import timezone
from django.views.decorators.http import require_POST
from django.db.models import Q, Sum, Min, Max
from django.urls import reverse
from functools import wraps
import csv
import json
from datetime import datetime, date, timedelta

from .bulk_upload import process_jobcard_upload, get_template_headers, get_template_example
from .models import JobCard, Production, Machine, Operator, Department, Material, Dispatch, UserProfile, ChangeLog, EditOverrideRequest

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


def build_audit_snapshot(entity_type, instance):
    config = AUDIT_CONFIG[entity_type]
    return {
        field_name: format_audit_value(getattr(instance, field_name))
        for field_name in config['fields']
    }


def build_change_summary(entity_type, before_snapshot, after_snapshot):
    config = AUDIT_CONFIG[entity_type]
    changes = {}

    for field_name in config['fields']:
        old_value = before_snapshot.get(field_name, '-')
        new_value = after_snapshot.get(field_name, '-')
        if old_value != new_value:
            changes[field_name] = {
                'label': config['labels'].get(field_name, field_name.replace('_', ' ').title()),
                'from': old_value,
                'to': new_value,
            }

    return changes


def log_change(entity_type, instance, before_snapshot, changed_by, action, reason=''):
    after_snapshot = build_audit_snapshot(entity_type, instance)
    summary = build_change_summary(entity_type, before_snapshot, after_snapshot)

    if action == 'delete' and not summary:
        summary = {
            'record_state': {
                'label': 'Record State',
                'from': 'Active',
                'to': 'Archived',
            }
        }
    elif action == 'restore' and not summary:
        summary = {
            'record_state': {
                'label': 'Record State',
                'from': 'Archived',
                'to': 'Active',
            }
        }

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


@login_required
def home(request):
    return render(request, 'home.html')


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

        result = process_jobcard_upload(file)
        context = result

        return render(request, "upload.html", context)

    return render(request, "upload.html", context)


@login_required
@require_POST
def quick_add_master(request):
    """Create master dropdown values for planner workflow without admin dependency."""
    master_type = (request.POST.get('type') or '').strip().lower()
    name = (request.POST.get('name') or '').strip()

    if master_type not in {'material', 'machine', 'department'}:
        return JsonResponse({'ok': False, 'error': 'Invalid master type.'}, status=400)

    if not name:
        return JsonResponse({'ok': False, 'error': 'Name is required.'}, status=400)

    model_map = {
        'material': Material,
        'machine': Machine,
        'department': Department,
    }
    model = model_map[master_type]

    # Case-insensitive duplicate check for better UX.
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
    edit_id = (request.POST.get('edit_id') or request.GET.get('edit') or '').strip()
    edit_record = get_active_record_or_404(JobCard, edit_id) if edit_id else None

    if request.method == "POST":
        try:
            change_reason = (request.POST.get('change_reason') or '').strip()
            job_card_no = (request.POST.get('job_card_no') or '').strip()
            sku = (request.POST.get('sku') or '').strip()
            order_qty = int(request.POST.get('order_qty') or 0)

            if edit_record and not change_reason:
                raise ValueError("Change reason is required when editing a job card")
            if not job_card_no:
                raise ValueError("Job card number is required")
            if not sku:
                raise ValueError("SKU is required")
            if order_qty <= 0:
                raise ValueError("Order quantity must be greater than 0")

            duplicate_query = JobCard.objects.filter(job_card_no=job_card_no)
            if edit_record:
                duplicate_query = duplicate_query.exclude(pk=edit_record.pk)
            if duplicate_query.exists():
                raise ValueError(f"Job card number {job_card_no} already exists")

            po_date_raw = (request.POST.get('po_date') or '').strip()
            po_date = datetime.strptime(po_date_raw, "%Y-%m-%d").date() if po_date_raw else None
            month_value = po_date.strftime("%B") if po_date else ((request.POST.get('month') or '').strip() or None)

            material = Material.objects.filter(id=request.POST.get('material')).first() if request.POST.get('material') else None
            machine = Machine.objects.filter(id=request.POST.get('machine_name')).first() if request.POST.get('machine_name') else None
            department = Department.objects.filter(id=request.POST.get('department')).first() if request.POST.get('department') else None

            payload = {
                'job_card_no': job_card_no,
                'month': month_value,
                'po_date': po_date,
                'PO_No': (request.POST.get('po_no') or '').strip() or None,
                'SKU': sku,
                'material': material,
                'colour': int(request.POST.get('colour') or 0) or None,
                'application': (request.POST.get('application') or '').strip() or None,
                'order_qty': order_qty,
                'total_impressions_required': int(request.POST.get('total_impressions_required') or 0) or None,
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
            log_change('job_card', job_card, {}, request.user, 'create', 'Initial entry created')

            messages.success(request, f"Job card {job_card_no} created successfully")
            return redirect('job_card_entry')

        except Exception as e:
            messages.error(request, f"Error creating job card: {str(e)}")

    context = {
        'today': edit_record.po_date if edit_record and edit_record.po_date else timezone.now().date(),
        'materials': Material.objects.all().order_by('name'),
        'machines': Machine.objects.filter(is_active=True).order_by('name'),
        'departments': Department.objects.all().order_by('name'),
        'edit_record': edit_record,
    }
    return render(request, 'job_card_entry.html', context)


@login_required
@permission_required('can_edit_jobcard')
def job_card_records(request):
    """Job card records list page"""
    query = (request.GET.get('q') or '').strip()
    status = (request.GET.get('status') or '').strip()

    jobcards = JobCard.objects.filter(is_active=True).select_related('material', 'machine_name', 'department').order_by('-created_at')

    if query:
        jobcards = jobcards.filter(
            Q(job_card_no__icontains=query) |
            Q(SKU__icontains=query) |
            Q(PO_No__icontains=query)
        )

    if status:
        jobcards = jobcards.filter(status=status)

    context = {
        'jobcards': jobcards,
        'q': query,
        'status': status,
    }
    return render(request, 'job_card_records.html', context)


@login_required
@permission_required('can_edit_production')
def production_entry(request):
    """Production data entry form for operators"""
    edit_id = (request.POST.get('edit_id') or request.GET.get('edit') or '').strip()
    edit_record = get_active_record_or_404(Production, edit_id) if edit_id else None
    if edit_record and not ensure_edit_lock_allowed(request, 'production', edit_record):
        return redirect('production_records')

    if request.method == "POST":
        job_card_id = request.POST.get('job_card')
        machine_id = request.POST.get('machine')
        operator_id = request.POST.get('operator')
        shift = request.POST.get('shift')
        date = request.POST.get('date')
        impressions = request.POST.get('impressions')
        output_sheets = request.POST.get('output_sheets')
        waste_sheets = request.POST.get('waste_sheets')
        downtime_minutes = request.POST.get('downtime_minutes')
        planned_time = request.POST.get('planned_time')
        run_time = request.POST.get('run_time')
        setup_time = request.POST.get('setup_time')
        downtime_category = request.POST.get('downtime_category')
        waste_reason = request.POST.get('waste_reason')
        remarks = request.POST.get('remarks')

        try:
            change_reason = (request.POST.get('change_reason') or '').strip()
            output_sheets_val = int(output_sheets) if output_sheets else 0
            waste_sheets_val = int(waste_sheets) if waste_sheets else 0
            downtime_val = float(downtime_minutes) if downtime_minutes else 0
            planned_time_val = float(planned_time) if planned_time else 0
            run_time_val = float(run_time) if run_time else 0
            setup_time_val = float(setup_time) if setup_time else 0

            if edit_record and not change_reason:
                raise ValueError("Change reason is required when editing production data")
            if planned_time_val <= 0:
                raise ValueError("Planned time must be greater than 0")
            if run_time_val <= 0:
                raise ValueError("Run time must be greater than 0")
            if downtime_val > 0 and not downtime_category:
                raise ValueError("Downtime category is required when downtime is greater than 0")
            if waste_sheets_val > 0 and not waste_reason:
                raise ValueError("Waste reason is required when waste sheets are greater than 0")
            if (run_time_val + downtime_val + setup_time_val) > planned_time_val:
                raise ValueError("Run time + downtime + setup time cannot exceed planned time")

            job_card = get_active_record_or_404(JobCard, job_card_id)
            machine = get_object_or_404(Machine, pk=machine_id)
            operator = get_object_or_404(Operator, pk=operator_id)

            payload = {
                'job_card': job_card,
                'machine': machine,
                'operator': operator,
                'shift': shift,
                'date': date,
                'impressions': int(impressions) if impressions else 0,
                'output_sheets': output_sheets_val,
                'waste_sheets': waste_sheets_val,
                'planned_time': planned_time_val,
                'run_time': run_time_val,
                'setup_time': setup_time_val,
                'downtime': downtime_val,
                'downtime_category': downtime_category,
                'waste_reason': waste_reason,
            }

            if edit_record:
                before_snapshot = build_audit_snapshot('production', edit_record)
                for field_name, value in payload.items():
                    setattr(edit_record, field_name, value)
                edit_record.save()

                if log_change('production', edit_record, before_snapshot, request.user, 'update', change_reason):
                    messages.success(request, f'Production record updated for Job Card {job_card.job_card_no}')
                else:
                    messages.success(request, f'No changes detected for Job Card {job_card.job_card_no}')
                return redirect('production_records')

            record = Production.objects.create(**payload)
            log_change('production', record, {}, request.user, 'create', 'Initial entry created')

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
        'today': edit_record.date if edit_record else timezone.now().date(),
        'edit_record': edit_record,
        'edit_lock_days': get_record_edit_lock_days(),
        'edit_lock_applies': bool(edit_record and record_is_time_locked('production', edit_record)),
    }

    return render(request, 'production_entry.html', context)


@login_required
@permission_required('can_edit_production')
def production_records(request):
    """Production records list page"""
    query = (request.GET.get('q') or '').strip()
    shift = (request.GET.get('shift') or '').strip()

    records = Production.objects.filter(is_active=True, job_card__is_active=True).select_related('job_card', 'machine', 'operator').order_by('-date', '-id')

    if query:
        records = records.filter(
            Q(job_card__job_card_no__icontains=query) |
            Q(machine__name__icontains=query) |
            Q(operator__name__icontains=query)
        )

    if shift:
        records = records.filter(shift=shift)

    cutoff = get_record_edit_lock_cutoff()
    pending_ids: set = set()
    approved_ids: set = set()
    if cutoff and not user_can_bypass_edit_lock(request.user):
        user_overrides = EditOverrideRequest.objects.filter(
            entity_type='production',
            requested_by=request.user,
        ).values('record_id', 'status', 'expires_at')
        for ov in user_overrides:
            if ov['status'] == 'pending':
                pending_ids.add(ov['record_id'])
            elif ov['status'] == 'approved' and ov['expires_at'] and ov['expires_at'] > timezone.now():
                approved_ids.add(ov['record_id'])

    context = {
        'records': records,
        'q': query,
        'shift': shift,
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
    edit_id = (request.POST.get('edit_id') or request.GET.get('edit') or '').strip()
    edit_record = get_active_record_or_404(Dispatch, edit_id) if edit_id else None
    if edit_record and not ensure_edit_lock_allowed(request, 'dispatch', edit_record):
        return redirect('dispatch_records')

    if request.method == 'POST':
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
    }
    return render(request, 'dispatch_entry.html', context)


@login_required
@permission_required('can_approve_dispatch')
def dispatch_records(request):
    """Dispatch records list page"""
    query = (request.GET.get('q') or '').strip()

    records = Dispatch.objects.filter(is_active=True, job_card__is_active=True).select_related('job_card').order_by('-dispatch_date', '-id')
    if query:
        records = records.filter(
            Q(job_card__job_card_no__icontains=query) |
            Q(dc_no__icontains=query)
        )

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
        'q': query,
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

    requested_entity = (request.GET.get('entity') or '').strip()
    entity_type = requested_entity if requested_entity in accessible_entities else accessible_entities[0]
    query = (request.GET.get('q') or '').strip()
    config = AUDIT_CONFIG[entity_type]

    records = config['model'].objects.filter(is_active=False).order_by('-id')
    if entity_type == 'job_card':
        records = records.select_related('material', 'machine_name', 'department').order_by('-created_at')
        if query:
            records = records.filter(
                Q(job_card_no__icontains=query) |
                Q(SKU__icontains=query) |
                Q(PO_No__icontains=query)
            )
    elif entity_type == 'production':
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

    context = {
        'accessible_entities': accessible_entities,
        'entity_type': entity_type,
        'query': query,
        'records': records,
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
    if config is None or entity_type not in ('production', 'dispatch'):
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

    # Get date range (default to last 7 days)
    try:
        days = int(request.GET.get('days', 7))
    except (TypeError, ValueError):
        days = 7
    if days < 1:
        days = 1

    end_date = timezone.now().date()
    start_date = end_date - timedelta(days=days - 1)

    period_productions = Production.objects.filter(is_active=True, job_card__is_active=True, date__gte=start_date, date__lte=end_date)
    period_dispatches = Dispatch.objects.filter(is_active=True, job_card__is_active=True, dispatch_date__gte=start_date, dispatch_date__lte=end_date)

    # Calculate period metrics
    total_impressions = period_productions.aggregate(total=Sum('impressions'))['total'] or 0
    total_downtime = period_productions.aggregate(total=Sum('downtime'))['total'] or 0
    total_output = period_productions.aggregate(total=Sum('output_sheets'))['total'] or 0
    total_waste = period_productions.aggregate(total=Sum('waste_sheets'))['total'] or 0

    # Dispatch metrics for same period
    total_dispatch_qty = period_dispatches.aggregate(total=Sum('dispatch_qty'))['total'] or 0
    dispatch_count = period_dispatches.count()
    dispatched_job_cards_count = period_dispatches.values('job_card').distinct().count()
    avg_dispatch_qty = (total_dispatch_qty / dispatch_count) if dispatch_count else 0
    dispatch_fulfillment_pct = (total_dispatch_qty / total_output * 100) if total_output > 0 else 0

    # Calculate OEE (simplified - you may want to adjust based on your machine standards)
    available_time_minutes = period_productions.count() * 480  # Assuming 8 hours per production record
    actual_run_time = available_time_minutes - total_downtime

    availability = (actual_run_time / available_time_minutes * 100) if available_time_minutes > 0 else 0
    performance = 85.0  # Placeholder - would need machine standards
    quality = ((total_output / (total_output + total_waste)) * 100) if (total_output + total_waste) > 0 else 100

    oee_value = (availability * performance * quality) / 10000

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

    # Downtime analysis
    downtime_by_category = period_productions\
        .exclude(downtime_category__isnull=True)\
        .exclude(downtime_category='')\
        .values('downtime_category')\
        .annotate(
            total_minutes=Sum('downtime'),
            count=Count('id')
        )\
        .order_by('-total_minutes')[:5]

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

    downtime_by_category_data = [
        {
            'downtime_category': item['downtime_category'],
            'total_minutes': float(item['total_minutes'] or 0),
            'count': int(item['count'] or 0),
        }
        for item in downtime_by_category
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
        'period_label': 'Today' if days == 1 else f'Last {days} Days',
        'oee_value': round(oee_value, 2),
        'availability_value': round(availability, 2),
        'performance_value': round(performance, 2),
        'quality_value': round(quality, 2),
        'total_impressions': total_impressions,
        'total_output': total_output,
        'total_waste': total_waste,
        'total_downtime': total_downtime,
        'total_dispatch_qty': total_dispatch_qty,
        'dispatch_count': dispatch_count,
        'dispatched_job_cards_count': dispatched_job_cards_count,
        'avg_dispatch_qty': round(avg_dispatch_qty, 2),
        'dispatch_fulfillment_pct': round(dispatch_fulfillment_pct, 2),
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
    }

    return render(request, 'production_dashboard.html', context)


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