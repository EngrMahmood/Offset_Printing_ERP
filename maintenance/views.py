from django.contrib import messages
from django.core.exceptions import PermissionDenied
from django.db.models import Avg, Case, IntegerField, Q, Sum, When
from django.http import HttpResponse
from django.shortcuts import get_object_or_404, redirect, render
from django.utils import timezone

from core.navigation import get_nav_permissions

from .decorators import maintenance_required, maintenance_staff_required
from .forms import (
    ComplaintForm, MaintenanceAttachmentForm, MaintenanceServiceJobForm,
    MaintenanceSparePartForm, PreventiveMaintenancePlanForm, TriageForm,
)
from .models import (
    MachineDowntime, MaintenanceAttachment, MaintenanceRecord, PreventiveMaintenancePlan,
)
from .services import (
    generate_record_no, log_activity, notify_complaint_raised, notify_status_update, notify_triaged,
    open_breakdown_downtime_if_needed, open_downtime, raise_service_demand, raise_spare_part_demand,
    soft_delete_record,
)

PRIORITY_RANK = Case(
    When(priority='CRITICAL', then=0),
    When(priority='MAJOR', then=1),
    When(priority='MEDIUM', then=2),
    When(priority='LOW', then=3),
    default=4,
    output_field=IntegerField(),
)

STAFF_ROLES = ('admin', 'manager', 'production_manager', 'maintenance_engineer')


def staff_users():
    from django.contrib.auth import get_user_model

    User = get_user_model()
    return User.objects.filter(is_active=True, profile__role__in=STAFF_ROLES).select_related('profile')

VALID_TRANSITIONS = {
    'REPORTED': {'DIAGNOSED', 'CANCELLED'},
    'DIAGNOSED': {'AWAITING_PARTS', 'AWAITING_VENDOR', 'IN_PROGRESS', 'CANCELLED'},
    'AWAITING_PARTS': {'IN_PROGRESS', 'CANCELLED'},
    'AWAITING_VENDOR': {'IN_PROGRESS', 'CANCELLED'},
    'IN_PROGRESS': {'COMPLETED', 'CANCELLED'},
    'COMPLETED': {'VERIFIED'},
    'VERIFIED': {'CLOSED'},
}


@maintenance_required
def dashboard(request):
    nav = get_nav_permissions(request)
    is_staff = nav.get('role') in STAFF_ROLES or request.user.is_superuser
    open_records = (
        MaintenanceRecord.objects.filter(is_active=True, status__in=MaintenanceRecord.OPEN_STATUSES)
        .select_related('machine')
        .annotate(priority_rank=PRIORITY_RANK)
        .order_by('priority_rank', 'created_at')
    )
    machines_down = MachineDowntime.objects.filter(end_at__isnull=True).select_related('machine')
    context = {
        'open_records': open_records[:20],
        'open_count': open_records.count(),
        'machines_down': machines_down,
        'is_staff': is_staff,
        'new_complaints_count': open_records.filter(status='REPORTED').count() if is_staff else 0,
        'my_complaints_count': open_records.filter(reported_by=request.user).count(),
        'my_tickets_count': open_records.filter(assigned_to=request.user).count() if is_staff else 0,
    }
    return render(request, 'maintenance/dashboard.html', context)


@maintenance_required
def record_list(request):
    nav = get_nav_permissions(request)
    is_staff = nav.get('role') in STAFF_ROLES or request.user.is_superuser
    records = MaintenanceRecord.objects.filter(is_active=True).select_related('machine')

    machine = request.GET.get('machine')
    status = request.GET.get('status')
    maintenance_type = request.GET.get('maintenance_type')
    priority = request.GET.get('priority')
    view = request.GET.get('view')
    q = request.GET.get('q', '').strip()

    if machine:
        records = records.filter(machine_id=machine)
    if status:
        records = records.filter(status=status)
    if maintenance_type:
        records = records.filter(maintenance_type=maintenance_type)
    if priority:
        records = records.filter(priority=priority)
    if view == 'mine':
        records = records.filter(reported_by=request.user)
    elif view == 'assigned' and is_staff:
        records = records.filter(assigned_to=request.user)
    if q:
        records = records.filter(
            Q(record_no__icontains=q) | Q(machine__name__icontains=q) | Q(fault_description__icontains=q)
        )

    if request.GET.get('export') == 'xlsx':
        return _export_records_xlsx(records)

    context = {
        'records': records,
        'status_choices': MaintenanceRecord.STATUS_CHOICES,
        'type_choices': MaintenanceRecord.MAINTENANCE_TYPE_CHOICES,
        'priority_choices': MaintenanceRecord.PRIORITY_CHOICES,
        'is_superuser': request.user.is_superuser,
        'is_staff': is_staff,
        'view': view or '',
        'q': q,
    }
    return render(request, 'maintenance/record_list.html', context)


def _export_records_xlsx(records):
    import openpyxl

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = 'Maintenance Records'
    sheet.append([
        'Record No.', 'Machine', 'Reported Date', 'Type', 'Execution', 'Priority', 'Status',
        'Fault Description', 'Assigned To', 'Labour Hours', 'Total Cost',
    ])
    for record in records:
        sheet.append([
            record.record_no, str(record.machine), str(record.reported_date), record.get_maintenance_type_display(),
            record.get_execution_type_display(), record.get_priority_display(), record.get_status_display(),
            record.fault_description, str(record.assigned_to or ''), float(record.labour_hours),
            float(record.total_cost),
        ])

    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = 'attachment; filename=maintenance_records.xlsx'
    workbook.save(response)
    return response


@maintenance_required
def complaint_create(request):
    """The shop-floor screen: pick a machine, describe the problem. Nothing
    technical is asked — the engineer classifies it at triage."""
    duplicate_records = None
    if request.method == 'POST':
        form = ComplaintForm(request.POST, request.FILES)
        if form.is_valid():
            machine = form.cleaned_data['machine']
            duplicate_records = MaintenanceRecord.objects.filter(
                machine=machine, is_active=True, status__in=MaintenanceRecord.OPEN_STATUSES,
            ).select_related('machine')
            if duplicate_records.exists() and request.POST.get('confirm_duplicate') != '1':
                pass  # fall through to re-render with the warning banner below
            else:
                stopped = form.cleaned_data['machine_stopped'] == 'no'
                record = form.save(commit=False)
                record.reported_by = request.user
                record.reported_date = timezone.now().date()
                record.status = 'REPORTED'
                record.priority = 'CRITICAL' if stopped else 'MEDIUM'
                record.save()
                generate_record_no(record)
                log_activity(record, request.user, 'Complaint raised', to_status=record.status)
                if stopped:
                    open_downtime(record, reason='BREAKDOWN')
                photo = form.cleaned_data.get('photo')
                if photo:
                    MaintenanceAttachment.objects.create(record=record, file=photo, uploaded_by=request.user)
                notify_complaint_raised(record, request.user)
                messages.success(request, f'Complaint {record.record_no} reported. Maintenance has been notified.')
                return redirect('maintenance:record_detail', pk=record.pk)
    else:
        form = ComplaintForm()
    return render(request, 'maintenance/complaint_form.html', {'form': form, 'duplicate_records': duplicate_records})


@maintenance_staff_required
def record_triage(request, pk):
    """The engineer's classification screen — turns a raw REPORTED complaint
    into a properly assessed record and moves it to DIAGNOSED."""
    record = get_object_or_404(MaintenanceRecord, pk=pk)
    if record.status != 'REPORTED':
        messages.error(request, 'This record has already been triaged.')
        return redirect('maintenance:record_detail', pk=record.pk)

    if request.method == 'POST':
        form = TriageForm(request.POST, instance=record)
        if form.is_valid():
            record = form.save(commit=False)
            record.status = 'DIAGNOSED'
            record.save()
            form.save_m2m()
            open_breakdown_downtime_if_needed(record)
            log_activity(
                record, request.user, 'Triaged & classified', from_status='REPORTED', to_status='DIAGNOSED',
            )
            notify_triaged(record, request.user)
            messages.success(request, f'{record.record_no} triaged and moved to Diagnosed.')
            return redirect('maintenance:record_detail', pk=record.pk)
    else:
        initial = {} if record.assigned_to_id else {'assigned_to': request.user}
        form = TriageForm(instance=record, initial=initial)
    return render(request, 'maintenance/record_triage.html', {'form': form, 'record': record})


@maintenance_required
def record_comment(request, pk):
    record = get_object_or_404(MaintenanceRecord, pk=pk)
    if request.method == 'POST':
        note = request.POST.get('note', '').strip()
        if note:
            log_activity(record, request.user, 'Comment', note=note)
            messages.success(request, 'Comment added.')
        else:
            messages.error(request, 'Comment cannot be empty.')
    return redirect('maintenance:record_detail', pk=record.pk)


@maintenance_required
def record_detail(request, pk):
    record = get_object_or_404(MaintenanceRecord.objects.select_related('machine', 'assigned_to'), pk=pk)
    nav = get_nav_permissions(request)
    can_work = nav.get('role') in STAFF_ROLES or request.user.is_superuser
    needs_triage = record.status == 'REPORTED'
    # DIAGNOSED is only reached through the dedicated triage screen, so it's
    # excluded from the generic status-move dropdown; CANCELLED stays available.
    next_statuses = sorted(VALID_TRANSITIONS.get(record.status, set()) - {'DIAGNOSED'})
    can_reassign = can_work and record.status not in ('CLOSED', 'CANCELLED', 'REJECTED') and not needs_triage
    can_raise_demand = can_work and (record.spare_parts_needed or record.execution_type == 'OUTSOURCE')
    context = {
        'record': record,
        'spare_parts': record.spare_parts.select_related('item_request', 'existing_sku'),
        'service_jobs': record.service_jobs.select_related('item_request'),
        'downtime_intervals': record.downtime_intervals.all(),
        'activity_log': record.activity_log.select_related('actor'),
        'attachments': record.attachments.select_related('uploaded_by'),
        'spare_part_form': MaintenanceSparePartForm(),
        'service_job_form': MaintenanceServiceJobForm(),
        'attachment_form': MaintenanceAttachmentForm(),
        'next_statuses': next_statuses,
        'can_work': can_work,
        'can_reassign': can_reassign,
        'can_raise_demand': can_raise_demand,
        'engineers': staff_users() if can_reassign else None,
        'needs_triage': needs_triage,
        'can_delete': request.user.is_superuser,
    }
    return render(request, 'maintenance/record_detail.html', context)


@maintenance_staff_required
def record_reassign(request, pk):
    record = get_object_or_404(MaintenanceRecord, pk=pk)
    if request.method == 'POST':
        engineer_id = request.POST.get('assigned_to')
        engineer = staff_users().filter(pk=engineer_id).first() if engineer_id else None
        if not engineer:
            messages.error(request, 'Pick a valid engineer to assign.')
        else:
            record.assigned_to = engineer
            record.save(update_fields=['assigned_to'])
            log_activity(record, request.user, 'Reassigned', note=f'Assigned to {engineer}')
            messages.success(request, f'{record.record_no} assigned to {engineer}.')
    return redirect('maintenance:record_detail', pk=record.pk)


@maintenance_staff_required
def spare_part_add(request, pk):
    record = get_object_or_404(MaintenanceRecord, pk=pk)
    if request.method == 'POST':
        form = MaintenanceSparePartForm(request.POST)
        if form.is_valid():
            spare_part = form.save(commit=False)
            spare_part.record = record
            spare_part.save()
            messages.success(request, 'Spare part line added.')
        else:
            messages.error(request, 'Could not add spare part line.')
    return redirect('maintenance:record_detail', pk=record.pk)


@maintenance_staff_required
def service_job_add(request, pk):
    record = get_object_or_404(MaintenanceRecord, pk=pk)
    if request.method == 'POST':
        form = MaintenanceServiceJobForm(request.POST)
        if form.is_valid():
            service_job = form.save(commit=False)
            service_job.record = record
            service_job.save()
            messages.success(request, 'Service job added.')
        else:
            messages.error(request, 'Could not add service job.')
    return redirect('maintenance:record_detail', pk=record.pk)


@maintenance_required
def attachment_add(request, pk):
    record = get_object_or_404(MaintenanceRecord, pk=pk)
    if request.method == 'POST':
        form = MaintenanceAttachmentForm(request.POST, request.FILES)
        if form.is_valid():
            attachment = form.save(commit=False)
            attachment.record = record
            attachment.uploaded_by = request.user
            attachment.save()
            messages.success(request, 'Attachment uploaded.')
        else:
            messages.error(request, 'Could not upload attachment.')
    return redirect('maintenance:record_detail', pk=record.pk)


@maintenance_staff_required
def raise_demand(request, pk):
    record = get_object_or_404(MaintenanceRecord, pk=pk)
    if request.method == 'POST':
        for spare_part in record.spare_parts.filter(item_request__isnull=True):
            raise_spare_part_demand(spare_part, request.user)
        for service_job in record.service_jobs.filter(item_request__isnull=True):
            raise_service_demand(service_job, request.user)
        messages.success(request, 'Item request(s) raised for pending spare parts / service jobs.')
    return redirect('maintenance:record_detail', pk=record.pk)


@maintenance_staff_required
def record_transition(request, pk):
    record = get_object_or_404(MaintenanceRecord, pk=pk)
    if request.method == 'POST':
        to_status = request.POST.get('to_status')
        note = request.POST.get('note', '')
        allowed = VALID_TRANSITIONS.get(record.status, set()) - {'DIAGNOSED'}
        if to_status not in allowed:
            messages.error(request, f'Cannot move from {record.status} to {to_status}.')
        else:
            from_status = record.status
            record.status = to_status
            if to_status == 'IN_PROGRESS' and not record.work_start_at:
                record.work_start_at = timezone.now()
            if to_status == 'COMPLETED':
                record.work_end_at = timezone.now()
                labour_hours = request.POST.get('labour_hours', '').strip()
                if labour_hours:
                    try:
                        record.labour_hours = float(labour_hours)
                    except ValueError:
                        pass
                open_downtime = record.downtime_intervals.filter(end_at__isnull=True).first()
                if open_downtime:
                    open_downtime.end_at = timezone.now()
                    open_downtime.save(update_fields=['end_at'])
            record.save()
            log_activity(record, request.user, 'Status changed', from_status=from_status, to_status=to_status, note=note)
            if to_status in ('COMPLETED', 'VERIFIED', 'CLOSED'):
                notify_status_update(record, request.user)
            messages.success(request, f'Record moved to {record.get_status_display()}.')
    return redirect('maintenance:record_detail', pk=record.pk)


@maintenance_required
def bulk_delete(request):
    if not request.user.is_superuser:
        raise PermissionDenied
    if request.method == 'POST':
        selected_ids = request.POST.getlist('selected_ids')
        reason = request.POST.get('reason', '')
        records = MaintenanceRecord.objects.filter(pk__in=selected_ids, is_active=True)
        count = 0
        for record in records:
            soft_delete_record(record, request.user, reason)
            count += 1
        messages.success(request, f'Deleted {count} maintenance record(s).')
    return redirect('maintenance:record_list')


@maintenance_required
def downtime_list(request):
    intervals = MachineDowntime.objects.select_related('machine', 'record').all()
    return render(request, 'maintenance/downtime_list.html', {'intervals': intervals})


@maintenance_staff_required
def pm_plan_list(request):
    plans = PreventiveMaintenancePlan.objects.select_related('machine').all()
    return render(request, 'maintenance/pm_plan_list.html', {'plans': plans})


@maintenance_staff_required
def pm_plan_create(request):
    if request.method == 'POST':
        form = PreventiveMaintenancePlanForm(request.POST)
        if form.is_valid():
            form.save()
            messages.success(request, 'Preventive maintenance plan created.')
            return redirect('maintenance:pm_plan_list')
    else:
        form = PreventiveMaintenancePlanForm()
    return render(request, 'maintenance/pm_plan_form.html', {'form': form, 'is_create': True})


@maintenance_required
def reports(request):
    active_records = MaintenanceRecord.objects.filter(is_active=True)
    closed_intervals = MachineDowntime.objects.filter(end_at__isnull=False)

    by_machine = []
    for machine_id, machine_name in (
        MachineDowntime.objects.values_list('machine_id', 'machine__name').distinct()
    ):
        machine_intervals = closed_intervals.filter(machine_id=machine_id)
        breakdown_count = machine_intervals.filter(reason='BREAKDOWN').count()
        avg_repair_minutes = None
        if breakdown_count:
            durations = [
                (interval.end_at - interval.start_at).total_seconds() / 60
                for interval in machine_intervals.filter(reason='BREAKDOWN')
            ]
            avg_repair_minutes = sum(durations) / len(durations)
        machine_records = active_records.filter(machine_id=machine_id)
        total_cost = sum((r.total_cost for r in machine_records), 0)
        by_machine.append({
            'machine': machine_name,
            'breakdown_count': breakdown_count,
            'mttr_minutes': avg_repair_minutes,
            'total_cost': total_cost,
        })

    context = {
        'by_machine': sorted(by_machine, key=lambda x: x['breakdown_count'], reverse=True),
        'open_count': active_records.filter(status__in=MaintenanceRecord.OPEN_STATUSES).count(),
        'closed_count': active_records.filter(status='CLOSED').count(),
    }
    return render(request, 'maintenance/reports.html', context)
