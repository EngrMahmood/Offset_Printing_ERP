from django.contrib import messages
from django.core.exceptions import PermissionDenied
from django.db.models import Avg, Sum
from django.http import HttpResponse
from django.shortcuts import get_object_or_404, redirect, render
from django.utils import timezone

from core.navigation import get_nav_permissions

from .decorators import maintenance_manager_required, maintenance_required
from .forms import (
    MaintenanceAttachmentForm, MaintenanceRecordForm, MaintenanceServiceJobForm, MaintenanceSparePartForm,
    PreventiveMaintenancePlanForm,
)
from .models import (
    MachineDowntime, MaintenanceRecord, PreventiveMaintenancePlan,
)
from .services import (
    generate_record_no, log_activity, raise_service_demand, raise_spare_part_demand, review_record,
    soft_delete_record, submit_for_approval,
)

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
    open_records = (
        MaintenanceRecord.objects.filter(is_active=True, status__in=MaintenanceRecord.OPEN_STATUSES)
        .select_related('machine')
    )
    machines_down = MachineDowntime.objects.filter(end_at__isnull=True).select_related('machine')
    pending_approval_count = open_records.filter(status='PENDING_APPROVAL').count()
    context = {
        'open_records': open_records[:20],
        'open_count': open_records.count(),
        'pending_approval_count': pending_approval_count,
        'machines_down': machines_down,
    }
    return render(request, 'maintenance/dashboard.html', context)


@maintenance_required
def record_list(request):
    records = MaintenanceRecord.objects.filter(is_active=True).select_related('machine')

    machine = request.GET.get('machine')
    status = request.GET.get('status')
    maintenance_type = request.GET.get('maintenance_type')
    priority = request.GET.get('priority')

    if machine:
        records = records.filter(machine_id=machine)
    if status:
        records = records.filter(status=status)
    if maintenance_type:
        records = records.filter(maintenance_type=maintenance_type)
    if priority:
        records = records.filter(priority=priority)

    if request.GET.get('export') == 'xlsx':
        return _export_records_xlsx(records)

    context = {
        'records': records,
        'status_choices': MaintenanceRecord.STATUS_CHOICES,
        'type_choices': MaintenanceRecord.MAINTENANCE_TYPE_CHOICES,
        'priority_choices': MaintenanceRecord.PRIORITY_CHOICES,
        'is_superuser': request.user.is_superuser,
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
def record_create(request):
    if request.method == 'POST':
        form = MaintenanceRecordForm(request.POST)
        if form.is_valid():
            record = form.save(commit=False)
            record.reported_by = request.user
            record.status = 'PENDING_APPROVAL'
            record.save()
            form.save_m2m()
            generate_record_no(record)
            log_activity(record, request.user, 'Record created', to_status=record.status)
            submit_for_approval(record, request.user)
            if record.maintenance_type == 'BREAKDOWN':
                MachineDowntime.objects.create(
                    machine=record.machine, record=record, start_at=timezone.now(), reason='BREAKDOWN',
                )
            messages.success(request, f'Maintenance record {record.record_no} created and sent for manager approval.')
            return redirect('maintenance:record_detail', pk=record.pk)
    else:
        form = MaintenanceRecordForm()
    return render(request, 'maintenance/record_form.html', {'form': form, 'is_create': True})


@maintenance_required
def record_detail(request, pk):
    record = get_object_or_404(MaintenanceRecord.objects.select_related('machine'), pk=pk)
    nav = get_nav_permissions(request)
    can_review = record.status == 'PENDING_APPROVAL' and nav.get('role') in ('admin', 'manager')
    context = {
        'record': record,
        'spare_parts': record.spare_parts.select_related('item_request', 'existing_sku'),
        'service_jobs': record.service_jobs.select_related('item_request'),
        'downtime_intervals': record.downtime_intervals.all(),
        'activity_log': record.activity_log.select_related('actor'),
        'approvals': record.approvals.select_related('actor'),
        'attachments': record.attachments.select_related('uploaded_by'),
        'spare_part_form': MaintenanceSparePartForm(),
        'service_job_form': MaintenanceServiceJobForm(),
        'attachment_form': MaintenanceAttachmentForm(),
        'next_statuses': sorted(VALID_TRANSITIONS.get(record.status, set())),
        'can_review': can_review,
        'can_delete': request.user.is_superuser,
    }
    return render(request, 'maintenance/record_detail.html', context)


@maintenance_required
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


@maintenance_required
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


@maintenance_required
def raise_demand(request, pk):
    record = get_object_or_404(MaintenanceRecord, pk=pk)
    if request.method == 'POST':
        for spare_part in record.spare_parts.filter(item_request__isnull=True):
            raise_spare_part_demand(spare_part, request.user)
        for service_job in record.service_jobs.filter(item_request__isnull=True):
            raise_service_demand(service_job, request.user)
        messages.success(request, 'Item request(s) raised for pending spare parts / service jobs.')
    return redirect('maintenance:record_detail', pk=record.pk)


@maintenance_required
def record_transition(request, pk):
    record = get_object_or_404(MaintenanceRecord, pk=pk)
    if request.method == 'POST':
        to_status = request.POST.get('to_status')
        note = request.POST.get('note', '')
        allowed = VALID_TRANSITIONS.get(record.status, set())
        if to_status not in allowed:
            messages.error(request, f'Cannot move from {record.status} to {to_status}.')
        else:
            from_status = record.status
            record.status = to_status
            if to_status == 'IN_PROGRESS' and not record.work_start_at:
                record.work_start_at = timezone.now()
            if to_status == 'COMPLETED':
                record.work_end_at = timezone.now()
                open_downtime = record.downtime_intervals.filter(end_at__isnull=True).first()
                if open_downtime:
                    open_downtime.end_at = timezone.now()
                    open_downtime.save(update_fields=['end_at'])
            record.save()
            log_activity(record, request.user, 'Status changed', from_status=from_status, to_status=to_status, note=note)
            messages.success(request, f'Record moved to {record.get_status_display()}.')
    return redirect('maintenance:record_detail', pk=record.pk)


@maintenance_manager_required
def approval_queue(request):
    records = (
        MaintenanceRecord.objects.filter(is_active=True, status='PENDING_APPROVAL')
        .select_related('machine', 'reported_by')
    )
    return render(request, 'maintenance/approval_queue.html', {'records': records})


@maintenance_manager_required
def record_review(request, pk):
    record = get_object_or_404(MaintenanceRecord, pk=pk)
    if request.method == 'POST':
        action = request.POST.get('action')
        comment = request.POST.get('comment', '')
        if action not in ('APPROVE', 'REJECT'):
            messages.error(request, 'Invalid review action.')
        else:
            try:
                review_record(record, request.user, action, comment)
                messages.success(request, f'Record {record.record_no} {action.lower()}d.')
            except ValueError as exc:
                messages.error(request, str(exc))
    return redirect('maintenance:approval_queue')


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


@maintenance_required
def pm_plan_list(request):
    plans = PreventiveMaintenancePlan.objects.select_related('machine').all()
    return render(request, 'maintenance/pm_plan_list.html', {'plans': plans})


@maintenance_required
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
