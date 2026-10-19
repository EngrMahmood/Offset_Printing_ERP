import logging
from datetime import date, timedelta
from decimal import Decimal, InvalidOperation

from django.contrib import messages
from django.core.exceptions import PermissionDenied
from django.db.models import Case, Count, IntegerField, Q, Sum, When
from django.http import HttpResponse
from django.db import transaction
from django.shortcuts import get_object_or_404, redirect, render
from django.utils import timezone

from .decorators import maintenance_required, maintenance_staff_required
from .reporting import build_report
from .permissions import STAFF_ROLES, is_maintenance_staff, require_record_access
from .workflow import (
    ACTION_LABELS, MACHINE_STATE_CHOICES, STATUS_HELP, TransitionError, allowed_transitions, transition_record,
)
from .forms import (
    ComplaintForm, MaintenanceAttachmentForm, MaintenanceServiceJobForm,
    MaintenanceSparePartForm, PreventiveMaintenancePlanForm, TriageForm,
)
from core.models import Production

from .models import (
    MachineDowntime, MaintenanceAttachment, MaintenanceRecord, PreventiveMaintenancePlan,
)
from .services import (
    generate_record_no, log_activity, notify_assigned, pm_overdue_since, pm_plan_status, notify_complaint_raised, notify_triaged,
    open_breakdown_downtime_if_needed, open_downtime, raise_service_demand, raise_spare_part_demand,
    soft_delete_record,
)
logger = logging.getLogger(__name__)

PRIORITY_RANK = Case(
    When(priority='CRITICAL', then=0),
    When(priority='MAJOR', then=1),
    When(priority='MEDIUM', then=2),
    When(priority='LOW', then=3),
    default=4,
    output_field=IntegerField(),
)


def staff_users():
    from django.contrib.auth import get_user_model

    User = get_user_model()
    return User.objects.filter(is_active=True, profile__role__in=STAFF_ROLES).select_related('profile')

@maintenance_required
def dashboard(request):
    is_staff = is_maintenance_staff(request.user)
    today = timezone.localdate()
    open_qs = MaintenanceRecord.objects.filter(is_active=True, status__in=MaintenanceRecord.OPEN_STATUSES)
    open_records = open_qs.select_related('machine').annotate(priority_rank=PRIORITY_RANK).order_by('priority_rank', 'created_at')
    machines_down = MachineDowntime.objects.filter(end_at__isnull=True).select_related('machine', 'record')
    my_tickets = (
        MaintenanceRecord.objects.filter(is_active=True, reported_by=request.user)
        .select_related('machine').order_by('-created_at')[:10]
    )
    context = {
        'is_staff': is_staff,
        'machines_down': machines_down,
        'my_tickets': [(record, STATUS_HELP.get(record.status, '')) for record in my_tickets],
        'my_complaints_count': open_qs.filter(reported_by=request.user).count(),
    }
    if is_staff:
        by_status = {row['status']: row['n'] for row in open_qs.values('status').annotate(n=Count('id'))}
        pm_open = open_qs.filter(pm_plan__isnull=False).select_related('machine', 'pm_plan')
        overdue_pm = [r for r in pm_open if (pm_overdue_since(r) or today) < today]
        context.update({
            'open_records': open_records[:20],
            'open_count': open_records.count(),
            'new_complaints_count': by_status.get('REPORTED', 0),
            'my_tickets_count': open_qs.filter(assigned_to=request.user).count(),
            'critical_records': open_records.filter(priority='CRITICAL')[:10],
            'in_progress_count': by_status.get('IN_PROGRESS', 0),
            'awaiting_parts_count': by_status.get('AWAITING_PARTS', 0),
            'awaiting_vendor_count': by_status.get('AWAITING_VENDOR', 0),
            'overdue_pm': overdue_pm,
            'overdue_triage': [r for r in open_qs.filter(status='REPORTED') if r.is_triage_overdue],
            'stale_in_progress': [r for r in open_qs.filter(status='IN_PROGRESS').select_related('machine') if r.is_stale_in_progress],
            'recent_completed': (
                MaintenanceRecord.objects.filter(is_active=True, status__in=('COMPLETED', 'VERIFIED', 'CLOSED'))
                .select_related('machine').order_by('-updated_at')[:5]
            ),
        })
    return render(request, 'maintenance/dashboard.html', context)


@maintenance_required
def record_list(request):
    is_staff = is_maintenance_staff(request.user)
    records = MaintenanceRecord.objects.filter(is_active=True).select_related('machine')
    if not is_staff:
        records = records.filter(reported_by=request.user)

    machine = request.GET.get('machine')
    status = request.GET.get('status')
    maintenance_type = request.GET.get('maintenance_type')
    priority = request.GET.get('priority')
    view = request.GET.get('view')
    q = request.GET.get('q', '').strip()
    assigned = request.GET.get('assigned_to')
    date_from = _parse_date(request.GET.get('date_from'))
    date_to = _parse_date(request.GET.get('date_to'))

    if machine:
        records = records.filter(machine_id=machine)
    if status:
        records = records.filter(status=status)
    if maintenance_type:
        records = records.filter(maintenance_type=maintenance_type)
    if priority:
        records = records.filter(priority=priority)
    if assigned and is_staff:
        records = records.filter(assigned_to_id=assigned)
    if date_from:
        records = records.filter(reported_date__gte=date_from)
    if date_to:
        records = records.filter(reported_date__lte=date_to)
    if view == 'mine':
        records = records.filter(reported_by=request.user)
    elif view == 'assigned' and is_staff:
        records = records.filter(assigned_to=request.user)
    if q:
        records = records.filter(
            Q(record_no__icontains=q) | Q(machine__name__icontains=q) | Q(fault_description__icontains=q)
        )

    if request.GET.get('export') == 'xlsx':
        if not is_staff:
            raise PermissionDenied
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
        'technicians': staff_users() if is_staff else [],
        'date_from': date_from.isoformat() if date_from else '',
        'date_to': date_to.isoformat() if date_to else '',
        'selected_technician': assigned or '',
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


DUPLICATE_SUBMIT_WINDOW = timedelta(minutes=2)


@maintenance_required
def complaint_create(request):
    """The shop-floor screen: pick a machine, describe the problem, say whether it
    has stopped. Nothing technical is asked - the engineer classifies it at triage."""
    duplicate_records = None
    if request.method == 'POST':
        form = ComplaintForm(request.POST, request.FILES)
        if form.is_valid():
            machine = form.cleaned_data['machine']
            description = form.cleaned_data['fault_description'].strip()

            # Double-click / refresh guard: the same person re-sending the same
            # complaint within moments gets the ticket they already created.
            recent = MaintenanceRecord.objects.filter(
                reported_by=request.user, machine=machine, fault_description=description, is_active=True,
                created_at__gte=timezone.now() - DUPLICATE_SUBMIT_WINDOW,
            ).first()
            if recent:
                messages.info(request, f'Complaint {recent.record_no} was already reported.')
                return redirect('maintenance:record_detail', pk=recent.pk)

            duplicate_records = MaintenanceRecord.objects.filter(
                machine=machine, is_active=True, status__in=MaintenanceRecord.OPEN_STATUSES,
            ).select_related('machine')
            if duplicate_records.exists() and request.POST.get('confirm_duplicate') != '1':
                pass  # fall through to re-render with the warning banner below
            else:
                stopped = form.cleaned_data['machine_stopped'] == 'yes'
                with transaction.atomic():
                    record = form.save(commit=False)
                    record.reported_by = request.user
                    record.reported_date = timezone.localdate()
                    record.status = 'REPORTED'
                    record.priority = 'CRITICAL' if stopped else 'MEDIUM'
                    record.machine_stopped = stopped
                    record.save()
                    form.save_m2m()
                    generate_record_no(record)
                    log_activity(record, request.user, 'Complaint raised', to_status=record.status)
                    if stopped:
                        open_downtime(record, reason='BREAKDOWN')
                    photo = form.cleaned_data.get('photo')
                    if photo:
                        MaintenanceAttachment.objects.create(record=record, file=photo, uploaded_by=request.user)
                try:
                    notify_complaint_raised(record, request.user)
                except Exception:
                    logger.exception('Complaint notification failed for %s', record.pk)
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
            notify_assigned(record, request.user)
            messages.success(request, f'{record.record_no} triaged and moved to Diagnosed.')
            return redirect('maintenance:record_detail', pk=record.pk)
    else:
        initial = {} if record.assigned_to_id else {'assigned_to': request.user}
        form = TriageForm(instance=record, initial=initial)
    return render(request, 'maintenance/record_triage.html', {'form': form, 'record': record})


@maintenance_required
def record_comment(request, pk):
    record = get_object_or_404(MaintenanceRecord, pk=pk, is_active=True)
    require_record_access(request.user, record)
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
    record = get_object_or_404(
        MaintenanceRecord.objects.select_related('machine', 'assigned_to', 'verified_by'), pk=pk, is_active=True,
    )
    require_record_access(request.user, record)
    can_work = is_maintenance_staff(request.user)
    needs_triage = record.status == 'REPORTED'
    # One authoritative rule set (workflow.py); DIAGNOSED comes only from triage.
    next_statuses = allowed_transitions(record, request.user)
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
        'checklist_items': record.checklist_items.select_related('done_by'),
        'next_statuses': next_statuses,
        'next_actions': [(status, ACTION_LABELS.get(status, status)) for status in next_statuses],
        'status_help': STATUS_HELP.get(record.status, ''),
        'machine_state_choices': MACHINE_STATE_CHOICES,
        'machine_open_downtime': record.machine_stopped and record.downtime_intervals.filter(end_at__isnull=True).exists(),
        'test_result_choices': MaintenanceRecord.TEST_RESULT_CHOICES,
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
            notify_assigned(record, request.user)
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
    record = get_object_or_404(MaintenanceRecord, pk=pk, is_active=True)
    require_record_access(request.user, record)
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
    """Create Supply Chain requests for lines that have none yet.

    Row-locked inside one transaction: a second click (or a retry) waits, then
    finds the lines already linked and creates nothing, so a part is never
    requested twice.
    """
    record = get_object_or_404(MaintenanceRecord, pk=pk, is_active=True)
    if request.method == 'POST':
        raised = 0
        with transaction.atomic():
            for spare_part in record.spare_parts.select_for_update().filter(item_request__isnull=True):
                raise_spare_part_demand(spare_part, request.user)
                raised += 1
            for service_job in record.service_jobs.select_for_update().filter(item_request__isnull=True):
                raise_service_demand(service_job, request.user)
                raised += 1
        if raised:
            messages.success(request, f'{raised} item request(s) raised in Supply Chain.')
        else:
            messages.info(request, 'Nothing new to request - every line already has a Supply Chain request.')
    return redirect('maintenance:record_detail', pk=record.pk)


@maintenance_staff_required
def record_transition(request, pk):
    record = get_object_or_404(MaintenanceRecord, pk=pk, is_active=True)
    if request.method == 'POST':
        try:
            record = transition_record(
                record,
                request.POST.get('to_status', ''),
                request.user,
                note=request.POST.get('note', ''),
                labour_hours=request.POST.get('labour_hours', '').strip() or None,
                machine_state=request.POST.get('machine_state', ''),
            )
        except TransitionError as exc:
            messages.error(request, str(exc))
        else:
            messages.success(request, f'Ticket is now {record.get_status_display()}.')
    return redirect('maintenance:record_detail', pk=record.pk)


@maintenance_staff_required
def record_checklist_update(request, pk):
    """Tick off PM checklist items and record findings against this ticket."""
    record = get_object_or_404(MaintenanceRecord, pk=pk, is_active=True)
    if request.method == 'POST':
        now = timezone.now()
        changed = 0
        for item in record.checklist_items.all():
            done = request.POST.get(f'done_{item.pk}') == 'on'
            finding = request.POST.get(f'finding_{item.pk}', '').strip()
            if done != item.is_done or finding != item.finding:
                item.is_done = done
                item.finding = finding
                if done and not item.done_at:
                    item.done_by, item.done_at = request.user, now
                if not done:
                    item.done_by, item.done_at = None, None
                item.save()
                changed += 1
        if changed:
            log_activity(record, request.user, 'Checklist updated', note=f'{changed} item(s) changed')
        messages.success(request, 'Checklist saved.')
    return redirect('maintenance:record_detail', pk=record.pk)


@maintenance_staff_required
def record_repair_update(request, pk):
    """Diagnosis, repair action, labour and machine-test result for the ticket."""
    record = get_object_or_404(MaintenanceRecord, pk=pk, is_active=True)
    if request.method == 'POST':
        test_result = request.POST.get('test_result', '')
        if test_result not in dict(MaintenanceRecord.TEST_RESULT_CHOICES):
            messages.error(request, 'Choose a valid machine test result.')
            return redirect('maintenance:record_detail', pk=record.pk)
        labour_raw = request.POST.get('labour_hours', '').strip()
        labour = record.labour_hours
        if labour_raw:
            try:
                labour = Decimal(labour_raw)
            except InvalidOperation:
                messages.error(request, 'Labour hours must be a number.')
                return redirect('maintenance:record_detail', pk=record.pk)
            if labour < 0:
                messages.error(request, 'Labour hours cannot be negative.')
                return redirect('maintenance:record_detail', pk=record.pk)
        record.root_cause = request.POST.get('root_cause', '').strip()
        record.repair_action = request.POST.get('repair_action', '').strip()
        record.test_result = test_result
        record.labour_hours = labour
        record.save(update_fields=['root_cause', 'repair_action', 'test_result', 'labour_hours', 'updated_at'])
        log_activity(record, request.user, 'Repair details updated', note=f'Test result: {record.get_test_result_display()}')
        messages.success(request, 'Repair details saved.')
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


@maintenance_staff_required
def downtime_list(request):
    intervals = MachineDowntime.objects.select_related('machine', 'record').all()
    return render(request, 'maintenance/downtime_list.html', {'intervals': intervals})


@maintenance_staff_required
def pm_plan_list(request):
    plans = list(PreventiveMaintenancePlan.objects.select_related('machine', 'assigned_to'))
    machine_ids = {plan.machine_id for plan in plans if plan.interval_type == 'IMPRESSIONS'}
    totals = {
        row['machine_id']: row['total'] or 0
        for row in Production.objects.filter(machine_id__in=machine_ids, is_active=True)
        .values('machine_id').annotate(total=Sum('impressions'))
    }
    rank = {'overdue': 0, 'soon': 1, 'upcoming': 2, 'unset': 3, 'inactive': 4}
    rows = []
    for plan in plans:
        state, label = pm_plan_status(plan, impression_total=totals.get(plan.machine_id, 0))
        rows.append({'plan': plan, 'state': state, 'label': label, 'impressions_run': totals.get(plan.machine_id)})
    rows.sort(key=lambda row: (rank[row['state']], row['plan'].machine.name))
    return render(request, 'maintenance/pm_plan_list.html', {'rows': rows})


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


@maintenance_staff_required
def pm_plan_edit(request, pk):
    plan = get_object_or_404(PreventiveMaintenancePlan, pk=pk)
    if request.method == 'POST':
        form = PreventiveMaintenancePlanForm(request.POST, instance=plan)
        if form.is_valid():
            form.save()
            messages.success(request, 'Preventive maintenance plan updated.')
            return redirect('maintenance:pm_plan_list')
    else:
        form = PreventiveMaintenancePlanForm(instance=plan)
    return render(request, 'maintenance/pm_plan_form.html', {'form': form, 'is_create': False, 'plan': plan})


@maintenance_staff_required
def pm_plan_toggle(request, pk):
    plan = get_object_or_404(PreventiveMaintenancePlan, pk=pk)
    if request.method == 'POST':
        plan.is_active = not plan.is_active
        plan.save(update_fields=['is_active'])
        messages.success(request, f'Plan "{plan.title}" {"enabled" if plan.is_active else "disabled"}.')
    return redirect('maintenance:pm_plan_list')


def _parse_date(value):
    try:
        return date.fromisoformat((value or '').strip())
    except ValueError:
        return None


@maintenance_staff_required
def reports(request):
    today = timezone.localdate()
    date_to = _parse_date(request.GET.get('date_to')) or today
    date_from = _parse_date(request.GET.get('date_from')) or (date_to - timedelta(days=90))
    if date_from > date_to:
        date_from, date_to = date_to, date_from
    context = build_report(date_from, date_to)
    context.update({'date_from': date_from.isoformat(), 'date_to': date_to.isoformat()})
    return render(request, 'maintenance/reports.html', context)
