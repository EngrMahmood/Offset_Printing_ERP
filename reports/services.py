from __future__ import annotations

from datetime import timedelta

from django.db.models import Avg, Count, Q, Sum
from django.utils import timezone

from core.models import ChangeLog, Dispatch, JobCard, Machine, Production, ProductionDowntime
from planning.models import PlanningJob, PoDocument, SkuRecipe


REPORT_CATALOG = [
    {
        'key': 'machine-planning',
        'title': 'Machine Planning',
        'description': 'Planned load by machine, current backlog, and machine capacity vs actual output.',
        'focus': 'Planning',
    },
    {
        'key': 'job-planning',
        'title': 'Job Planning',
        'description': 'Job status mix, due-date pressure, repeat mix, and planning readiness.',
        'focus': 'Planning',
    },
    {
        'key': 'plates-planning',
        'title': 'Plates Planning',
        'description': 'Plate set readiness, blocked jobs, and QC gate visibility.',
        'focus': 'Planning',
    },
    {
        'key': 'production-insights',
        'title': 'Production Insights',
        'description': 'Output, waste, downtime, OEE proxies, and machine/shift efficiency.',
        'focus': 'Execution',
    },
    {
        'key': 'qc-approvals',
        'title': 'QC Approvals',
        'description': 'Approval queue status, rejection trends, and QC turnaround time.',
        'focus': 'Quality',
    },
    {
        'key': 'dispatch-tracking',
        'title': 'Dispatch Tracking',
        'description': 'Fulfillment rates, dispatch completion, and delivery backlog.',
        'focus': 'Execution',
    },
    {
        'key': 'raw-material-cutting-request',
        'title': 'Raw Material Cutting Request',
        'description': 'Material cutting requirements for released jobs, including sheet sizes and quantities.',
        'focus': 'Execution',
    },
]


def _get_year(request):
    """Parse year from request, default to current year."""
    year_value = request.GET.get('year', '')
    current_year = timezone.now().year
    try:
        year = int(year_value) if year_value else current_year
        # Limit to reasonable range (2020-2030)
        if 2020 <= year <= 2030:
            return year
    except (ValueError, TypeError):
        pass
    return current_year


def _parse_days(request, default_days=30):
    raw_days = request.GET.get('days', default_days)
    try:
        days = int(raw_days)
    except (TypeError, ValueError):
        days = default_days
    return max(7, min(days, 365))


def _date_window(request, default_days=30):
    days = _parse_days(request, default_days=default_days)
    end = timezone.localdate()
    start = end - timedelta(days=days - 1)
    return start, end, days


def _search_jobs(queryset, search_value):
    if not search_value:
        return queryset
    return queryset.filter(
        Q(jc_number__icontains=search_value)
        | Q(po_number__icontains=search_value)
        | Q(sku__icontains=search_value)
        | Q(job_name__icontains=search_value)
        | Q(machine_name__icontains=search_value)
        | Q(department__icontains=search_value)
    )


def _status_count_map(rows, key='status'):
    return {row[key]: row['count'] for row in rows}


def build_overview_context(request):
    year = _get_year(request)
    
    planning_jobs = PlanningJob.objects.filter(is_active=True, plan_date__year=year)
    job_cards = JobCard.objects.filter(is_active=True, created_at__year=year)
    productions = Production.objects.filter(is_active=True, date__year=year)
    dispatches = Dispatch.objects.filter(is_active=True, dispatch_date__year=year)
    sku_recipes = SkuRecipe.objects.filter(is_active=True)
    po_documents = PoDocument.objects.filter(created_at__year=year)

    planning_status_rows = planning_jobs.values('status').annotate(count=Count('id')).order_by('status')
    job_card_status_rows = job_cards.values('status').annotate(count=Count('id')).order_by('status')
    recipe_status_rows = sku_recipes.values('master_data_status').annotate(count=Count('id')).order_by('master_data_status')

    module_cards = [
        {
            'title': 'Planning Jobs',
            'value': planning_jobs.count(),
            'subtitle': 'Active jobs in planning',
            'hint': f"{_status_count_map(planning_status_rows).get('pending_qc', 0)} waiting for QC",
        },
        {
            'title': 'Job Cards',
            'value': job_cards.count(),
            'subtitle': 'Execution records',
            'hint': f"{_status_count_map(job_card_status_rows).get('released', 0)} released for production",
        },
        {
            'title': 'Production Entries',
            'value': productions.count(),
            'subtitle': 'Logged production runs',
            'hint': f"{productions.aggregate(total_output=Sum('output_sheets'))['total_output'] or 0:,} output sheets",
        },
        {
            'title': 'Dispatch Rows',
            'value': dispatches.count(),
            'subtitle': 'Dispatch activity',
            'hint': f"{dispatches.aggregate(total_qty=Sum('dispatch_qty'))['total_qty'] or 0:,} dispatched qty",
        },
        {
            'title': 'SKU Recipes',
            'value': sku_recipes.count(),
            'subtitle': 'Master data rows',
            'hint': f"{_status_count_map(recipe_status_rows, key='master_data_status').get('approved', 0)} approved",
        },
        {
            'title': 'PO Documents',
            'value': po_documents.count(),
            'subtitle': 'Imported source docs',
            'hint': f"{po_documents.filter(extraction_status='pending').count()} pending extraction",
        },
    ]

    attention_cards = [
        {
            'title': 'Jobs needing machine assignment',
            'value': planning_jobs.filter(Q(machine_name__isnull=True) | Q(machine_name='')).count(),
            'hint': 'Planning load cannot be balanced until a machine is assigned.',
        },
        {
            'title': 'Jobs missing plate set',
            'value': planning_jobs.filter(Q(plate_set_no__isnull=True) | Q(plate_set_no='')).count(),
            'hint': 'Plate readiness is a common QC blocker.',
        },
        {
            'title': 'Jobs on hold',
            'value': planning_jobs.filter(is_on_hold=True).count(),
            'hint': 'Use this to spot blocked work and aging issues.',
        },
        {
            'title': 'Unapproved SKUs',
            'value': sku_recipes.exclude(master_data_status='approved').count(),
            'hint': 'These may slow down new-job finalization.',
        },
    ]

    recent_changes = ChangeLog.objects.filter(created_at__year=year).select_related('changed_by').order_by('-created_at')[:10]

    return {
        'report_cards': REPORT_CATALOG,
        'module_cards': module_cards,
        'attention_cards': attention_cards,
        'recent_changes': recent_changes,
        'planning_status_rows': planning_status_rows,
        'job_card_status_rows': job_card_status_rows,
        'recipe_status_rows': recipe_status_rows,
        'year': year,
    }


def build_machine_planning_context(request):
    start, end, days = _date_window(request, default_days=45)
    year = _get_year(request)
    search_value = (request.GET.get('q') or '').strip()
    status_value = (request.GET.get('status') or '').strip()

    jobs = PlanningJob.objects.filter(is_active=True, plan_date__year=year)
    jobs = _search_jobs(jobs, search_value)
    if status_value:
        jobs = jobs.filter(status=status_value)

    planned_window = jobs.filter(Q(plan_date__range=(start, end)) | Q(plan_date__isnull=True))
    machine_rows = list(
        planned_window.values('machine_name')
        .annotate(
            job_count=Count('id'),
            total_order_qty=Sum('order_qty'),
            total_sheet_qty=Sum('actual_sheet_required'),
            total_balance_qty=Sum('balance_qty'),
            avg_aging_days=Avg('aging_days'),
            on_hold_count=Count('id', filter=Q(is_on_hold=True)),
        )
        .order_by('-total_sheet_qty', '-job_count')
    )

    machine_map = {machine.name.lower(): machine for machine in Machine.objects.filter(is_active=True)}
    for row in machine_rows:
        machine_name = (row['machine_name'] or '').strip()
        machine = machine_map.get(machine_name.lower()) if machine_name else None
        row['machine_display'] = machine_name or 'Unassigned'
        row['machine_capacity_per_hour'] = machine.standard_impressions_per_hour if machine else None
        row['machine_setup_minutes'] = machine.standard_setup_minutes_per_color if machine else None

    unassigned_jobs = jobs.filter(Q(machine_name__isnull=True) | Q(machine_name='')).order_by('plan_date', 'delivery_date', '-aging_days')[:10]
    urgent_jobs = jobs.filter(Q(plan_date__lt=start) | Q(delivery_date__lt=end)).order_by('delivery_date', 'plan_date')[:12]
    status_rows = jobs.values('status').annotate(count=Count('id')).order_by('status')

    summary = jobs.aggregate(
        total_jobs=Count('id'),
        total_order_qty=Sum('order_qty'),
        total_sheets=Sum('actual_sheet_required'),
        avg_aging_days=Avg('aging_days'),
    )

    actual_runs = Production.objects.filter(is_active=True, date__range=(start, end))
    actual_rows = list(
        actual_runs.values('machine__name')
        .annotate(
            run_count=Count('id'),
            output_sheets=Sum('output_sheets'),
            waste_sheets=Sum('waste_sheets'),
            impressions=Sum('impressions'),
            run_time=Sum('run_time'),
            downtime_minutes=Sum('downtime_minutes'),
            make_ready_time=Sum('make_ready_time'),
        )
        .order_by('-output_sheets', '-run_count')
    )
    for row in actual_rows:
        machine_name = (row['machine__name'] or '').strip()
        machine = machine_map.get(machine_name.lower()) if machine_name else None
        row['machine_display'] = machine_name or 'Unassigned'
        row['waste_rate'] = round((row['waste_sheets'] or 0) * 100 / max((row['output_sheets'] or 0) + (row['waste_sheets'] or 0), 1), 1)
        row['throughput_per_hour'] = round((row['impressions'] or 0) * 60 / max(row['run_time'] or 0, 1), 1)
        row['capacity_utilization'] = round(
            (row['throughput_per_hour'] or 0) * 100 / max((machine.standard_impressions_per_hour if machine else 0) or 1, 1),
            1,
        ) if machine else None

    return {
        'report': next(item for item in REPORT_CATALOG if item['key'] == 'machine-planning'),
        'filters': {
            'q': search_value,
            'status': status_value,
            'days': days,
            'start': start,
            'end': end,
            'year': year,
        },
        'summary': summary,
        'status_rows': status_rows,
        'machine_rows': machine_rows,
        'actual_rows': actual_rows,
        'unassigned_jobs': unassigned_jobs,
        'urgent_jobs': urgent_jobs,
        'status_choices': PlanningJob._meta.get_field('status').choices,
        'machine_choices': Machine.objects.filter(is_active=True).order_by('name').values_list('name', flat=True),
    }


def build_job_planning_context(request):
    start, end, days = _date_window(request, default_days=30)
    year = _get_year(request)
    search_value = (request.GET.get('q') or '').strip()
    status_value = (request.GET.get('status') or '').strip()
    department_value = (request.GET.get('department') or '').strip()

    jobs = PlanningJob.objects.filter(is_active=True, plan_date__year=year)
    jobs = _search_jobs(jobs, search_value)
    if status_value:
        jobs = jobs.filter(status=status_value)
    if department_value:
        jobs = jobs.filter(department__iexact=department_value)

    summary = jobs.aggregate(
        total_jobs=Count('id'),
        total_order_qty=Sum('order_qty'),
        total_sheets=Sum('actual_sheet_required'),
        total_balance=Sum('balance_qty'),
        avg_aging_days=Avg('aging_days'),
    )

    status_rows = jobs.values('status').annotate(count=Count('id')).order_by('status')
    repeat_rows = jobs.values('repeat_flag').annotate(count=Count('id')).order_by('-count', 'repeat_flag')
    department_rows = jobs.exclude(department='').values('department').annotate(count=Count('id')).order_by('-count', 'department')

    overdue_jobs = jobs.filter(delivery_date__lt=timezone.localdate()).exclude(status='completed').order_by('delivery_date', 'plan_date')[:12]
    due_soon_jobs = jobs.filter(delivery_date__range=(start, end)).exclude(status='completed').order_by('delivery_date')[:12]
    missing_readiness = jobs.filter(
        Q(machine_name__isnull=True) | Q(machine_name='') |
        Q(plate_set_no__isnull=True) | Q(plate_set_no='') |
        Q(total_colors__isnull=True) |
        Q(ups__isnull=True)
    ).order_by('-aging_days', 'delivery_date')[:15]

    return {
        'report': next(item for item in REPORT_CATALOG if item['key'] == 'job-planning'),
        'filters': {
            'q': search_value,
            'status': status_value,
            'department': department_value,
            'days': days,
            'start': start,
            'end': end,
            'year': year,
        },
        'summary': summary,
        'status_rows': status_rows,
        'repeat_rows': repeat_rows,
        'department_rows': department_rows,
        'overdue_jobs': overdue_jobs,
        'due_soon_jobs': due_soon_jobs,
        'missing_readiness': missing_readiness,
        'status_choices': PlanningJob._meta.get_field('status').choices,
        'department_choices': sorted({value for value in jobs.values_list('department', flat=True) if value}),
    }


def build_plates_planning_context(request):
    start, end, days = _date_window(request, default_days=30)
    search_value = (request.GET.get('q') or '').strip()
    status_value = (request.GET.get('status') or '').strip()
    plate_state = (request.GET.get('plate_state') or '').strip()

    jobs = PlanningJob.objects.filter(is_active=True)
    jobs = _search_jobs(jobs, search_value)
    if status_value:
        jobs = jobs.filter(status=status_value)
    if plate_state == 'missing':
        jobs = jobs.filter(Q(plate_set_no__isnull=True) | Q(plate_set_no=''))
    elif plate_state == 'ready':
        jobs = jobs.exclude(Q(plate_set_no__isnull=True) | Q(plate_set_no=''))

    summary = jobs.aggregate(
        total_jobs=Count('id'),
        missing_plate_count=Count('id', filter=Q(plate_set_no__isnull=True) | Q(plate_set_no='')),
        ready_plate_count=Count('id', filter=~(Q(plate_set_no__isnull=True) | Q(plate_set_no=''))),
        qc_blockers=Count('id', filter=Q(status__in=['pending_qc', 'qc_approved']) & (Q(plate_set_no__isnull=True) | Q(plate_set_no=''))),
    )

    missing_plate_jobs = jobs.filter(Q(plate_set_no__isnull=True) | Q(plate_set_no='')).order_by('-aging_days', 'delivery_date')[:15]
    ready_jobs = jobs.exclude(Q(plate_set_no__isnull=True) | Q(plate_set_no='')).order_by('delivery_date', '-updated_at')[:15]
    department_rows = jobs.exclude(department='').values('department').annotate(count=Count('id')).order_by('-count', 'department')
    qc_blocked_jobs = jobs.filter(
        status__in=['pending_qc', 'qc_approved']
    ).filter(Q(plate_set_no__isnull=True) | Q(plate_set_no='')).order_by('delivery_date', '-aging_days')[:12]

    return {
        'report': next(item for item in REPORT_CATALOG if item['key'] == 'plates-planning'),
        'filters': {
            'q': search_value,
            'status': status_value,
            'plate_state': plate_state,
            'days': days,
            'start': start,
            'end': end,
        },
        'summary': summary,
        'missing_plate_jobs': missing_plate_jobs,
        'ready_jobs': ready_jobs,
        'department_rows': department_rows,
        'qc_blocked_jobs': qc_blocked_jobs,
        'status_choices': PlanningJob._meta.get_field('status').choices,
    }


def build_production_insights_context(request):
    start, end, days = _date_window(request, default_days=30)
    year = _get_year(request)
    search_value = (request.GET.get('q') or '').strip()
    machine_filter = (request.GET.get('machine') or '').strip()
    shift_filter = (request.GET.get('shift') or '').strip()

    productions = Production.objects.filter(is_active=True, date__range=(start, end), date__year=year)
    if search_value:
        productions = productions.filter(
            Q(job_card__job_card_no__icontains=search_value)
            | Q(job_card__PO_No__icontains=search_value)
            | Q(job_card__SKU__icontains=search_value)
        )
    if machine_filter:
        productions = productions.filter(machine__name__iexact=machine_filter)
    if shift_filter:
        productions = productions.filter(shift=shift_filter)

    summary = productions.aggregate(
        total_runs=Count('id'),
        total_output_sheets=Sum('output_sheets'),
        total_waste_sheets=Sum('waste_sheets'),
        total_impressions=Sum('impressions'),
        total_planned_time=Sum('planned_time'),
        total_run_time=Sum('run_time'),
        total_downtime=Sum('downtime_minutes'),
        total_make_ready_time=Sum('make_ready_time'),
        avg_run_rate=Avg('ideal_run_rate'),
    )

    by_machine = list(
        productions.values('machine__name')
        .annotate(
            run_count=Count('id'),
            output_sheets=Sum('output_sheets'),
            waste_sheets=Sum('waste_sheets'),
            impressions=Sum('impressions'),
            planned_time=Sum('planned_time'),
            run_time=Sum('run_time'),
            downtime_minutes=Sum('downtime_minutes'),
            make_ready_time=Sum('make_ready_time'),
            avg_ideal_rate=Avg('ideal_run_rate'),
        )
        .order_by('-output_sheets', '-run_count')
    )
    machine_map = {machine.name.lower(): machine for machine in Machine.objects.filter(is_active=True)}
    for row in by_machine:
        machine_name = (row['machine__name'] or '').strip()
        machine = machine_map.get(machine_name.lower()) if machine_name else None
        row['machine_display'] = machine_name or 'Unassigned'
        row['waste_rate'] = round((row['waste_sheets'] or 0) * 100 / max((row['output_sheets'] or 0) + (row['waste_sheets'] or 0), 1), 1)
        row['throughput_per_hour'] = round((row['impressions'] or 0) * 60 / max(row['run_time'] or 0, 1), 1)
        row['capacity_utilization'] = round((row['throughput_per_hour'] or 0) * 100 / max((machine.standard_impressions_per_hour if machine else 0) or 1, 1), 1) if machine else None

    by_shift = productions.values('shift').annotate(
        run_count=Count('id'),
        output_sheets=Sum('output_sheets'),
        waste_sheets=Sum('waste_sheets'),
        impressions=Sum('impressions'),
        downtime_minutes=Sum('downtime_minutes'),
        run_time=Sum('run_time'),
    ).order_by('shift')

    waste_by_reason = productions.exclude(waste_reason='').values('waste_reason').annotate(
        count=Count('id'),
        waste_sheets=Sum('waste_sheets'),
    ).order_by('-waste_sheets', '-count')
    downtime_by_category = ProductionDowntime.objects.filter(production__is_active=True, production__date__range=(start, end)).values('category').annotate(
        count=Count('id'),
        minutes=Sum('minutes'),
    ).order_by('-minutes', '-count')
    top_runs = productions.select_related('job_card', 'machine', 'operator').order_by('-impressions', '-output_sheets')[:15]

    return {
        'report': next(item for item in REPORT_CATALOG if item['key'] == 'production-insights'),
        'filters': {
            'q': search_value,
            'machine': machine_filter,
            'shift': shift_filter,
            'days': days,
            'start': start,
            'end': end,
            'year': year,
        },
        'summary': summary,
        'by_machine': by_machine,
        'by_shift': by_shift,
        'waste_by_reason': waste_by_reason,
        'downtime_by_category': downtime_by_category,
        'top_runs': top_runs,
        'machine_choices': Machine.objects.filter(is_active=True).order_by('name').values_list('name', flat=True),
        'shift_choices': Production._meta.get_field('shift').choices,
    }


def build_qc_approvals_context(request):
    start, end, days = _date_window(request, default_days=30)
    year = _get_year(request)
    search_value = (request.GET.get('q') or '').strip()
    status_value = (request.GET.get('status') or '').strip()

    job_cards = JobCard.objects.filter(is_active=True, created_at__year=year)
    if search_value:
        job_cards = job_cards.filter(
            Q(job_card_no__icontains=search_value)
            | Q(PO_No__icontains=search_value)
            | Q(SKU__icontains=search_value)
        )
    if status_value:
        job_cards = job_cards.filter(status=status_value)

    qc_statuses = ['pending_qc', 'qc_approved', 'qc_rejected', 'pending_pm_approval', 'production_approved', 'pm_rejected']
    qc_jobs = job_cards.filter(status__in=qc_statuses)

    summary = qc_jobs.aggregate(
        total_in_qc=Count('id'),
        pending_qc=Count('id', filter=Q(status='pending_qc')),
        qc_approved=Count('id', filter=Q(status='qc_approved')),
        qc_rejected=Count('id', filter=Q(status='qc_rejected')),
        pending_pm=Count('id', filter=Q(status='pending_pm_approval')),
        production_approved=Count('id', filter=Q(status='production_approved')),
    )

    status_rows = qc_jobs.values('status').annotate(count=Count('id')).order_by('status')
    pending_qc_jobs = job_cards.filter(status='pending_qc').select_related('machine_name', 'planning_job').order_by('created_at')[:15]
    rejected_jobs = job_cards.filter(status__in=['qc_rejected', 'pm_rejected']).select_related('machine_name', 'planning_job').order_by('-updated_at')[:15]
    approved_recent = job_cards.filter(status__in=['qc_approved', 'production_approved']).select_related('machine_name', 'planning_job').order_by('-updated_at')[:15]

    return {
        'report': next(item for item in REPORT_CATALOG if item['key'] == 'qc-approvals'),
        'filters': {
            'q': search_value,
            'status': status_value,
            'days': days,
            'start': start,
            'end': end,
            'year': year,
        },
        'summary': summary,
        'status_rows': status_rows,
        'pending_qc_jobs': pending_qc_jobs,
        'rejected_jobs': rejected_jobs,
        'approved_recent': approved_recent,
        'status_choices': JobCard._meta.get_field('status').choices,
    }


def build_dispatch_tracking_context(request):
    start, end, days = _date_window(request, default_days=30)
    year = _get_year(request)
    search_value = (request.GET.get('q') or '').strip()

    dispatches = Dispatch.objects.filter(is_active=True, dispatch_date__range=(start, end), dispatch_date__year=year)
    if search_value:
        dispatches = dispatches.filter(
            Q(job_card__job_card_no__icontains=search_value)
            | Q(job_card__PO_No__icontains=search_value)
            | Q(job_card__SKU__icontains=search_value)
            | Q(dc_no__icontains=search_value)
        )

    summary = dispatches.aggregate(
        total_dispatches=Count('id'),
        total_dispatch_qty=Sum('dispatch_qty'),
    )

    job_cards_with_dispatch = JobCard.objects.filter(
        is_active=True,
        dispatch__is_active=True,
        dispatch__dispatch_date__range=(start, end)
    ).distinct().select_related('machine_name', 'planning_job')

    fulfillment_summary = job_cards_with_dispatch.aggregate(
        total_jobs=Count('id'),
        avg_completion=Avg('order_qty'),
    )

    by_date = dispatches.values('dispatch_date').annotate(
        dispatch_count=Count('id'),
        total_qty=Sum('dispatch_qty'),
    ).order_by('-dispatch_date')[:15]

    by_dc = dispatches.exclude(dc_no='').values('dc_no').annotate(
        dispatch_count=Count('id'),
        total_qty=Sum('dispatch_qty'),
    ).order_by('-total_qty')[:15]

    recent_dispatches = dispatches.select_related('job_card', 'job_card__machine_name', 'job_card__planning_job').order_by('-dispatch_date', '-id')[:20]

    backlog_jobs = JobCard.objects.filter(
        is_active=True,
        status__in=['in_production', 'completed']
    ).exclude(
        id__in=Dispatch.objects.filter(is_active=True).values_list('job_card_id', flat=True)
    ).select_related('machine_name', 'planning_job').order_by('-updated_at')[:15]

    return {
        'report': next(item for item in REPORT_CATALOG if item['key'] == 'dispatch-tracking'),
        'filters': {
            'q': search_value,
            'days': days,
            'start': start,
            'end': end,
            'year': year,
        },
        'summary': {**summary, **fulfillment_summary},
        'by_date': by_date,
        'by_dc': by_dc,
        'recent_dispatches': recent_dispatches,
        'backlog_jobs': backlog_jobs,
    }


def build_raw_material_cutting_context(request):
    start, end, days = _date_window(request, default_days=30)
    year = _get_year(request)
    search_value = (request.GET.get('q') or '').strip()

    # Get jobs that are released for production
    jobs = PlanningJob.objects.filter(
        is_active=True,
        status__in=['released', 'in_production', 'completed'],
        plan_date__year=year,
    ).order_by('plan_date', 'jc_number')
    
    jobs = _search_jobs(jobs, search_value)

    # Build the cutting request rows
    cutting_request_rows = []
    for idx, job in enumerate(jobs, 1):
        row = {
            'sno': idx,
            'date': job.created_at.strftime('%d/%m/%Y') if job.created_at else '',
            'po': job.po_number or '-',
            'jc': job.jc_number or '-',
            'sku': job.sku or '-',
            'order_qty': job.order_qty or 0,
            'material': job.material_display or '-',
            'purchase_sheet_ups': job.purchase_sheet_ups_display or '-',
            'purchase_sheet_size': job.purchase_sheet_size_display or '-',
            'purchase_sheet_required': job.purchase_sheet_required_display or 0,
            'print_sheet_size': job.print_sheet_size_display or '-',
            'print_sheet_required': job.calculated_sheets_required or 0,
            'wastage_sheets': job.wastage_sheets or 0,
            'plate_set_no': job.plate_set_no or '-',
            'remarks': job.remarks or '-',
            'status': job.get_status_display() or '-',
        }
        cutting_request_rows.append(row)

    summary = jobs.aggregate(
        total_jobs=Count('id'),
        total_order_qty=Sum('order_qty'),
        total_purchase_sheets=Sum('purchase_sheet_required'),
        total_print_sheets=Sum('actual_sheet_required'),
        total_wastage=Sum('wastage_sheets'),
    )

    return {
        'report': next(item for item in REPORT_CATALOG if item['key'] == 'raw-material-cutting-request'),
        'filters': {
            'q': search_value,
            'days': days,
            'start': start,
            'end': end,
            'year': year,
        },
        'summary': summary,
        'cutting_request_rows': cutting_request_rows,
    }


def build_report_context(report_type, request):
    builders = {
        'machine-planning': build_machine_planning_context,
        'job-planning': build_job_planning_context,
        'plates-planning': build_plates_planning_context,
        'production-insights': build_production_insights_context,
        'qc-approvals': build_qc_approvals_context,
        'dispatch-tracking': build_dispatch_tracking_context,
        'raw-material-cutting-request': build_raw_material_cutting_context,
    }
    builder = builders.get(report_type)
    if builder is None:
        return None
    return builder(request)
