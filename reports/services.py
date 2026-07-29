from __future__ import annotations

from datetime import datetime, timedelta

from django.db.models import (
    Avg,
    Count,
    ExpressionWrapper,
    F,
    FloatField,
    OuterRef,
    Prefetch,
    Q,
    Subquery,
    Sum,
    Value,
)
from django.db.models.functions import Coalesce
from django.utils import timezone

from core.models import (
    JOB_CARD_DISPATCHABLE_STATUSES,
    ChangeLog,
    Dispatch,
    JobCard,
    Machine,
    Production,
    ProductionDowntime,
)
from planning.models import PlanningJob, PoDocument, SkuRecipe

PLANNING_NOT_RELEASED_STATUSES = ('draft', 'pending_qc', 'qc_approved')
PRODUCTION_WIP_STATUSES = ('released', 'in_production')


REPORT_CATALOG = [
    {
        'key': 'daily-production',
        'title': 'Daily Production',
        'description': 'Day-by-day printing impressions, packing output, and dispatch quantities.',
        'focus': 'Execution',
    },
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
    {
        'key': 'wastage-report',
        'title': 'Wastage Report',
        'description': 'Process-wise wastage analysis including printing, sorting, and dispatch gaps (tentative vs finalized).',
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


def _parse_period_filter(request, default_period='month'):
    """Resolve dashboard period presets or explicit date range."""
    today = timezone.localdate()
    requested_period = (request.GET.get('period') or '').strip().lower()
    period = (requested_period or default_period).strip().lower()

    if period == 'all':
        start = None
        end = None
        label = 'All Time'
        return start, end, period, label, '', ''

    date_from_raw = (request.GET.get('date_from') or '').strip()
    date_to_raw = (request.GET.get('date_to') or '').strip()

    # An explicitly chosen preset wins over date_from/date_to. The filter forms
    # render the *resolved* range of the current preset into the date inputs, so
    # those fields are always populated; without this guard they would silently
    # override every preset the user picks and the report would never change.
    preset_wins = bool(requested_period) and requested_period != 'custom'

    if not preset_wins and date_from_raw and date_to_raw:
        try:
            start = datetime.strptime(date_from_raw, '%Y-%m-%d').date()
            end = datetime.strptime(date_to_raw, '%Y-%m-%d').date()
            if start > end:
                start, end = end, start
            label = f'{start.strftime("%d %b")} – {end.strftime("%d %b %Y")}'
            return start, end, 'custom', label, start.isoformat(), end.isoformat()
        except ValueError:
            pass

    if period == 'today':
        start = end = today
        label = 'Today'
    elif period == 'week':
        start = today - timedelta(days=today.weekday())
        end = today
        label = 'This Week'
    elif period == 'month':
        start = today.replace(day=1)
        end = today
        label = 'This Month'
    elif period == 'days':
        days = _parse_days(request, default_days=30)
        end = today
        start = end - timedelta(days=days - 1)
        label = f'Last {days} days'
    elif period == 'all':
        start = None
        end = None
        label = 'All Time'
    else:
        start = today.replace(day=1)
        end = today
        period = 'month'
        label = 'This Month'

    return start, end, period, label, start.isoformat() if start else '', end.isoformat() if end else ''


def _filter_planning_jobs_by_period(queryset, start, end):
    if start is None or end is None:
        return queryset
    return queryset.filter(
        Q(plan_date__range=(start, end))
        | (Q(plan_date__isnull=True) & Q(created_at__date__range=(start, end)))
    )


def _filter_job_cards_by_period(queryset, start, end):
    if start is None or end is None:
        return queryset
    # Filter on created_at (system-entry date) which is what we display as Plan Date.
    # Fall back to planning_job__po_approval_date or planning_job__plan_date only
    # when created_at is somehow absent (practically never happens on JobCard).
    return queryset.filter(
        Q(created_at__date__range=(start, end))
        | (
            Q(created_at__isnull=True)
            & Q(planning_job__po_approval_date__range=(start, end))
        )
        | (
            Q(created_at__isnull=True)
            & Q(planning_job__po_approval_date__isnull=True)
            & Q(planning_job__plan_date__range=(start, end))
        )
    )


def _annotate_dispatch_balance(queryset):
    dispatched_subquery = (
        Dispatch.objects.filter(job_card_id=OuterRef('pk'), is_active=True)
        .values('job_card_id')
        .annotate(total=Sum('dispatch_qty'))
        .values('total')
    )
    return queryset.annotate(
        dispatched_total=Coalesce(Subquery(dispatched_subquery), Value(0)),
        balance_dispatch_qty=F('order_qty') - F('dispatched_total'),
    )


def _active_productions(job_card):
    cached = getattr(job_card, '_prefetched_objects_cache', {}).get('productions')
    if cached is not None:
        return [production for production in cached if production.is_active]
    return list(job_card.productions.filter(is_active=True))


def _sum_balance_impressions(job_cards):
    total = 0
    for job_card in job_cards:
        used = sum(
            production.impressions or 0
            for production in _active_productions(job_card)
            if production.entry_type == 'printing'
        )
        allowed = job_card.total_impressions_allowed_with_tolerance or 0
        total += max(0, allowed - used)
    return total


def _sum_make_ready_colors(job_cards):
    total = 0
    for job_card in job_cards:
        colors = job_card.total_colors
        if colors:
            total += int(colors)
    return total


def _sum_balance_packing(job_cards):
    total = 0
    for job_card in job_cards:
        packing_entries = [
            production
            for production in _active_productions(job_card)
            if production.entry_type == 'packing'
        ]
        packed = sum(production.packing_qty or 0 for production in packing_entries)
        waste = sum(production.sorting_waste_qty or 0 for production in packing_entries)
        used = int(packed) + int(waste)
        if job_card.is_print_job:
            printed = sum(
                production.pcs_produced
                for production in _active_productions(job_card)
                if production.entry_type == 'printing'
            )
            limit = int(printed or 0)
        else:
            limit = int(job_card.order_qty or 0)
        total += max(0, limit - used)
    return total


def build_dashboard_context(request):
    start, end, period, period_label, date_from, date_to = _parse_period_filter(request)
    year = _get_year(request)

    pending_planning_jobs = _filter_planning_jobs_by_period(
        PlanningJob.objects.filter(
            is_active=True,
            is_on_hold=False,
            issued_to_production=False,
            status__in=PLANNING_NOT_RELEASED_STATUSES,
        ),
        start,
        end,
    )
    planning_summary = pending_planning_jobs.aggregate(
        pending_jobs=Count('id'),
        pending_order_qty=Sum('order_qty'),
    )
    planning_pending_skus = (
        pending_planning_jobs.exclude(sku='')
        .values('sku')
        .distinct()
        .count()
    )

    production_job_cards = list(
        _filter_job_cards_by_period(
            JobCard.objects.filter(
                is_active=True,
                status__in=PRODUCTION_WIP_STATUSES,
            ).select_related('planning_job'),
            start,
            end,
        ).prefetch_related(
            Prefetch(
                'productions',
                queryset=Production.objects.filter(is_active=True),
            ),
        )
    )
    production_jobs_count = len(production_job_cards)
    production_balance_impressions = _sum_balance_impressions(production_job_cards)
    production_make_ready_colors = _sum_make_ready_colors(production_job_cards)
    production_balance_packing = _sum_balance_packing(production_job_cards)

    period_production_activity = Production.objects.filter(
        is_active=True,
        job_card__is_active=True,
        date__range=(start, end),
    )
    period_printing_impressions = (
        period_production_activity.filter(entry_type='printing').aggregate(total=Sum('impressions'))['total'] or 0
    )
    period_packed_qty = (
        period_production_activity.filter(entry_type='packing').aggregate(total=Sum('packing_qty'))['total'] or 0
    )

    dispatch_jobs = _annotate_dispatch_balance(
        _filter_job_cards_by_period(
            JobCard.objects.filter(
                is_active=True,
                status__in=JOB_CARD_DISPATCHABLE_STATUSES,
            ).select_related('planning_job'),
            start,
            end,
        )
    )
    dispatch_open_jobs = dispatch_jobs.filter(balance_dispatch_qty__gt=0)
    dispatch_summary = dispatch_open_jobs.aggregate(
        open_jobs=Count('id'),
        balance_dispatch_qty=Sum('balance_dispatch_qty'),
    )
    period_dispatches = Dispatch.objects.filter(
        is_active=True,
        job_card__is_active=True,
        dispatch_date__range=(start, end),
    )
    period_dispatch_summary = period_dispatches.aggregate(
        dispatch_rows=Count('id'),
        total_dispatch_qty=Sum('dispatch_qty'),
    )
    period_dispatch_jobs = period_dispatches.values('job_card_id').distinct().count()

    return {
        'dashboard': {
            'planning': {
                'pending_jobs': planning_summary['pending_jobs'] or 0,
                'pending_skus': planning_pending_skus,
                'pending_order_qty': planning_summary['pending_order_qty'] or 0,
            },
            'production': {
                'jobs': production_jobs_count,
                'balance_impressions': production_balance_impressions,
                'make_ready_colors': production_make_ready_colors,
                'balance_packing': production_balance_packing,
                'period_impressions': period_printing_impressions,
                'period_packed_qty': period_packed_qty,
            },
            'dispatch': {
                'open_jobs': dispatch_summary['open_jobs'] or 0,
                'balance_dispatch_qty': dispatch_summary['balance_dispatch_qty'] or 0,
                'period_dispatch_qty': period_dispatch_summary['total_dispatch_qty'] or 0,
                'period_dispatch_jobs': period_dispatch_jobs,
                'period_dispatch_rows': period_dispatch_summary['dispatch_rows'] or 0,
            },
        },
        'filters': {
            'period': period,
            'period_label': period_label,
            'start': start,
            'end': end,
            'date_from': date_from,
            'date_to': date_to,
            'year': year,
        },
    }


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
    dashboard_context = build_dashboard_context(request)
    year = dashboard_context['filters']['year']
    
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
        **dashboard_context,
        'report_cards': REPORT_CATALOG,
        'module_cards': module_cards,
        'attention_cards': attention_cards,
        'recent_changes': recent_changes,
        'planning_status_rows': planning_status_rows,
        'job_card_status_rows': job_card_status_rows,
        'recipe_status_rows': recipe_status_rows,
        'year': year,
    }


PARTIAL_PRODUCTION_DONE_TOLERANCE_PCT = 5  # remaining <= 5% of planned qty counts as fully done
# Manager has all planning-console permissions a planner/admin has (short of
# deletion/superuser-only actions), so it's included here alongside them.
MACHINE_PLANNING_SELECTION_ROLES = {'planner', 'admin', 'manager'}


def _can_edit_jc_selection(user):
    """V2 plan item 3: Planner/Admin/Manager (or superuser) may toggle JC selection."""
    if not user or not user.is_authenticated:
        return False
    if user.is_superuser:
        return True
    profile = getattr(user, 'profile', None)
    return bool(profile and profile.role in MACHINE_PLANNING_SELECTION_ROLES)


def _production_state(job, produced_sheets_by_jobcard):
    """Classify a plannable job's production progress (V2 plan item 1).

    Returns not_started / partially_produced / done, plus the remaining
    sheet quantity to use for report totals instead of the full order qty.
    """
    total_planned = job.actual_sheet_required_display or 0
    job_card = getattr(job, 'job_card', None)
    produced = produced_sheets_by_jobcard.get(job_card.id, 0) if job_card else 0

    remaining_sheets = max(total_planned - produced, 0)
    percent_done = round(produced * 100 / total_planned, 1) if total_planned else 0.0

    if produced <= 0:
        state = 'not_started'
    elif total_planned and remaining_sheets <= total_planned * PARTIAL_PRODUCTION_DONE_TOLERANCE_PCT / 100:
        state = 'done'
    else:
        state = 'partially_produced'

    return {
        'state': state,
        'produced_sheets': produced,
        'remaining_sheets': remaining_sheets if state != 'not_started' else total_planned,
        'percent_done': percent_done,
        'total_planned_sheets': total_planned,
    }


def build_machine_planning_context(request):
    from django.utils import timezone
    from collections import Counter
    from datetime import date, datetime, timedelta

    start, end, days = _date_window(request, default_days=45)
    year = _get_year(request)
    search_value = (request.GET.get('q') or '').strip()
    status_value = (request.GET.get('status') or '').strip()

    # Get all pending printing jobs (excluding cut_and_pack, completed, and
    # on-hold - is_on_hold is the planner's exclude toggle, V2 plan item 1).
    base_jobs = (
        PlanningJob.objects.filter(is_active=True)
        .exclude(status='completed')
        .exclude(job_process_type='cut_and_pack')
        .exclude(is_on_hold=True)
        .select_related('job_card')
    )
    base_jobs = _search_jobs(base_jobs, search_value)
    if status_value:
        base_jobs = base_jobs.filter(status=status_value)

    # Classify production state (not_started / partially_produced / done) per
    # job so fully-produced jobs (>=95% of planned sheets already run) leave
    # the plannable set, while partially-produced jobs stay visible showing
    # their remaining balance (V2 plan item 1).
    candidate_jobs = list(base_jobs)
    job_card_ids = [j.job_card.id for j in candidate_jobs if getattr(j, 'job_card', None)]
    produced_lookup = dict(
        Production.objects.filter(is_active=True, entry_type='printing', job_card_id__in=job_card_ids)
        .values('job_card_id').annotate(total=Sum('output_sheets'))
        .values_list('job_card_id', 'total')
    )
    production_state_by_job_id = {}
    done_ids = []
    for j in candidate_jobs:
        state = _production_state(j, produced_lookup)
        production_state_by_job_id[j.id] = state
        if state['state'] == 'done':
            done_ids.append(j.id)

    jobs = base_jobs.exclude(id__in=done_ids) if done_ids else base_jobs

    # Prefetch SkuRecipes to avoid N+1 queries in loops
    from planning.models import SkuRecipe
    recipes = SkuRecipe.objects.filter(is_active=True).order_by('-updated_at')
    recipe_dict = {}
    for r in recipes:
        s_val = (r.sku or '').lower().strip()
        if s_val and s_val not in recipe_dict:
            recipe_dict[s_val] = r

    # V2 plan item 3: shared planner/admin JC opt-out. Excluded JCs are
    # dropped from merged report totals/exports but stay visible (marked
    # excluded) in the planner console.
    from reports.models import MachinePlanningJcSelection
    excluded_jc_numbers = set(
        MachinePlanningJcSelection.objects.filter(is_excluded=True).values_list('jc_number', flat=True)
    )

    for job in jobs:
        s_val = (job.sku or '').lower().strip()
        job._cached_sku_recipe = recipe_dict.get(s_val)
        job._production_state = production_state_by_job_id.get(job.id) or _production_state(job, produced_lookup)
        job._is_excluded_by_planner = job.jc_number in excluded_jc_numbers

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

    all_active_machines = list(Machine.objects.filter(is_active=True))
    machine_map = {machine.name.lower(): machine for machine in all_active_machines}

    from core.machine_routing import build_pools, route_job
    machine_pools = build_pools(all_active_machines)
    size_gate_code = next(
        (m.machine_group_code for m in all_active_machines if m.default_colors and m.default_colors >= 5),
        None,
    )
    maintenance_machines = sorted(
        {m.name for pool in machine_pools.values() for m in pool.maintenance_members}
    )

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

    # Group all pending printing jobs by (colour/size-routed machine pool, SKU) to
    # consolidate same-SKU runs. A job that's already explicitly assigned to a
    # machine stays grouped under that machine's own pool (e.g. SM74 never gets
    # folded into a GTO pool just because its colour count would fit one) -
    # only unassigned jobs get auto-routed by colour/size.
    from core.machine_routing import (
        find_pool_for_machine, find_pool_by_group_code_text, color_class as _color_class,
        parse_sheet_size_mm, pool_fits,
    )
    import math as _math

    sku_groups = {}
    today_date = timezone.localdate()

    for job in jobs:
        explicit_name = (job.machine_name or '').strip()
        routed = None
        planner_override = False
        size_warning = None
        if explicit_name:
            explicit_machine = machine_map.get(explicit_name.lower())
            pool = find_pool_for_machine(machine_pools, explicit_machine) if explicit_machine else None
            if pool is None and explicit_machine is None:
                # No exact machine match (e.g. legacy data says "GTO 1" with
                # no per-unit suffix) - fall back to matching the group code
                # so it still collapses into the combined pool tab.
                pool = find_pool_by_group_code_text(machine_pools, explicit_name)
            if pool and pool.members:
                colors = _color_class(job.color_spec_display)
                # Pools are defined by colour capability, and the job's colour
                # requirement - not a stale machine assignment - decides its
                # pool. When the assigned machine degraded (e.g. a GTO2 unit
                # dropped to operational_colors=1 and folded into the GTO1
                # pool), its 2-colour jobs must NOT follow it into the
                # 1-colour tab: they re-route by colour/size back to the
                # remaining capable pool, spreading that load over the
                # machines still in it. Multi-pass routing (a 3-colour job on
                # a 2-colour pool) is still allowed - only a pool whose class
                # a single-colour machine physically can't serve in its class
                # (colour class > pool colour capacity where a capable pool
                # exists) triggers the re-route.
                auto_routed = route_job(job.color_spec_display, job.print_sheet_size_display, machine_pools, size_gate_code)
                if colors and colors > pool.effective_colors and auto_routed and auto_routed['pool_key'] != pool.group_code:
                    routed = auto_routed
                    pool = None
                else:
                    passes = max(1, _math.ceil(colors / max(pool.effective_colors, 1))) if colors else 1
                    routed = {
                        'pool_key': pool.group_code,
                        'pool_label': pool.label,
                        'member_machines': [m.name for m in pool.members],
                        'passes': passes,
                        'color_class': colors,
                    }
                    # V2 plan item 5a: an explicit assignment the pool can
                    # actually serve wins over auto-routing - flag it as a
                    # planner override so it's clear why (e.g. a job small
                    # enough for GTO sitting on SM74).
                    if auto_routed and auto_routed['pool_key'] != pool.group_code:
                        planner_override = True
                    # V2 plan item 5b: warn (don't auto-move) when a job explicitly
                    # parked on a non-size-gate pool (a GTO) doesn't actually fit
                    # that pool's max sheet size.
                    if pool.group_code != size_gate_code:
                        size_mm = parse_sheet_size_mm(job.print_sheet_size_display)
                        if size_mm and not pool_fits(pool, size_mm):
                            size_warning = f"Exceeds {pool.group_code} max print size - move to the size-gate machine."
            if routed:
                m_name = routed['pool_label']
            elif explicit_machine:
                # Canonicalize to the stored Machine.name so case variants of
                # the same machine (e.g. 'Konica Minolta' vs 'KONICA MINOLTA'
                # in legacy job data) collapse into one tab.
                m_name = explicit_machine.name
            else:
                m_name = explicit_name
        else:
            routed = route_job(job.color_spec_display, job.print_sheet_size_display, machine_pools, size_gate_code)
            m_name = routed['pool_label'] if routed else 'Unassigned'
        job._routed_pool = routed
        job._planner_override = planner_override
        job._size_warning = size_warning
        sku_val = (job.sku or 'Unknown SKU').strip()
        key = (m_name, sku_val)
        if key not in sku_groups:
            sku_groups[key] = []
        sku_groups[key].append(job)

    # Pre-calculate dominant attributes per machine for similarity clustering scoring
    machine_dominant_attrs = {}
    for (m_name, sku_val), job_list in sku_groups.items():
        if m_name not in machine_dominant_attrs:
            machine_dominant_attrs[m_name] = {'materials': [], 'sizes': [], 'colors': []}
        for job in job_list:
            if job.material_display:
                machine_dominant_attrs[m_name]['materials'].append(job.material_display)
            if job.print_sheet_size_display:
                machine_dominant_attrs[m_name]['sizes'].append(job.print_sheet_size_display)
            if job.color_spec_display:
                machine_dominant_attrs[m_name]['colors'].append(job.color_spec_display)

    dominant_values = {}
    for m_name, attrs in machine_dominant_attrs.items():
        dominant_values[m_name] = {
            'material': Counter(attrs['materials']).most_common(1)[0][0] if attrs['materials'] else None,
            'size': Counter(attrs['sizes']).most_common(1)[0][0] if attrs['sizes'] else None,
            'color': Counter(attrs['colors']).most_common(1)[0][0] if attrs['colors'] else None,
        }

    # Process and score each merged SKU group
    all_merged_rows = []
    for (m_name, sku_val), all_group_jobs in sku_groups.items():
        # V2 plan item 3: JCs the planner deselected are excluded from the
        # merged totals/exports below but stay in all_group_jobs so the
        # planner console can still show them (marked excluded).
        group_jobs = [j for j in all_group_jobs if not getattr(j, '_is_excluded_by_planner', False)]
        if not group_jobs:
            continue
        first_job = group_jobs[0]

        # Aggregations - partially-produced jobs contribute only their
        # remaining balance, not the full order qty (V2 plan item 1/2).
        def _remaining_order_qty(j):
            state = getattr(j, '_production_state', None)
            total_sheets = (state or {}).get('total_planned_sheets') or 0
            if not state or state['state'] != 'partially_produced' or not total_sheets:
                return j.order_qty or 0
            ratio = state['remaining_sheets'] / total_sheets
            return round((j.order_qty or 0) * ratio)

        finish_qty = sum(_remaining_order_qty(j) for j in group_jobs)
        print_sheet_qty = sum((j._production_state or {}).get('remaining_sheets', j.actual_sheet_required_display or 0) for j in group_jobs)

        # Physical print passes (e.g. 1+1 front/back = 2 passes) come from
        # planning's own field (PlanningJob.effective_print_passes, synced
        # from the SKU master) - the authoritative source the planner
        # already maintains. Only fall back to the machine-pool merge-pass
        # count (how many runs a colour-count needs on a smaller-colour
        # pool) when planning hasn't set an explicit pass count.
        def _job_passes(j):
            return j.effective_print_passes or (j._routed_pool or {}).get('passes') or 1

        def _job_remaining_sheets(j):
            return (j._production_state or {}).get('remaining_sheets', j.actual_sheet_required_display or 0)

        total_impressions = sum(_job_remaining_sheets(j) * _job_passes(j) for j in group_jobs)
        row_passes = _job_passes(first_job)

        jc_numbers = sorted(list({j.jc_number for j in group_jobs if j.jc_number}))
        po_numbers = sorted(list({j.po_number for j in group_jobs if j.po_number}))
        po_count = len(po_numbers)
        
        # Max priority
        max_priority = max(j.priority for j in group_jobs)
        
        # Oldest PO date for aging
        dates = []
        for j in group_jobs:
            d_val = j.po_approval_date or j.plan_date
            if d_val:
                dates.append(d_val)
            elif j.created_at:
                dates.append(j.created_at.date())
        oldest_date = min(dates) if dates else today_date
        po_age_days = (today_date - oldest_date).days

        # Determine dominant features of this machine
        m_dominants = dominant_values.get(m_name, {'material': None, 'size': None, 'color': None})
        
        # 1. Business Priority (30% weight)
        priority_weights = {4: 100, 3: 75, 2: 50, 1: 25, 0: 10}
        priority_points = priority_weights.get(max_priority, 25)
        priority_score = priority_points * 0.30

        # 2. PO Aging (25% weight)
        aging_points = min(po_age_days, 14) / 14 * 100
        aging_score = aging_points * 0.25

        # 3. Same SKU Merge Benefit (20% weight)
        merge_points = 100 if po_count > 1 else 0
        merge_score = merge_points * 0.20

        # 4. Same Material (10% weight)
        mat_display = first_job.material_display
        material_points = 100 if m_dominants['material'] and mat_display == m_dominants['material'] else 20
        material_score = material_points * 0.10

        # 5. Same Color Setup (7% weight)
        col_display = first_job.color_spec_display
        color_points = 100 if m_dominants['color'] and col_display == m_dominants['color'] else 20
        color_score = color_points * 0.07

        # 6. Same Print Sheet Size (5% weight)
        size_display = first_job.print_sheet_size_display
        size_points = 100 if m_dominants['size'] and size_display == m_dominants['size'] else 20
        size_score = size_points * 0.05

        # 7. Machine Load Balance (3% weight)
        load_score = 50 * 0.03  # Baseline load balance score

        ai_score = round(priority_score + aging_score + merge_score + material_score + color_score + size_score + load_score, 1)

        # AI Reason generation
        reasons = []
        if max_priority == 4:
            reasons.append("Critical business priority")
        elif max_priority == 3:
            reasons.append("High business priority")

        if po_age_days >= 14:
            reasons.append("PO overdue (exceeds 14-day SLA)")
        elif po_age_days >= 11:
            reasons.append("SLA Risk: Due soon")

        if po_count > 1:
            reasons.append(f"Same SKU merged from {po_count} POs")

        if m_dominants['material'] and mat_display == m_dominants['material']:
            reasons.append("Optimizes machine material transitions")
        if m_dominants['size'] and size_display == m_dominants['size']:
            reasons.append("Minimizes size layout setup changes")
        if m_dominants['color'] and col_display == m_dominants['color']:
            reasons.append("Avoids printing color wash-up")

        ai_reason = " | ".join(reasons) if reasons else "Optimized sequence queue"

        # Determine SLA Status
        if po_age_days >= 14:
            status_pill = "Overdue"  # Red
        elif po_age_days >= 11:
            status_pill = "Due Soon"  # Orange
        elif po_age_days >= 7:
            status_pill = "Attention"  # Yellow
        else:
            status_pill = "Safe"  # Green

        row = {
            'id': first_job.id,
            'machine_name': m_name,
            'sku': sku_val,
            'sku_description': first_job.job_name or '-',
            'material': mat_display or '-',
            'colors': col_display or '-',
            'ups': first_job.ups_display or 0,
            'print_sheet_size': size_display or '-',
            'print_sheet_quantity': print_sheet_qty,
            'finish_quantity': finish_qty,
            'po_age_days': po_age_days,
            'po_count': po_count,
            'po_numbers': ", ".join(po_numbers),
            'job_card_numbers': ", ".join(jc_numbers),
            'priority': max_priority,
            'priority_display': first_job.get_priority_display(),
            'ai_score': ai_score,
            'ai_reason': ai_reason,
            'status': status_pill,
            'job_ids': [j.id for j in group_jobs],
            'passes': row_passes,
            'planned_machine_pool': m_name,
            'actual_machine': _actual_machine_for_jobs(group_jobs),
            'jobs_detail': [
                {
                    'job_id': j.id,
                    'jc_number': j.jc_number,
                    'stage': j.get_status_display(),
                    'production_state': (j._production_state or {}).get('state', 'not_started'),
                    'percent_done': (j._production_state or {}).get('percent_done', 0),
                    'remaining_sheets': (j._production_state or {}).get('remaining_sheets'),
                    'is_excluded': getattr(j, '_is_excluded_by_planner', False),
                }
                for j in all_group_jobs
            ],
            'has_partial_production': any((j._production_state or {}).get('state') == 'partially_produced' for j in group_jobs),
            'planner_override': any(getattr(j, '_planner_override', False) for j in group_jobs),
            'size_warnings': sorted({j._size_warning for j in group_jobs if getattr(j, '_size_warning', None)}),
            'total_impressions': total_impressions,
        }
        all_merged_rows.append(row)

    # Average speed/setup across a pool's members, for pool tabs (label != any
    # single machine name) where a plain machine_map lookup would miss.
    pool_label_settings = {}
    for pool in machine_pools.values():
        if pool.members:
            speeds = [m.standard_impressions_per_hour for m in pool.members if m.standard_impressions_per_hour]
            setups = [m.standard_setup_minutes_per_color for m in pool.members if m.standard_setup_minutes_per_color]
            pool_label_settings[pool.label] = (
                sum(speeds) / len(speeds) if speeds else 5000,
                sum(setups) / len(setups) if setups else 30,
            )

    # Group rows by machine and sort each machine's rows by AI Score descending
    machine_reports = {}
    unique_machines = sorted(list({row['machine_name'] for row in all_merged_rows}))

    for m_name in unique_machines:
        m_rows = [row for row in all_merged_rows if row['machine_name'] == m_name]
        m_rows.sort(key=lambda x: x['ai_score'], reverse=True)

        # Add sequence numbers
        for idx, row in enumerate(m_rows, 1):
            row['sequence'] = idx

        # Machine settings
        if m_name in pool_label_settings:
            speed, setup_minutes = pool_label_settings[m_name]
        else:
            machine = machine_map.get(m_name.lower())
            speed = machine.standard_impressions_per_hour if machine and machine.standard_impressions_per_hour else 5000
            setup_minutes = machine.standard_setup_minutes_per_color if machine and machine.standard_setup_minutes_per_color else 30

        # Per-row estimated hours for the supervisor overview (run + setup).
        # Runtime is driven by impressions (sheets x passes), not raw sheet
        # count, so a 1+1 (2-pass) job correctly takes twice as long as a
        # single-pass job of the same sheet quantity.
        for row in m_rows:
            row['estimated_hours'] = round(row['total_impressions'] / max(speed, 1000) + setup_minutes / 60, 2)

        total_print_sheets = sum(r['print_sheet_quantity'] for r in m_rows)
        total_impressions_sum = sum(r['total_impressions'] for r in m_rows)
        total_job_cards = sum(len(r['job_ids']) for r in m_rows)
        total_planned_jobs = len(m_rows)

        # Runtime and Setup calculations
        est_runtime_hours = round(total_impressions_sum / max(speed, 1000), 2)
        est_setup_time_hours = round((total_planned_jobs * setup_minutes) / 60, 2)
        
        # Setup Saved calculations (merging savings)
        setup_changes_saved = max(total_job_cards - total_planned_jobs, 0)
        setup_time_saved_hours = round((setup_changes_saved * setup_minutes) / 60, 2)

        # SLA risks
        sla_risks_count = sum(1 for r in m_rows if r['po_age_days'] >= 11)

        # Utilization (estimate based on a standard 22 available hours per day capacity)
        total_load_hours = est_runtime_hours + est_setup_time_hours
        utilization_pct = min(round((total_load_hours * 100) / 22, 1), 100.0)

        # Average AI Score
        avg_score = round(sum(r['ai_score'] for r in m_rows) / max(total_planned_jobs, 1), 1)

        # Completion Time calculation
        completion_time = datetime.now() + timedelta(hours=total_load_hours)

        machine_reports[m_name] = {
            'rows': m_rows,
            'summary': {
                'machine_name': m_name,
                'ai_planning_score': avg_score,
                'machine_utilization': utilization_pct,
                'total_planned_jobs': total_planned_jobs,
                'merged_runs': sum(1 for r in m_rows if r['po_count'] > 1),
                'individual_runs': sum(1 for r in m_rows if r['po_count'] == 1),
                'total_pos': sum(r['po_count'] for r in m_rows),
                'total_job_cards': total_job_cards,
                'total_print_sheets': total_print_sheets,
                'total_impressions': total_impressions_sum,
                'total_load_hours': round(total_load_hours, 2),
                'estimated_runtime': est_runtime_hours,
                'estimated_setup_time': est_setup_time_hours,
                'setup_changes_saved': setup_changes_saved,
                'setup_time_saved': setup_time_saved_hours,
                'jobs_at_sla_risk': sla_risks_count,
                'expected_completion_time': completion_time.strftime('%Y-%m-%d %I:%M %p'),
                'size_violations_count': sum(1 for r in m_rows if r['size_warnings']),
            }
        }

    # For exports, return detailed sequenced planning rows instead of aggregate counts
    export_machine = request.GET.get('machine')
    detailed_rows = []
    if export_machine and export_machine in machine_reports:
        detailed_rows = machine_reports[export_machine]['rows']
    else:
        for m_name in unique_machines:
            detailed_rows.extend(machine_reports[m_name]['rows'])

    headers = [
        'sequence', 'po_numbers', 'job_card_numbers', 'sku', 'po_count',
        'po_age_days', 'status', 'priority_display', 'ai_score', 'machine_name',
        'material', 'colors', 'ups', 'print_sheet_size', 'print_sheet_quantity', 'finish_quantity',
        'total_impressions', 'passes', 'estimated_hours',
    ]
    header_labels = {
        'sequence': 'S#',
        'po_numbers': 'PO Numbers',
        'job_card_numbers': 'Job Card Numbers',
        'sku': 'SKU',
        'po_count': 'PO#',
        'po_age_days': 'Age',
        'status': 'Status',
        'priority_display': 'Priority',
        'ai_score': 'AI Score',
        'machine_name': 'Machine',
        'material': 'Material',
        'colors': 'Colors',
        'ups': 'Ups',
        'print_sheet_size': 'Size',
        'print_sheet_quantity': 'Sheet Qty',
        'finish_quantity': 'Finish Qty',
        'total_impressions': 'Impr.',
        'passes': 'Passes',
        'estimated_hours': 'Hours',
    }

    return {
        'report': next(item for item in REPORT_CATALOG if item['key'] == 'machine-planning'),
        'filters': {
            'q': search_value,
            'status': status_value,
            'days': days,
            'start': start,
            'end': end,
            'year': year,
            'machine': export_machine,
        },
        'summary': summary,
        'status_rows': status_rows,
        'machine_rows': detailed_rows,
        'actual_rows': actual_rows,
        'unassigned_jobs': unassigned_jobs,
        'urgent_jobs': urgent_jobs,
        'status_choices': PlanningJob._meta.get_field('status').choices,
        'machine_choices': Machine.objects.filter(is_active=True).order_by('name').values_list('name', flat=True),
        'priority_choices': PlanningJob.PRIORITY_CHOICES,
        'machine_reports': machine_reports,
        'headers': headers,
        'header_labels': header_labels,
        'maintenance_machines': maintenance_machines,
        'can_edit_jc_selection': _can_edit_jc_selection(getattr(request, 'user', None)),
    }


def _actual_machine_for_jobs(planning_jobs):
    """Name(s) of the machine(s) that actually ran production for these jobs
    (plan-vs-actual tracking, Part C). Falls back to '-' when nothing has
    been produced yet."""
    jc_numbers = {j.jc_number for j in planning_jobs if j.jc_number}
    if not jc_numbers:
        return '-'
    names = sorted({
        p.machine.name
        for p in Production.objects.filter(
            is_active=True,
            job_card__job_card_no__in=jc_numbers,
            machine__isnull=False,
        ).select_related('machine')
    })
    return ", ".join(names) if names else '-'


def build_job_planning_context(request):
    start, end, days = _date_window(request, default_days=30)
    year = _get_year(request)
    search_value = (request.GET.get('q') or '').strip()
    status_value = (request.GET.get('status') or '').strip()
    department_value = (request.GET.get('department') or '').strip()

    jobs = PlanningJob.objects.filter(is_active=True, plan_date__year=year).exclude(status='completed')
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

    jobs = PlanningJob.objects.filter(is_active=True).exclude(status='completed')
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

    job_cards = JobCard.objects.filter(is_active=True, created_at__year=year).exclude(status='completed')
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


def build_daily_production_context(request):
    """Day-by-day printing / packing / dispatch output for the selected window.

    Each stream is aggregated independently and then aligned on the date so the
    Overview tab can show one row per day across all three.
    """
    start, end, period, period_label, date_from, date_to = _parse_period_filter(request, default_period='month')

    productions = Production.objects.filter(is_active=True)
    dispatches = Dispatch.objects.filter(is_active=True)
    if start and end:
        productions = productions.filter(date__range=(start, end))
        dispatches = dispatches.filter(dispatch_date__range=(start, end))

    machine_filter = (request.GET.get('machine') or '').strip()
    if machine_filter:
        productions = productions.filter(machine__name=machine_filter)

    # Shift lives on Production only — Dispatch has no shift, so a shift filter
    # narrows printing/packing and leaves the dispatch stream untouched.
    shift_filter = (request.GET.get('shift') or '').strip().upper()
    if shift_filter in {choice[0] for choice in Production.SHIFT_CHOICES}:
        productions = productions.filter(shift=shift_filter)
    else:
        shift_filter = ''

    # Production.expected_impressions is a Python property, so the DB equivalent
    # (run time in hours x machine rated speed) is rebuilt here as an expression.
    expected_expr = ExpressionWrapper(
        Coalesce(F('run_time'), Value(0.0)) / Value(60.0)
        * Coalesce(F('machine__standard_impressions_per_hour'), Value(0.0)),
        output_field=FloatField(),
    )
    # Printing waste is recorded in sheets; converting to pcs (to combine
    # with packing's pcs-native sorting waste into one "Process Wastage"
    # figure) needs each row's own job_card.ups, done here as a DB
    # expression rather than a Python join.
    waste_pcs_expr = ExpressionWrapper(
        Coalesce(F('waste_sheets'), Value(0)) * Coalesce(F('job_card__ups'), Value(1)),
        output_field=FloatField(),
    )
    output_pcs_expr = ExpressionWrapper(
        Coalesce(F('output_sheets'), Value(0)) * Coalesce(F('job_card__ups'), Value(1)),
        output_field=FloatField(),
    )

    printing_rows = list(
        productions.filter(entry_type='printing')
        .values('date')
        .annotate(
            # waste_pcs/output_pcs/expected_impressions must be annotated
            # before waste_sheets/output_sheets/run_time below — within a
            # single .annotate() call, F() resolves against annotations
            # already added earlier in the same call, and would otherwise
            # pick up the Sum('waste_sheets') aggregate instead of the raw
            # column (Django raises "is an aggregate" for that).
            waste_pcs=Coalesce(Sum(waste_pcs_expr), Value(0.0)),
            output_pcs=Coalesce(Sum(output_pcs_expr), Value(0.0)),
            expected_impressions=Coalesce(Sum(expected_expr), Value(0.0)),
            impressions=Coalesce(Sum('impressions'), Value(0)),
            output_sheets=Coalesce(Sum('output_sheets'), Value(0)),
            waste_sheets=Coalesce(Sum('waste_sheets'), Value(0)),
            downtime_minutes=Coalesce(Sum('downtime_minutes'), Value(0.0)),
            entries=Count('id'),
        )
        .order_by('-date')
    )
    # The report engine serialises dates to ISO strings before they reach the
    # template, so build the display label here while they are still dates.
    def _label(day):
        return day.strftime('%Y-%m-%d (%a)') if day else '—'

    for row in printing_rows:
        row['date_label'] = _label(row['date'])
        gross = (row['output_sheets'] or 0) + (row['waste_sheets'] or 0)
        row['waste_pct'] = round((row['waste_sheets'] or 0) * 100.0 / gross, 2) if gross else 0.0
        expected = row['expected_impressions'] or 0
        row['efficiency_pct'] = round((row['impressions'] or 0) * 100.0 / expected, 2) if expected else None

    packing_rows = list(
        productions.filter(entry_type='packing')
        .values('date')
        .annotate(
            packing_qty=Coalesce(Sum('packing_qty'), Value(0)),
            sorting_waste_qty=Coalesce(Sum('sorting_waste_qty'), Value(0)),
            entries=Count('id'),
        )
        .order_by('-date')
    )
    for row in packing_rows:
        row['date_label'] = _label(row['date'])
        gross = (row['packing_qty'] or 0) + (row['sorting_waste_qty'] or 0)
        row['waste_pct'] = round((row['sorting_waste_qty'] or 0) * 100.0 / gross, 2) if gross else 0.0

    dispatch_rows = list(
        dispatches.values('dispatch_date')
        .annotate(
            dispatch_qty=Coalesce(Sum('dispatch_qty'), Value(0)),
            dc_count=Count('dc_no', distinct=True),
            entries=Count('id'),
        )
        .order_by('-dispatch_date')
    )
    for row in dispatch_rows:
        row['date_label'] = _label(row['dispatch_date'])

    # Align the three streams on the calendar date for the Overview tab.
    printing_by_date = {r['date']: r for r in printing_rows}
    packing_by_date = {r['date']: r for r in packing_rows}
    dispatch_by_date = {r['dispatch_date']: r for r in dispatch_rows}

    overview_rows = []
    for day in sorted(set(printing_by_date) | set(packing_by_date) | set(dispatch_by_date), reverse=True):
        printing = printing_by_date.get(day, {})
        packing = packing_by_date.get(day, {})
        dispatch = dispatch_by_date.get(day, {})
        day_waste_pcs = (printing.get('waste_pcs') or 0.0) + (packing.get('sorting_waste_qty') or 0)
        day_gross_pcs = day_waste_pcs + (printing.get('output_pcs') or 0.0) + (packing.get('packing_qty') or 0)
        overview_rows.append({
            'date': day,
            'date_label': _label(day),
            'impressions': printing.get('impressions', 0),
            'output_sheets': printing.get('output_sheets', 0),
            'printing_waste': printing.get('waste_sheets', 0),
            'packing_qty': packing.get('packing_qty', 0),
            'packing_waste': packing.get('sorting_waste_qty', 0),
            'dispatch_qty': dispatch.get('dispatch_qty', 0),
            'dc_count': dispatch.get('dc_count', 0),
            'process_wastage_pcs': int(round(day_waste_pcs)),
            'process_wastage_pct': round((day_waste_pcs * 100.0 / day_gross_pcs), 2) if day_gross_pcs else 0.0,
        })

    def _total(rows, key):
        return sum((r.get(key) or 0) for r in rows)

    totals = {
        'impressions': _total(printing_rows, 'impressions'),
        'output_sheets': _total(printing_rows, 'output_sheets'),
        'printed_pcs': int(round(_total(printing_rows, 'output_pcs'))),
        'printing_waste': _total(printing_rows, 'waste_sheets'),
        'packing_qty': _total(packing_rows, 'packing_qty'),
        'packing_waste': _total(packing_rows, 'sorting_waste_qty'),
        'dispatch_qty': _total(dispatch_rows, 'dispatch_qty'),
        'dc_count': _total(dispatch_rows, 'dc_count'),
        'days': len(overview_rows),
    }
    totals['avg_impressions_per_day'] = round(totals['impressions'] / totals['days']) if totals['days'] else 0

    # Process Wastage = printing waste (pcs) + sorting waste (pcs), as logged
    # in real time on production entries. Deliberately excludes dispatch-gap
    # wastage — that component can't be attributed to a specific day/week
    # since a dispatch may land weeks after the production it settles; it
    # stays exclusive to the Wastage Report, finalized only at job completion.
    process_wastage_bucket = {}
    for row in printing_rows:
        day = row['date']
        if not day:
            continue
        bucket = process_wastage_bucket.setdefault(day, {'output_pcs': 0.0, 'waste_pcs': 0.0})
        bucket['output_pcs'] += row['output_pcs'] or 0.0
        bucket['waste_pcs'] += row['waste_pcs'] or 0.0
    for row in packing_rows:
        day = row['date']
        if not day:
            continue
        bucket = process_wastage_bucket.setdefault(day, {'output_pcs': 0.0, 'waste_pcs': 0.0})
        bucket['output_pcs'] += row['packing_qty'] or 0
        bucket['waste_pcs'] += row['sorting_waste_qty'] or 0

    def _process_wastage_row(period_label, bucket):
        gross = bucket['output_pcs'] + bucket['waste_pcs']
        pct = round((bucket['waste_pcs'] * 100.0 / gross), 2) if gross else 0.0
        return {
            'period_label': period_label,
            'process_wastage_pcs': int(round(bucket['waste_pcs'])),
            'process_wastage_pct': pct,
        }

    process_wastage_rows = [
        {**_process_wastage_row(_label(day), process_wastage_bucket[day]), 'date': day}
        for day in sorted(process_wastage_bucket, reverse=True)
    ]

    weekly_bucket = {}
    for day, bucket in process_wastage_bucket.items():
        week_start = day - timedelta(days=day.weekday())
        wk = weekly_bucket.setdefault(week_start, {'output_pcs': 0.0, 'waste_pcs': 0.0})
        wk['output_pcs'] += bucket['output_pcs']
        wk['waste_pcs'] += bucket['waste_pcs']

    process_wastage_weekly = []
    for week_start in sorted(weekly_bucket, reverse=True):
        week_end = week_start + timedelta(days=6)
        label = f"{week_start.strftime('%Y-%m-%d')} to {week_end.strftime('%Y-%m-%d')}"
        process_wastage_weekly.append(_process_wastage_row(label, weekly_bucket[week_start]))

    totals['process_wastage_pcs'] = sum(b['waste_pcs'] for b in process_wastage_bucket.values())
    _process_gross = totals['process_wastage_pcs'] + sum(b['output_pcs'] for b in process_wastage_bucket.values())
    totals['process_wastage_pct'] = (
        round((totals['process_wastage_pcs'] * 100.0 / _process_gross), 2) if _process_gross else 0.0
    )
    totals['process_wastage_pcs'] = int(round(totals['process_wastage_pcs']))

    # Consumption-anomaly flag: job cards, active in this window, that used
    # more sheets (output + waste) than their planned allowance + tolerance.
    # This is a correlation signal for possibly-unreported waste, not proof —
    # there's no independent stock-issuance record to check against, only
    # what operators logged. Surfaced here so it sits next to the wastage
    # figures instead of only living on Production Records.
    flagged_job_ids = set(productions.filter(entry_type='printing').values_list('job_card_id', flat=True))
    flagged_job_count = 0
    if flagged_job_ids:
        for job_card in JobCard.objects.filter(id__in=flagged_job_ids, is_active=True):
            if job_card.extra_sheets_used > job_card.tolerance_sheets:
                flagged_job_count += 1
    totals['flagged_job_count'] = flagged_job_count

    # Machine split for the printing tab.
    printing_by_machine = list(
        productions.filter(entry_type='printing')
        .values('machine__name')
        .annotate(
            impressions=Coalesce(Sum('impressions'), Value(0)),
            output_sheets=Coalesce(Sum('output_sheets'), Value(0)),
            waste_sheets=Coalesce(Sum('waste_sheets'), Value(0)),
        )
        .order_by('-impressions')
    )

    # Shift performance split.
    shift_labels = dict(Production.SHIFT_CHOICES)
    printing_by_shift = list(
        productions.filter(entry_type='printing')
        .values('shift')
        .annotate(
            impressions=Coalesce(Sum('impressions'), Value(0)),
            output_sheets=Coalesce(Sum('output_sheets'), Value(0)),
            waste_sheets=Coalesce(Sum('waste_sheets'), Value(0)),
            entries=Count('id'),
        )
        .order_by('shift')
    )
    packing_by_shift = list(
        productions.filter(entry_type='packing')
        .values('shift')
        .annotate(
            packing_qty=Coalesce(Sum('packing_qty'), Value(0)),
            sorting_waste_qty=Coalesce(Sum('sorting_waste_qty'), Value(0)),
            entries=Count('id'),
        )
        .order_by('shift')
    )
    for row in printing_by_shift + packing_by_shift:
        row['shift_label'] = shift_labels.get(row['shift'], row['shift'] or '—')

    # The generic exporter picks up data['export_rows'] with data['headers'];
    # ?tab= selects which of the four tables gets exported.
    export_tabs = {
        'overview': (overview_rows, [
            ('date_label', 'Date'), ('impressions', 'Impressions'),
            ('output_sheets', 'Printed Sheets'), ('printing_waste', 'Printing Waste'),
            ('packing_qty', 'Packed Pcs'), ('packing_waste', 'Packing Waste'),
            ('dispatch_qty', 'Dispatched Pcs'), ('dc_count', 'DCs'),
        ]),
        'printing': (printing_rows, [
            ('date_label', 'Date'), ('entries', 'Entries'), ('impressions', 'Impressions'),
            ('output_sheets', 'Output Sheets'), ('waste_sheets', 'Waste Sheets'),
            ('waste_pct', 'Waste %'), ('downtime_minutes', 'Downtime (min)'),
            ('efficiency_pct', 'Efficiency %'),
        ]),
        'packing': (packing_rows, [
            ('date_label', 'Date'), ('entries', 'Entries'), ('packing_qty', 'Packed Pcs'),
            ('sorting_waste_qty', 'Sorting Waste'), ('waste_pct', 'Waste %'),
        ]),
        'dispatch': (dispatch_rows, [
            ('date_label', 'Date'), ('entries', 'Dispatch Entries'),
            ('dc_count', 'Delivery Challans'), ('dispatch_qty', 'Dispatched Pcs'),
        ]),
        'wastage': (process_wastage_rows, [
            ('period_label', 'Date'), ('process_wastage_pcs', 'Process Wastage (Pcs)'),
            ('process_wastage_pct', 'Process Wastage %'),
        ]),
    }
    export_tab = (request.GET.get('tab') or 'overview').strip().lower()
    if export_tab not in export_tabs:
        export_tab = 'overview'
    export_rows, export_columns = export_tabs[export_tab]

    return {
        'export_rows': export_rows,
        'headers': [key for key, _ in export_columns],
        'header_labels': {key: label for key, label in export_columns},
        'export_tab': export_tab,
        'printing_by_shift': printing_by_shift,
        'packing_by_shift': packing_by_shift,
        'shift_filter': shift_filter,
        'shift_choices': list(Production.SHIFT_CHOICES),
        'overview_rows': overview_rows,
        'printing_rows': printing_rows,
        'packing_rows': packing_rows,
        'dispatch_rows': dispatch_rows,
        'process_wastage_rows': process_wastage_rows,
        'process_wastage_weekly': process_wastage_weekly,
        'printing_by_machine': printing_by_machine,
        'totals': totals,
        'period': period,
        'period_label': period_label,
        'date_from': date_from,
        'date_to': date_to,
        'machine_filter': machine_filter,
        'machines': list(
            Machine.objects.filter(is_active=True).values_list('name', flat=True).order_by('name')
        ),
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

    jobs = PlanningJob.objects.filter(
        is_active=True,
        status__in=['released', 'in_production'],
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


def build_wastage_report_context(request):
    start, end, period, period_label, date_from, date_to = _parse_period_filter(request)
    year = _get_year(request)
    search_value = (request.GET.get('q') or '').strip()
    status_filter = (request.GET.get('status') or '').strip()
    sku_filter = (request.GET.get('sku') or '').strip()
    po_filter = (request.GET.get('po') or '').strip()
    job_card_filter = (request.GET.get('job_card') or '').strip()
    machine_filter = (request.GET.get('machine') or '').strip()
    wastage_status_filter = (request.GET.get('wastage_status') or '').strip().lower()
    high_wastage_filter = request.GET.get('high_wastage') == 'true'

    job_cards = JobCard.objects.filter(is_active=True)

    # Apply date window filter if set
    if start and end:
        job_cards = _filter_job_cards_by_period(job_cards, start, end)

    # Universal search filter
    if search_value:
        job_cards = job_cards.filter(
            Q(job_card_no__icontains=search_value)
            | Q(PO_No__icontains=search_value)
            | Q(SKU__icontains=search_value)
        )

    # Individual filters
    if status_filter:
        job_cards = job_cards.filter(status=status_filter)
    if sku_filter:
        job_cards = job_cards.filter(SKU__icontains=sku_filter)
    if po_filter:
        job_cards = job_cards.filter(PO_No__icontains=po_filter)
    if job_card_filter:
        job_cards = job_cards.filter(job_card_no__icontains=job_card_filter)
    if machine_filter:
        job_cards = job_cards.filter(machine_name__name__icontains=machine_filter)

    # Prefetch related data to avoid N+1 queries
    job_cards = job_cards.select_related('planning_job', 'machine_name').prefetch_related(
        Prefetch('productions', queryset=Production.objects.filter(is_active=True)),
        Prefetch('dispatch_set', queryset=Dispatch.objects.filter(is_active=True)),
    )

    # Order chronologically by system-entry date (= created_at, same as displayed Plan Date)
    job_cards = job_cards.order_by('created_at')

    wastage_rows = []
    
    # Summary KPI totals
    total_plan_qty = 0
    total_dispatch_qty = 0
    total_printing_waste_pcs = 0
    total_sorting_waste_pcs = 0
    total_dispatch_gap_pcs = 0
    total_wastage_pcs = 0
    total_finalized_waste_pcs = 0
    total_tentative_waste_pcs = 0
    jobs_needing_reconciliation = 0

    s_no_counter = 0

    for job in job_cards:
        ups = job.ups or 1
        plan_qty_pcs = int(job.total_sheets_planned * ups)

        # Python-level summation of prefetched relations to avoid N+1 queries
        dispatch_qty_pcs = sum(d.dispatch_qty for d in job.dispatch_set.all())
        printing_waste_sheets = sum(p.waste_sheets for p in job.productions.all() if p.entry_type == 'printing')
        printing_waste_pcs = printing_waste_sheets * ups
        sorting_waste_pcs = sum(p.sorting_waste_qty for p in job.productions.all() if p.entry_type == 'packing')

        # Dispatch Gap (Pcs): Plan Qty (Pcs) - Dispatch Qty
        # Clamped to 0 minimum as per discussion
        dispatch_gap_pcs = max(plan_qty_pcs - dispatch_qty_pcs, 0)

        # Status of Wastage (Tentative vs Finalized)
        is_completed = (job.status in ('completed', 'closed') or job.job_status == 'Completed')
        wastage_status = "Finalized" if is_completed else "Tentative"

        # Apply wastage_status filter in Python to ensure consistency
        if wastage_status_filter == 'finalized' and not is_completed:
            continue
        if wastage_status_filter == 'tentative' and is_completed:
            continue

        # Extract plan date (= when the job entered the system, NOT the scheduled production date)
        # Priority:
        #   1. job.created_at.date()     — JobCard creation timestamp (actual system-entry date)
        #   2. planning_job.po_approval_date — PO approval date (user-specified fallback)
        #   3. planning_job.plan_date    — last resort; may hold a future scheduled date from CSV import
        _pj = job.planning_job
        if job.created_at:
            plan_date = job.created_at.date().strftime('%Y-%m-%d')
        elif _pj and _pj.po_approval_date:
            plan_date = _pj.po_approval_date.strftime('%Y-%m-%d')
        elif _pj and _pj.plan_date:
            plan_date = _pj.plan_date.strftime('%Y-%m-%d')
        else:
            plan_date = ''
        # plan_month: derive from the resolved plan_date (not from stored labels which may
        # be based on the CSV scheduled month — e.g. "August" — rather than system-entry month)
        if plan_date:
            plan_month = datetime.strptime(plan_date, '%Y-%m-%d').strftime('%B')
        elif _pj and _pj.plan_month:
            plan_month = _pj.plan_month
        elif job.month:
            plan_month = job.month
        else:
            plan_month = ''

        # Total Wastage (Pcs)
        job_total_waste_pcs = printing_waste_pcs + sorting_waste_pcs + dispatch_gap_pcs

        # Percentages based on plan_qty_pcs
        printing_waste_pct = round((printing_waste_pcs / plan_qty_pcs * 100), 2) if plan_qty_pcs > 0 else 0.0
        sorting_waste_pct = round((sorting_waste_pcs / plan_qty_pcs * 100), 2) if plan_qty_pcs > 0 else 0.0
        dispatch_gap_pct = round((dispatch_gap_pcs / plan_qty_pcs * 100), 2) if plan_qty_pcs > 0 else 0.0
        total_waste_pct = round((job_total_waste_pcs / plan_qty_pcs * 100), 2) if plan_qty_pcs > 0 else 0.0

        # Apply high_wastage filter (>5%)
        if high_wastage_filter and total_waste_pct <= 5.0:
            continue

        # Dispatch-gap is the part of total wastage that reported process
        # wastage (printing + sorting) did NOT explain — it only becomes
        # knowable once the job is Finalized and the real dispatch count is
        # in. Flag it against the job's own planned tolerance so a wrong or
        # understated in-process waste entry gets caught and pointed at.
        tolerance_pct = float(job.production_tolerance_percent or 5)
        needs_reconciliation_review = is_completed and dispatch_gap_pct > tolerance_pct

        s_no_counter += 1

        row = {
            's_no': s_no_counter,
            'job_card_no': job.job_card_no,
            'sku': job.SKU,
            'plan_date': plan_date,
            'plan_month': plan_month,
            'plan_qty': plan_qty_pcs,
            'dispatch_qty': dispatch_qty_pcs,
            'printing_waste_sheets': printing_waste_sheets,
            'printing_waste_pcs': printing_waste_pcs,
            'printing_waste_pct': f"{printing_waste_pct}%",
            'sorting_waste_pcs': sorting_waste_pcs,
            'sorting_waste_pct': f"{sorting_waste_pct}%",
            'difference_pcs': dispatch_gap_pcs,
            'difference_pct': f"{dispatch_gap_pct}%",
            'wastage_status': wastage_status,
            'total_wastage_pcs': job_total_waste_pcs,
            'total_wastage_pct': f"{total_waste_pct}%",
            'needs_reconciliation_review': 'Yes' if needs_reconciliation_review else '',
        }
        wastage_rows.append(row)

        # Accumulate totals for KPIs
        total_plan_qty += plan_qty_pcs
        total_dispatch_qty += dispatch_qty_pcs
        total_printing_waste_pcs += printing_waste_pcs
        total_sorting_waste_pcs += sorting_waste_pcs
        total_dispatch_gap_pcs += dispatch_gap_pcs
        total_wastage_pcs += job_total_waste_pcs

        if is_completed:
            total_finalized_waste_pcs += job_total_waste_pcs
        else:
            total_tentative_waste_pcs += job_total_waste_pcs

        if needs_reconciliation_review:
            jobs_needing_reconciliation += 1

    overall_wastage_pct = round((total_wastage_pcs / total_plan_qty * 100), 2) if total_plan_qty > 0 else 0.0

    summary = {
        'total_plan_qty': total_plan_qty,
        'total_dispatch_qty': total_dispatch_qty,
        'printing_waste_pcs': total_printing_waste_pcs,
        'sorting_waste_pcs': total_sorting_waste_pcs,
        'dispatch_gap_pcs': total_dispatch_gap_pcs,
        'total_wastage_pcs': total_wastage_pcs,
        'overall_wastage_pct': overall_wastage_pct,
        'finalized_wastage_pcs': total_finalized_waste_pcs,
        'tentative_wastage_pcs': total_tentative_waste_pcs,
        'jobs_needing_reconciliation': jobs_needing_reconciliation,
    }

    headers = [
        's_no',
        'job_card_no',
        'sku',
        'plan_date',
        'plan_month',
        'plan_qty',
        'dispatch_qty',
        'printing_waste_sheets',
        'printing_waste_pcs',
        'printing_waste_pct',
        'sorting_waste_pcs',
        'sorting_waste_pct',
        'difference_pcs',
        'difference_pct',
        'wastage_status',
        'total_wastage_pcs',
        'total_wastage_pct',
        'needs_reconciliation_review',
    ]

    header_labels = {
        's_no': 'S.No.',
        'job_card_no': 'Job Card No',
        'sku': 'SKU',
        'plan_date': 'Plan Date',
        'plan_month': 'Plan Month',
        'plan_qty': 'Plan Qty (Pcs)',
        'dispatch_qty': 'Dispatched Qty (Pcs)',
        'printing_waste_sheets': 'Printing Waste (Sheets)',
        'printing_waste_pcs': 'Printing Waste (Pcs)',
        'printing_waste_pct': 'Printing Waste %',
        'sorting_waste_pcs': 'Sorting Waste (Pcs)',
        'sorting_waste_pct': 'Sorting Waste %',
        'difference_pcs': 'Dispatch Gap (Pcs)',
        'difference_pct': 'Dispatch Gap %',
        'wastage_status': 'Wastage Status',
        'total_wastage_pcs': 'Total Wastage (Pcs)',
        'total_wastage_pct': 'Total Wastage %',
        'needs_reconciliation_review': 'Needs Review',
    }

    # Determine the date range to display in the report subtitle.
    # If the user has selected a specific date range, show that range.
    # Otherwise (all time / no filter), show the actual min/max from the data.
    if start and end:
        # User has an active date filter — show exactly what they selected
        display_start_date = start.strftime('%Y-%m-%d')
        display_end_date = end.strftime('%Y-%m-%d')
    else:
        # All-time or no date bounds — compute from actual data rows
        all_dates = sorted(row['plan_date'] for row in wastage_rows if row['plan_date'])
        display_start_date = all_dates[0] if all_dates else ''
        display_end_date = all_dates[-1] if all_dates else ''

    # Pagination logic (100 entries per page)
    is_export = request.GET.get('_export') == 'true'
    total_rows = len(wastage_rows)
    page_size = 100
    total_pages = max(1, (total_rows + page_size - 1) // page_size)

    try:
        current_page = int(request.GET.get('page') or 1)
    except ValueError:
        current_page = 1

    if current_page < 1:
        current_page = 1
    elif current_page > total_pages:
        current_page = total_pages

    if not is_export:
        start_idx = (current_page - 1) * page_size
        end_idx = start_idx + page_size
        wastage_rows_paginated = wastage_rows[start_idx:end_idx]
    else:
        wastage_rows_paginated = wastage_rows

    pagination_data = {
        'current_page': current_page,
        'total_pages': total_pages,
        'total_rows': total_rows,
        'page_size': page_size,
        'has_next': current_page < total_pages,
        'has_prev': current_page > 1,
    }

    # Day / week rollups of the same wastage figures, mirroring the Daily
    # Production report's day-by-day breakdown. Grouped from the already
    # per-job wastage_rows (keyed on each job's resolved plan_date) rather
    # than re-querying, since per-job UPS/plan_qty is only easy to resolve
    # once, at the per-job level above.
    daily_bucket = {}
    weekly_bucket = {}
    for row in wastage_rows:
        if not row['plan_date']:
            continue
        day = datetime.strptime(row['plan_date'], '%Y-%m-%d').date()
        week_start = day - timedelta(days=day.weekday())

        d = daily_bucket.setdefault(day, {'plan_qty': 0, 'total_wastage_pcs': 0})
        d['plan_qty'] += row['plan_qty']
        d['total_wastage_pcs'] += row['total_wastage_pcs']

        w = weekly_bucket.setdefault(week_start, {'plan_qty': 0, 'total_wastage_pcs': 0})
        w['plan_qty'] += row['plan_qty']
        w['total_wastage_pcs'] += row['total_wastage_pcs']

    daily_wastage = []
    for day in sorted(daily_bucket, reverse=True):
        bucket = daily_bucket[day]
        pct = round((bucket['total_wastage_pcs'] / bucket['plan_qty'] * 100), 2) if bucket['plan_qty'] > 0 else 0.0
        daily_wastage.append({
            'period_label': day.strftime('%Y-%m-%d (%a)'),
            'plan_qty': bucket['plan_qty'],
            'total_wastage_pcs': bucket['total_wastage_pcs'],
            'total_wastage_pct': pct,
        })

    weekly_wastage = []
    for week_start in sorted(weekly_bucket, reverse=True):
        bucket = weekly_bucket[week_start]
        week_end = week_start + timedelta(days=6)
        pct = round((bucket['total_wastage_pcs'] / bucket['plan_qty'] * 100), 2) if bucket['plan_qty'] > 0 else 0.0
        weekly_wastage.append({
            'period_label': f"{week_start.strftime('%Y-%m-%d')} to {week_end.strftime('%Y-%m-%d')}",
            'plan_qty': bucket['plan_qty'],
            'total_wastage_pcs': bucket['total_wastage_pcs'],
            'total_wastage_pct': pct,
        })

    return {
        'report': next(item for item in REPORT_CATALOG if item['key'] == 'wastage-report'),
        'filters': {
            'q': search_value,
            'status': status_filter,
            'sku': sku_filter,
            'po': po_filter,
            'job_card': job_card_filter,
            'machine': machine_filter,
            'wastage_status': wastage_status_filter,
            'high_wastage': 'true' if high_wastage_filter else '',
            'page': current_page,
            'period': period,
            'period_label': period_label,
            'date_from': date_from,
            'date_to': date_to,
            'start': start,
            'end': end,
            'year': year,
        },
        'summary': summary,
        'wastage_rows': wastage_rows_paginated,
        'daily_wastage': daily_wastage,
        'weekly_wastage': weekly_wastage,
        'pagination': pagination_data,
        'status_choices': JobCard._meta.get_field('status').choices,
        'headers': headers,
        'header_labels': header_labels,
        'start_date': display_start_date,
        'end_date': display_end_date,
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
        'wastage-report': build_wastage_report_context,
    }
    builder = builders.get(report_type)
    if builder is None:
        return None
    return builder(request)
