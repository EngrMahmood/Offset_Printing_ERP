"""Aggregation functions powering the production-floor TV dashboard.

Each `get_*` function returns a plain dict of numbers/lists ready for JSON
serialization and template rendering — one function per rotating screen,
plus `get_dashboard_data()` which bundles all of them for the polling
endpoint. Figures are scoped to "today" (server local date); nothing here
persists state — everything is computed live from Production/Dispatch/
JobCard rows, same as the existing reports app.
"""
from datetime import timedelta

from django.db.models import Sum
from django.utils import timezone

from core.models import JobCard, Production, Dispatch, Machine
from core.services import compute_job_card_wastage_metrics
from production.services import OEECalculator
from .models import DailyTarget

FINALIZED_STATUSES = ('completed', 'closed')


def _today():
    return timezone.localdate()


def _effective_date():
    """Today if there's any printing/packing/dispatch activity logged yet,
    otherwise the most recent day that has activity — so the TV doesn't sit
    on an all-zero screen before the shift starts or on a day off. Returns
    (date, is_fallback)."""
    today = _today()
    if Production.objects.filter(date=today).exists() or Dispatch.objects.filter(dispatch_date=today).exists():
        return today, False

    latest_production = Production.objects.order_by('-date').values_list('date', flat=True).first()
    latest_dispatch = Dispatch.objects.order_by('-dispatch_date').values_list('dispatch_date', flat=True).first()
    candidates = [d for d in (latest_production, latest_dispatch) if d]
    if candidates:
        return max(candidates), True
    return today, False


def _printing_qs(date):
    return Production.objects.filter(is_active=True, entry_type='printing', date=date)


def _packing_qs(date):
    return Production.objects.filter(is_active=True, entry_type='packing', date=date)


def _dispatch_qs(date):
    return Dispatch.objects.filter(is_active=True, dispatch_date=date)


def _period_metrics(start_date, end_date):
    """Core plant KPIs over an arbitrary [start_date, end_date] range —
    used both by get_plant_overview (range collapsed to a single day) and
    by get_period_summary (today/yesterday/week/month ranges)."""
    printing = Production.objects.filter(
        is_active=True, entry_type='printing', date__range=(start_date, end_date),
    )
    packing = Production.objects.filter(
        is_active=True, entry_type='packing', date__range=(start_date, end_date),
    )
    dispatch = Dispatch.objects.filter(
        is_active=True, dispatch_date__range=(start_date, end_date),
    )

    printed_pcs = 0
    expected_impressions_total = 0.0
    impressions_total = 0
    for row in printing.select_related('job_card', 'machine'):
        ups = row.job_card.ups or 1
        printed_pcs += (row.output_sheets or 0) * ups
        impressions_total += row.impressions or 0
        speed = row.machine.standard_impressions_per_hour if row.machine else 4000
        expected_impressions_total += speed * ((row.run_time or 0) / 60.0)

    efficiency_pct = round((impressions_total / expected_impressions_total * 100), 1) if expected_impressions_total else 0.0

    packed_pcs = packing.aggregate(total=Sum('packing_qty'))['total'] or 0
    sorting_waste_pcs = packing.aggregate(total=Sum('sorting_waste_qty'))['total'] or 0
    dispatched_pcs = dispatch.aggregate(total=Sum('dispatch_qty'))['total'] or 0

    process_wastage_pcs = sorting_waste_pcs
    for row in printing.select_related('job_card'):
        process_wastage_pcs += (row.waste_sheets or 0) * (row.job_card.ups or 1)
    produced_total = printed_pcs + packed_pcs + process_wastage_pcs
    process_wastage_pct = round((process_wastage_pcs / produced_total * 100), 1) if produced_total else 0.0

    return {
        'printed_pcs': printed_pcs,
        'packed_pcs': packed_pcs,
        'dispatched_pcs': dispatched_pcs,
        'process_wastage_pcs': process_wastage_pcs,
        'process_wastage_pct': process_wastage_pct,
        'efficiency_pct': efficiency_pct,
    }


def _period_bounds(today=None):
    today = today or _today()
    monday = today - timedelta(days=today.weekday())
    first_of_month = today.replace(day=1)
    yesterday = today - timedelta(days=1)
    return [
        {'key': 'today', 'label': 'Today', 'start': today, 'end': today},
        {'key': 'yesterday', 'label': 'Yesterday', 'start': yesterday, 'end': yesterday},
        {'key': 'week', 'label': 'This Week', 'start': monday, 'end': today},
        {'key': 'month', 'label': 'This Month', 'start': first_of_month, 'end': today},
    ]


def get_period_summary(today=None):
    periods = []
    for period in _period_bounds(today):
        metrics = _period_metrics(period['start'], period['end'])
        if period['start'] == period['end']:
            range_label = period['start'].strftime('%b %d')
        else:
            range_label = f"{period['start'].strftime('%b %d')} - {period['end'].strftime('%b %d')}"
        periods.append({
            'key': period['key'],
            'label': period['label'],
            'range_label': range_label,
            **metrics,
        })
    return {'periods': periods}


def get_plant_overview(date=None):
    date = date or _today()

    jobs_today = JobCard.objects.filter(is_active=True, productions__date=date).distinct()
    running_jobs = jobs_today.filter(status__in=('in_production', 'released')).count()
    completed_jobs = jobs_today.filter(status__in=FINALIZED_STATUSES).count()
    pending_jobs = JobCard.objects.filter(
        is_active=True, status__in=('released', 'production_approved'),
    ).exclude(productions__date=date).distinct().count()

    printing = _printing_qs(date)

    period_metrics = _period_metrics(date, date)
    printed_pcs = period_metrics['printed_pcs']
    packed_pcs = period_metrics['packed_pcs']
    dispatched_pcs = period_metrics['dispatched_pcs']
    process_wastage_pct = period_metrics['process_wastage_pct']
    efficiency_pct = period_metrics['efficiency_pct']

    waste_sheets_total = printing.aggregate(total=Sum('waste_sheets'))['total'] or 0

    active_machines = printing.exclude(machine__isnull=True).values('machine').distinct().count()
    total_machines = Machine.objects.count()
    active_operators = printing.exclude(operator__isnull=True).values('operator').distinct().count()

    target = _resolve_target(date)

    return {
        'date': date.isoformat(),
        'time': timezone.localtime().strftime('%H:%M'),
        'running_jobs': running_jobs,
        'completed_jobs': completed_jobs,
        'pending_jobs': pending_jobs,
        'efficiency_pct': efficiency_pct,
        'printed_pcs': printed_pcs,
        'packed_pcs': packed_pcs,
        'dispatched_pcs': dispatched_pcs,
        'process_wastage_pct': process_wastage_pct,
        'active_machines': active_machines,
        'total_machines': total_machines,
        'active_operators': active_operators,
        'plant_utilization_pct': round((active_machines / total_machines * 100), 1) if total_machines else 0.0,
        'target_qty': target['target_qty'],
        'target_is_estimated': target['is_estimated'],
        'target_progress_pct': round((packed_pcs / target['target_qty'] * 100), 1) if target['target_qty'] else 0.0,
        'target_remaining': max(target['target_qty'] - packed_pcs, 0),
        'waste_sheets_total': waste_sheets_total,
    }


def _resolve_target(date, shift=None):
    qs = DailyTarget.objects.filter(date=date)
    row = qs.filter(shift=shift).first() if shift else qs.filter(shift__isnull=True).first()
    if row:
        return {'target_qty': row.target_qty, 'is_estimated': False}

    # No DailyTarget row set for today — estimate from job cards that actually
    # have printing/packing activity logged today, not the entire historical
    # backlog of released/in-production/completed job cards.
    estimated = JobCard.objects.filter(
        is_active=True, productions__date=date,
    ).distinct().aggregate(total=Sum('order_qty'))['total'] or 0
    return {'target_qty': estimated, 'is_estimated': True}


def get_printing_performance(date=None):
    date = date or _today()
    printing = _printing_qs(date).select_related('job_card', 'machine', 'operator')

    impressions_total = printing.aggregate(total=Sum('impressions'))['total'] or 0
    output_sheets_total = printing.aggregate(total=Sum('output_sheets'))['total'] or 0
    waste_sheets_total = printing.aggregate(total=Sum('waste_sheets'))['total'] or 0
    waste_pct = round((waste_sheets_total / (output_sheets_total + waste_sheets_total) * 100), 1) \
        if (output_sheets_total + waste_sheets_total) else 0.0

    run_time_total = printing.aggregate(total=Sum('run_time'))['total'] or 0.0
    downtime_total = printing.aggregate(total=Sum('downtime_minutes'))['total'] or 0.0

    machines = _machine_leaderboard(printing)
    operators = _operator_leaderboard(printing)

    return {
        'impressions_total': impressions_total,
        'output_sheets_total': output_sheets_total,
        'waste_sheets_total': waste_sheets_total,
        'waste_pct': waste_pct,
        'avg_run_speed': round(impressions_total / (run_time_total / 60.0), 0) if run_time_total else 0,
        'avg_downtime_minutes': round(downtime_total / printing.count(), 1) if printing.count() else 0,
        'top_machines': machines[:5],
        'top_operators': operators[:5],
        'fastest_machine': machines[0]['name'] if machines else None,
        'lowest_waste_machine': min(machines, key=lambda m: m['waste_pct'])['name'] if machines else None,
    }


def _machine_leaderboard(printing_qs):
    by_machine = {}
    for row in printing_qs:
        if not row.machine_id:
            continue
        entry = by_machine.setdefault(row.machine_id, {
            'name': row.machine.name,
            'output_sheets': 0,
            'waste_sheets': 0,
            'impressions': 0,
            'expected_impressions': 0.0,
        })
        entry['output_sheets'] += row.output_sheets or 0
        entry['waste_sheets'] += row.waste_sheets or 0
        entry['impressions'] += row.impressions or 0
        speed = row.machine.standard_impressions_per_hour or 4000
        entry['expected_impressions'] += speed * ((row.run_time or 0) / 60.0)

    leaderboard = []
    for entry in by_machine.values():
        total_sheets = entry['output_sheets'] + entry['waste_sheets']
        efficiency_pct = round((entry['impressions'] / entry['expected_impressions'] * 100), 1) if entry['expected_impressions'] else 0.0
        waste_pct = round((entry['waste_sheets'] / total_sheets * 100), 1) if total_sheets else 0.0
        leaderboard.append({
            'name': entry['name'],
            'output_sheets': entry['output_sheets'],
            'efficiency_pct': efficiency_pct,
            'waste_pct': waste_pct,
        })
    return sorted(leaderboard, key=lambda m: m['efficiency_pct'], reverse=True)


def _operator_leaderboard(printing_qs):
    by_operator = {}
    for row in printing_qs:
        if not row.operator_id:
            continue
        entry = by_operator.setdefault(row.operator_id, {
            'name': row.operator.name,
            'output_sheets': 0,
            'waste_sheets': 0,
            'impressions': 0,
            'expected_impressions': 0.0,
        })
        entry['output_sheets'] += row.output_sheets or 0
        entry['waste_sheets'] += row.waste_sheets or 0
        entry['impressions'] += row.impressions or 0
        speed = row.machine.standard_impressions_per_hour if row.machine else 4000
        entry['expected_impressions'] += speed * ((row.run_time or 0) / 60.0)

    leaderboard = []
    for entry in by_operator.values():
        total_sheets = entry['output_sheets'] + entry['waste_sheets']
        efficiency_pct = round((entry['impressions'] / entry['expected_impressions'] * 100), 1) if entry['expected_impressions'] else 0.0
        waste_pct = round((entry['waste_sheets'] / total_sheets * 100), 1) if total_sheets else 0.0
        leaderboard.append({
            'name': entry['name'],
            'output_sheets': entry['output_sheets'],
            'efficiency_pct': efficiency_pct,
            'waste_pct': waste_pct,
        })
    return sorted(leaderboard, key=lambda m: m['efficiency_pct'], reverse=True)


def _sorter_leaderboard(packing_qs):
    by_sorter = {}
    for row in packing_qs:
        if not row.sorter_id:
            continue
        entry = by_sorter.setdefault(row.sorter_id, {'name': row.sorter.name, 'packed_qty': 0, 'sorting_waste_qty': 0})
        entry['packed_qty'] += row.packing_qty or 0
        entry['sorting_waste_qty'] += row.sorting_waste_qty or 0

    leaderboard = []
    for entry in by_sorter.values():
        total = entry['packed_qty'] + entry['sorting_waste_qty']
        waste_pct = round((entry['sorting_waste_qty'] / total * 100), 1) if total else 0.0
        leaderboard.append({'name': entry['name'], 'packed_qty': entry['packed_qty'], 'waste_pct': waste_pct})
    return sorted(leaderboard, key=lambda s: s['packed_qty'], reverse=True)


def get_packing_performance(date=None):
    date = date or _today()
    packing = _packing_qs(date).select_related('sorter')

    packed_pcs = packing.aggregate(total=Sum('packing_qty'))['total'] or 0
    sorting_waste_pcs = packing.aggregate(total=Sum('sorting_waste_qty'))['total'] or 0
    total = packed_pcs + sorting_waste_pcs
    waste_pct = round((sorting_waste_pcs / total * 100), 1) if total else 0.0

    target = _resolve_target(date)
    sorters = _sorter_leaderboard(packing)

    return {
        'packed_pcs': packed_pcs,
        'sorting_waste_pcs': sorting_waste_pcs,
        'waste_pct': waste_pct,
        'target_qty': target['target_qty'],
        'target_is_estimated': target['is_estimated'],
        'target_progress_pct': round((packed_pcs / target['target_qty'] * 100), 1) if target['target_qty'] else 0.0,
        'top_sorters': sorters[:5],
        'zero_waste_sorter': next((s['name'] for s in sorters if s['waste_pct'] == 0 and s['packed_qty'] > 0), None),
    }


def get_dispatch_performance(date=None):
    date = date or _today()
    dispatch = _dispatch_qs(date)

    dispatch_qty = dispatch.aggregate(total=Sum('dispatch_qty'))['total'] or 0
    dc_count = dispatch.values('dc_no').distinct().count()

    packed_pcs = _packing_qs(date).aggregate(total=Sum('packing_qty'))['total'] or 0
    completion_pct = round((dispatch_qty / packed_pcs * 100), 1) if packed_pcs else 0.0

    ready_jobs = JobCard.objects.filter(
        is_active=True, status__in=('completed', 'closed'),
    ).exclude(dispatch__is_active=True, dispatch__dispatch_date=date).distinct().count()

    return {
        'dispatch_qty': dispatch_qty,
        'dc_count': dc_count,
        'completion_pct': completion_pct,
        'ready_jobs': ready_jobs,
        'note': 'Dispatch is not attributed to an individual dispatcher or shift in this system today, '
                'so this screen shows plant-wide totals only.',
    }


def get_shift_comparison(date=None):
    date = date or _today()
    shifts = []
    for shift_code, shift_label in Production.SHIFT_CHOICES:
        printing = _printing_qs(date).filter(shift=shift_code).select_related('job_card', 'machine')
        packing = _packing_qs(date).filter(shift=shift_code)

        impressions = printing.aggregate(total=Sum('impressions'))['total'] or 0
        output_sheets = printing.aggregate(total=Sum('output_sheets'))['total'] or 0
        waste_sheets = printing.aggregate(total=Sum('waste_sheets'))['total'] or 0
        packed_pcs = packing.aggregate(total=Sum('packing_qty'))['total'] or 0
        sorting_waste = packing.aggregate(total=Sum('sorting_waste_qty'))['total'] or 0

        expected_impressions = 0.0
        for row in printing:
            speed = row.machine.standard_impressions_per_hour if row.machine else 4000
            expected_impressions += speed * ((row.run_time or 0) / 60.0)
        efficiency_pct = round((impressions / expected_impressions * 100), 1) if expected_impressions else 0.0

        waste_total = waste_sheets + sorting_waste
        produced_total = output_sheets + packed_pcs + waste_total
        waste_pct = round((waste_total / produced_total * 100), 1) if produced_total else 0.0

        score = round(efficiency_pct - waste_pct, 1)
        shifts.append({
            'shift': shift_code,
            'label': shift_label,
            'output_sheets': output_sheets,
            'packed_pcs': packed_pcs,
            'efficiency_pct': efficiency_pct,
            'waste_pct': waste_pct,
            'score': score,
        })

    ranked = sorted(shifts, key=lambda s: s['score'], reverse=True)
    return {
        'shifts': shifts,
        'winning_shift': ranked[0]['label'] if ranked and ranked[0]['score'] > 0 else None,
        'needs_improvement_shift': ranked[-1]['label'] if len(ranked) > 1 else None,
    }


def get_machine_leaderboard(date=None):
    date = date or _today()
    printing = _printing_qs(date).select_related('machine')
    leaderboard = _machine_leaderboard(printing)
    return {
        'machines': leaderboard,
        'machine_of_the_day': leaderboard[0]['name'] if leaderboard else None,
    }


def get_operator_leaderboard(date=None):
    date = date or _today()
    printing = _printing_qs(date).select_related('operator', 'machine')
    packing = _packing_qs(date).select_related('sorter')
    return {
        'operators': _operator_leaderboard(printing),
        'sorters': _sorter_leaderboard(packing),
    }


def get_wastage_quality(date=None):
    date = date or _today()
    printing_waste_sheets = _printing_qs(date).aggregate(total=Sum('waste_sheets'))['total'] or 0
    sorting_waste_pcs = _packing_qs(date).aggregate(total=Sum('sorting_waste_qty'))['total'] or 0

    printing_waste_pcs = 0
    for row in _printing_qs(date).select_related('job_card'):
        printing_waste_pcs += (row.waste_sheets or 0) * (row.job_card.ups or 1)

    flagged_jobs = []
    for job in JobCard.objects.filter(is_active=True, productions__date=date).distinct():
        if job.extra_sheets_used > job.tolerance_sheets:
            flagged_jobs.append(job.job_card_no)

    reconciliation_flagged = []
    for job in JobCard.objects.filter(is_active=True, status__in=FINALIZED_STATUSES):
        metrics = compute_job_card_wastage_metrics(job)
        if metrics and metrics['needs_reconciliation_review']:
            reconciliation_flagged.append(job.job_card_no)

    total_process_wastage = printing_waste_pcs + sorting_waste_pcs
    quality_score = max(round(100 - (total_process_wastage / 1000), 1), 0) if total_process_wastage else 100.0

    return {
        'printing_waste_sheets': printing_waste_sheets,
        'printing_waste_pcs': printing_waste_pcs,
        'sorting_waste_pcs': sorting_waste_pcs,
        'total_process_wastage_pcs': total_process_wastage,
        'flagged_job_count': len(flagged_jobs),
        'flagged_jobs': flagged_jobs[:5],
        'jobs_needing_reconciliation': len(reconciliation_flagged),
        'quality_score': quality_score,
    }


def get_target_achievement(date=None):
    date = date or _today()
    target = _resolve_target(date)
    packed_pcs = _packing_qs(date).aggregate(total=Sum('packing_qty'))['total'] or 0

    now = timezone.localtime()
    start_of_day = now.replace(hour=8, minute=0, second=0, microsecond=0)
    hours_elapsed = max((now - start_of_day).total_seconds() / 3600.0, 0.1)
    hourly_rate = packed_pcs / hours_elapsed
    end_of_day = now.replace(hour=20, minute=0, second=0, microsecond=0)
    hours_remaining = max((end_of_day - now).total_seconds() / 3600.0, 0)
    forecast_qty = packed_pcs + (hourly_rate * hours_remaining)

    if target['target_qty'] <= 0:
        prediction = 'No target set'
    elif forecast_qty >= target['target_qty']:
        prediction = 'Likely to Achieve'
    elif forecast_qty >= target['target_qty'] * 0.85:
        prediction = 'At Risk'
    else:
        prediction = 'Behind Schedule'

    return {
        'target_qty': target['target_qty'],
        'target_is_estimated': target['is_estimated'],
        'achieved_qty': packed_pcs,
        'remaining_qty': max(target['target_qty'] - packed_pcs, 0),
        'hourly_rate': round(hourly_rate, 0),
        'forecast_qty': round(forecast_qty, 0),
        'prediction': prediction,
    }


def get_recognition(date=None):
    date = date or _today()
    printing = get_printing_performance(date)
    packing = get_packing_performance(date)
    machines = get_machine_leaderboard(date)

    badges = []
    if printing['top_operators']:
        top_op = printing['top_operators'][0]
        badges.append({'title': 'Employee of the Shift', 'name': top_op['name'], 'detail': f"{top_op['efficiency_pct']}% efficiency"})
    if packing['top_sorters']:
        top_sorter = packing['top_sorters'][0]
        badges.append({'title': 'Best Sorter', 'name': top_sorter['name'], 'detail': f"{top_sorter['packed_qty']} pcs packed"})
    if machines['machine_of_the_day']:
        badges.append({'title': 'Machine of the Day', 'name': machines['machine_of_the_day'], 'detail': ''})
    zero_waste = packing.get('zero_waste_sorter')
    if zero_waste:
        badges.append({'title': 'Zero Waste Performer', 'name': zero_waste, 'detail': ''})

    quotes = [
        "Quality is everyone's responsibility.",
        "Every sheet saved is profit earned.",
        "Small improvements every day create big results.",
        "Today's waste is tomorrow's cost.",
        "Teamwork builds success.",
    ]
    quote_index = date.toordinal() % len(quotes)

    return {
        'badges': badges,
        'motivational_message': quotes[quote_index],
    }


def get_dashboard_data(date=None):
    is_fallback = False
    if date is None:
        date, is_fallback = _effective_date()
    return {
        'generated_at': timezone.localtime().isoformat(),
        'display_date': date.isoformat(),
        'is_fallback_date': is_fallback,
        'plant_overview': get_plant_overview(date),
        'period_summary': get_period_summary(),
        'printing_performance': get_printing_performance(date),
        'packing_performance': get_packing_performance(date),
        'dispatch_performance': get_dispatch_performance(date),
        'shift_comparison': get_shift_comparison(date),
        'machine_leaderboard': get_machine_leaderboard(date),
        'operator_leaderboard': get_operator_leaderboard(date),
        'wastage_quality': get_wastage_quality(date),
        'target_achievement': get_target_achievement(date),
        'recognition': get_recognition(date),
    }
