"""Operator efficiency and ranking: how much each press operator prints, how fast, and how cleanly.

Built from the printing entries (one per operator per job per shift; the derived rows of a combined
layout are left out so the operator is credited once with the whole run).

Operators are compared with the people who ran the *same machines*: a GTO 2C operator is not judged
against an SM 74 operator. Every index below reads 100 = the average of those peers.
"""
from __future__ import annotations

import datetime
from collections import defaultdict

from django.utils import timezone

MIN_SHIFTS_TO_RANK = 5          # never ranked below this many shifts, whatever the period
QUALIFY_PCT = 50                # ...and never with less than this % of the shifts of the operators who turned up most
SHIFT_MINUTES = 720             # a 12 hour shift
INDEX_CAP = 150.0
WEIGHTS = (('speed_index', 0.30), ('throughput_index', 0.25), ('quality_index', 0.20), ('changeover_index', 0.15), ('running_index', 0.10))

PERIOD_LABELS = {
    'this_month': 'This month', 'last_month': 'Last month', 'last_7': 'Last 7 days', 'last_30': 'Last 30 days',
    'last_90': 'Last 90 days', 'custom': 'Custom dates',
}

ROW_HEADERS = [
    'rank', 'operator', 'machines', 'days', 'shifts', 'attendance_pct', 'jobs', 'changeovers', 'impressions', 'run_hours',
    'run_hours_per_shift', 'impressions_per_hour', 'impressions_per_shift', 'waste_pct', 'downtime_hours', 'running_pct',
    'make_ready_min_per_changeover', 'score',
]
ROW_LABELS = {
    'rank': 'Rank', 'operator': 'Operator', 'machines': 'Machines', 'days': 'Days worked', 'shifts': 'Shifts',
    'attendance_pct': 'Attendance %', 'jobs': 'Jobs', 'changeovers': 'Changeovers', 'impressions': 'Impressions',
    'run_hours': 'Run hours', 'run_hours_per_shift': 'Run h / shift', 'impressions_per_hour': 'Impressions / run hour',
    'impressions_per_shift': 'Impressions / shift', 'waste_pct': 'Waste %', 'downtime_hours': 'Downtime (h)',
    'running_pct': 'Press running % of shift', 'make_ready_min_per_changeover': 'Make-ready min / changeover',
    'score': 'Score (100 = peer average)',
}


def resolve_period(period, date_from, date_to, today=None):
    today = today or timezone.localdate()
    period = period if period in PERIOD_LABELS else 'last_30'
    if period == 'custom':
        try:
            start = datetime.date.fromisoformat(date_from)
            end = datetime.date.fromisoformat(date_to) if date_to else today
            if start <= end:
                return period, start, end
        except (TypeError, ValueError):
            pass
        period = 'last_30'
    if period == 'this_month':
        return period, today.replace(day=1), today
    if period == 'last_month':
        end = today.replace(day=1) - datetime.timedelta(days=1)
        return period, end.replace(day=1), end
    days = {'last_7': 7, 'last_30': 30, 'last_90': 90}[period]
    return period, today - datetime.timedelta(days=days - 1), today


def _entries(start, end, machine='', shift=''):
    from core.models import Production

    qs = Production.objects.filter(
        entry_type='printing', is_active=True, merge_parent__isnull=True, operator__isnull=False,
        date__gte=start, date__lte=end,
    )
    if machine:
        qs = qs.filter(machine__name=machine)
    if shift in ('A', 'B'):
        qs = qs.filter(shift=shift)
    return list(qs.values(
        'id', 'operator_id', 'operator__name', 'machine_id', 'machine__name', 'date', 'shift', 'job_card_id',
        'job_card__job_card_no', 'job_card__SKU', 'impressions', 'run_time', 'downtime_minutes', 'make_ready_time',
        'waste_sheets', 'output_sheets',
    ))


def _clamp(value, high=INDEX_CAP):
    return max(0.0, min(value, high))


def compute(start, end, machine='', shift=''):
    """Ranked operator rows plus machine benchmarks for the period."""
    entries = _entries(start, end, machine, shift)
    bench = defaultdict(lambda: {'imp': 0, 'run_min': 0.0, 'sheets': 0, 'waste': 0, 'shifts': set(), 'changeovers': set(), 'mr': 0.0, 'mr_n': 0, 'name': ''})
    ops = {}
    for e in entries:
        m, o = e['machine_id'], e['operator_id']
        imp, run_min = int(e['impressions'] or 0), float(e['run_time'] or 0)
        sheets = int(e['output_sheets'] or 0) + int(e['waste_sheets'] or 0)
        slot = (o, e['date'], e['shift'])
        change = (o, e['date'], e['shift'], m, e['job_card_id'])               # one make-ready: this operator starting this job on this shift
        b = bench[m]
        b['name'] = e['machine__name'] or '-'
        b['imp'] += imp
        b['run_min'] += run_min
        b['sheets'] += sheets
        b['waste'] += int(e['waste_sheets'] or 0)
        b['shifts'].add(slot)
        b['changeovers'].add(change)
        if e['make_ready_time']:
            b['mr'] += float(e['make_ready_time'])
            b['mr_n'] += 1
        op = ops.setdefault(o, {
            'id': o, 'name': e['operator__name'], 'imp': 0, 'run_min': 0.0, 'downtime': 0.0, 'mr': 0.0, 'mr_n': 0, 'sheets': 0, 'waste': 0,
            'shifts': set(), 'days': set(), 'jobs': set(), 'changeovers': set(), 'machines': defaultdict(lambda: {'imp': 0, 'run_min': 0.0, 'sheets': 0, 'waste': 0, 'shifts': set(), 'mr': 0.0, 'mr_n': 0}),
        })
        op['imp'] += imp
        op['run_min'] += run_min
        op['downtime'] += float(e['downtime_minutes'] or 0)
        op['sheets'] += sheets
        op['waste'] += int(e['waste_sheets'] or 0)
        op['shifts'].add((e['date'], e['shift']))
        op['days'].add(e['date'])
        op['jobs'].add(e['job_card_id'])
        op['changeovers'].add(change)
        pm = op['machines'][m]
        pm['imp'] += imp
        pm['run_min'] += run_min
        pm['sheets'] += sheets
        pm['waste'] += int(e['waste_sheets'] or 0)
        pm['shifts'].add((e['date'], e['shift']))
        if e['make_ready_time']:
            op['mr'] += float(e['make_ready_time'])
            op['mr_n'] += 1
            pm['mr'] += float(e['make_ready_time'])
            pm['mr_n'] += 1

    # what an average operator achieves on each machine (the yardstick for every index)
    rate = {m: (b['imp'] / (b['run_min'] / 60)) if b['run_min'] else 0 for m, b in bench.items()}
    per_shift = {m: (b['imp'] / len(b['shifts'])) if b['shifts'] else 0 for m, b in bench.items()}
    waste_rate = {m: (b['waste'] / b['sheets']) if b['sheets'] else 0 for m, b in bench.items()}
    run_per_shift = {m: (b['run_min'] / len(b['shifts'])) if b['shifts'] else 0 for m, b in bench.items()}
    mr_per_change = {m: (b['mr'] / b['mr_n']) if b['mr_n'] else 0 for m, b in bench.items()}

    shift_counts = sorted((len(op['shifts']) for op in ops.values()), reverse=True)
    top_shifts = shift_counts[min(2, len(shift_counts) - 1)] if shift_counts else 0     # 3rd-most-present operator: one outlier cannot raise the bar
    need_shifts = max(MIN_SHIFTS_TO_RANK, int(top_shifts * QUALIFY_PCT / 100 + 0.999))

    rows = []
    for op in ops.values():
        exp_imp = exp_shift_imp = exp_waste = exp_run = exp_mr = 0.0
        for m, pm in op['machines'].items():
            exp_imp += pm['run_min'] / 60 * rate[m]
            exp_shift_imp += len(pm['shifts']) * per_shift[m]
            exp_waste += pm['sheets'] * waste_rate[m]
            exp_run += len(pm['shifts']) * run_per_shift[m]
            exp_mr += pm['mr_n'] * mr_per_change[m]
        shifts = len(op['shifts'])
        speed = _clamp(op['imp'] / exp_imp * 100) if exp_imp else 100.0
        throughput = _clamp(op['imp'] / exp_shift_imp * 100) if exp_shift_imp else 100.0
        running = _clamp(op['run_min'] / exp_run * 100) if exp_run else 100.0
        if exp_waste <= 0:
            quality = 100.0
        elif op['waste'] == 0:
            quality = INDEX_CAP
        else:
            quality = _clamp(exp_waste / op['waste'] * 100)
        changeover = (INDEX_CAP if op['mr'] == 0 else _clamp(exp_mr / op['mr'] * 100)) if exp_mr else 100.0
        indices = {'speed_index': speed, 'throughput_index': throughput, 'quality_index': quality, 'changeover_index': changeover, 'running_index': running}
        score = sum(indices[key] * weight for key, weight in WEIGHTS)
        changeovers = len(op['changeovers'])
        rows.append({
            'operator_id': op['id'], 'operator': op['name'], 'shifts': shifts, 'days': len(op['days']), 'jobs': len(op['jobs']),
            'machines': ', '.join(sorted({pm_name for pm_name in (bench[m]['name'] for m in op['machines'])})),
            'attendance_pct': round(shifts * 100 / top_shifts) if top_shifts else 0, 'changeovers': changeovers,
            'impressions': op['imp'], 'run_hours': round(op['run_min'] / 60, 1),
            'run_hours_per_shift': round(op['run_min'] / 60 / shifts, 1) if shifts else 0.0,
            'impressions_per_hour': round(op['imp'] / (op['run_min'] / 60)) if op['run_min'] else 0,
            'impressions_per_shift': round(op['imp'] / shifts) if shifts else 0,
            'waste_pct': round(op['waste'] * 100 / op['sheets'], 2) if op['sheets'] else 0.0,
            'downtime_hours': round(op['downtime'] / 60, 1),
            'running_pct': round(op['run_min'] * 100 / (shifts * SHIFT_MINUTES), 1) if shifts else 0.0,
            'make_ready_min_per_changeover': round(op['mr'] / op['mr_n']) if op['mr_n'] else 0,
            'speed_index': round(speed), 'throughput_index': round(throughput), 'quality_index': round(quality),
            'changeover_index': round(changeover), 'running_index': round(running),
            'score': round(score, 1), 'ranked': shifts >= need_shifts,
        })
    ranked = sorted((r for r in rows if r['ranked']), key=lambda r: (-r['score'], -r['impressions']))
    for n, r in enumerate(ranked, start=1):
        r['rank'] = n
    unranked = sorted((r for r in rows if not r['ranked']), key=lambda r: -r['impressions'])
    for r in unranked:
        r['rank'] = ''
    machines = sorted(({
        'name': b['name'], 'impressions': b['imp'], 'run_hours': round(b['run_min'] / 60, 1),
        'impressions_per_hour': round(rate[m]), 'impressions_per_shift': round(per_shift[m]),
        'waste_pct': round(waste_rate[m] * 100, 2), 'make_ready_min': round(mr_per_change[m]),
    } for m, b in bench.items()), key=lambda r: r['name'])
    return {'rows': ranked + unranked, 'ranked': ranked, 'machines': machines, 'entries': entries, 'need_shifts': need_shifts, 'top_shifts': top_shifts}


def champions(ranked):
    """Per-measure leaders among the ranked (qualified) operators only."""
    if not ranked:
        return []
    out = []
    best = max(ranked, key=lambda r: r['speed_index'])
    out.append({'title': 'Fastest on the press', 'name': best['operator'], 'detail': f"{best['impressions_per_hour']:,} impressions / run hour"})
    best = max(ranked, key=lambda r: r['impressions'])
    out.append({'title': 'Most impressions', 'name': best['operator'], 'detail': f"{best['impressions']:,} in {best['shifts']} shifts"})
    best = max(ranked, key=lambda r: r['impressions_per_shift'])
    out.append({'title': 'Best per shift', 'name': best['operator'], 'detail': f"{best['impressions_per_shift']:,} impressions / shift"})
    clean = [r for r in ranked if r['impressions'] >= 0.5 * (sum(x['impressions'] for x in ranked) / len(ranked))]
    if clean:
        best = min(clean, key=lambda r: r['waste_pct'])
        out.append({'title': 'Cleanest prints', 'name': best['operator'], 'detail': f"{best['waste_pct']}% waste"})
    timed = [r for r in ranked if r['make_ready_min_per_changeover']]
    if timed:
        best = min(timed, key=lambda r: r['make_ready_min_per_changeover'])
        out.append({'title': 'Quickest changeovers', 'name': best['operator'], 'detail': f"{best['make_ready_min_per_changeover']} min make-ready, {best['changeovers']} changeovers"})
    best = max(ranked, key=lambda r: r['running_pct'])
    out.append({'title': 'Press kept running', 'name': best['operator'], 'detail': f"{best['running_pct']}% of the shift"})
    return out


def operator_detail(entries, operator_id):
    """Machine and day breakdown plus the job list for one operator, from the period's entries."""
    mine = [e for e in entries if e['operator_id'] == operator_id]
    if not mine:
        return None
    by_machine, by_day, jobs = defaultdict(lambda: [0, 0.0, 0, 0, set()]), defaultdict(lambda: [0, 0.0, set()]), defaultdict(lambda: [0, 0.0, 0, 0])
    for e in mine:
        imp, hours = int(e['impressions'] or 0), float(e['run_time'] or 0) / 60
        a = by_machine[e['machine__name'] or '-']
        a[0] += imp
        a[1] += hours
        a[2] += int(e['waste_sheets'] or 0)
        a[3] += int(e['output_sheets'] or 0) + int(e['waste_sheets'] or 0)
        a[4].add((e['date'], e['shift']))
        d = by_day[e['date']]
        d[0] += imp
        d[1] += hours
        d[2].add(e['shift'])
        j = jobs[(e['job_card_id'], e['job_card__job_card_no'], e['job_card__SKU'])]
        j[0] += imp
        j[1] += hours
        j[2] += int(e['waste_sheets'] or 0)
        j[3] += int(e['output_sheets'] or 0) + int(e['waste_sheets'] or 0)
    from core.models import JobCard
    from reports.sku_history import _detail_url

    cards = {c.id: c for c in JobCard.objects.filter(id__in={e['job_card_id'] for e in mine}).select_related('planning_job')}
    return {
        'name': mine[0]['operator__name'],
        'machines': [{'name': k, 'impressions': v[0], 'run_hours': round(v[1], 1), 'shifts': len(v[4]),
                      'impressions_per_hour': round(v[0] / v[1]) if v[1] else 0,
                      'waste_pct': round(v[2] * 100 / v[3], 2) if v[3] else 0.0} for k, v in sorted(by_machine.items())],
        'days': [{'date': k, 'impressions': v[0], 'run_hours': round(v[1], 1), 'shifts': ''.join(sorted(v[2])),
                  'impressions_per_hour': round(v[0] / v[1]) if v[1] else 0} for k, v in sorted(by_day.items(), reverse=True)],
        'jobs': [{'job_card_id': k[0], 'jc': k[1], 'sku': k[2], 'url': _detail_url(cards.get(k[0]), getattr(cards.get(k[0]), 'planning_job', None)), 'impressions': v[0], 'run_hours': round(v[1], 1),
                  'impressions_per_hour': round(v[0] / v[1]) if v[1] else 0,
                  'waste_pct': round(v[2] * 100 / v[3], 2) if v[3] else 0.0}
                 for k, v in sorted(jobs.items(), key=lambda kv: -kv[1][0])],
    }


def build_context(request):
    """Report context for the page, the exports and the notice-board print."""
    q = request.GET
    period, start, end = resolve_period((q.get('period') or '').strip(), (q.get('date_from') or '').strip(), (q.get('date_to') or '').strip())
    machine = (q.get('machine') or '').strip()
    shift = (q.get('shift') or '').strip().upper()
    result = compute(start, end, machine, shift)
    ranked = result['ranked']
    try:
        operator_id = int(q.get('operator') or 0)
    except ValueError:
        operator_id = 0
    from core.models import Machine

    total_imp = sum(r['impressions'] for r in result['rows'])
    from django.db.models import Count, ExpressionWrapper, F, IntegerField, Sum

    from coating.models import CoatingRun

    coating_qs = CoatingRun.objects.filter(is_active=True, date__gte=start, date__lte=end)
    if machine:
        coating_qs = coating_qs.filter(machine__name=machine)
    if shift in ('A', 'B'):
        coating_qs = coating_qs.filter(shift=shift)
    coating = list(
        coating_qs.values('operator__name').annotate(
            runs=Count('id'), sheets=Sum('sheets_coated'), waste=Sum('waste_sheets'), minutes=Sum('run_minutes'),
            impressions=Sum(ExpressionWrapper((F('sheets_coated') + F('waste_sheets')) * F('coats'), output_field=IntegerField())),
        ).order_by('-impressions')
    )
    for row in coating:
        row['operator'] = row.pop('operator__name') or 'No operator'
        row['run_hours'] = round((row['minutes'] or 0) / 60, 1)
        row['per_hour'] = round((row['impressions'] or 0) / row['run_hours']) if row['run_hours'] else ''
    return {
        'coating': coating,
        'filters': {'period': period, 'date_from': start.isoformat(), 'date_to': end.isoformat(), 'machine': machine, 'shift': shift},
        'period_label': PERIOD_LABELS[period], 'start': start, 'end': end,
        'headers': ROW_HEADERS, 'header_labels': ROW_LABELS,
        'rows': result['rows'], 'export_rows': result['rows'], 'top3': ranked[:3], 'champions': champions(ranked),
        'machines': result['machines'], 'machine_names': list(Machine.objects.filter(is_active=True, production__entry_type='printing').distinct().order_by('name').values_list('name', flat=True)),
        'period_choices': list(PERIOD_LABELS.items()),
        'summary': {'operators': len(result['rows']), 'ranked': len(ranked), 'impressions': total_imp,
                    'shifts': sum(r['shifts'] for r in result['rows'])},
        'min_shifts': result['need_shifts'], 'top_shifts': result['top_shifts'], 'qualify_pct': QUALIFY_PCT,
        'detail': operator_detail(result['entries'], operator_id) if operator_id else None, 'operator_id': operator_id,
        'export_notes': [f'Period: {start:%d %b %Y} to {end:%d %b %Y}' + (f' | Machine: {machine}' if machine else '') + (f' | Shift {shift}' if shift else ''),
                         'Score: 30% speed + 25% impressions per shift + 20% waste + 15% changeover time + 10% press running, each against operators on the same machines (100 = peer average). Ranked only with enough shifts.'],
    }
