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

MIN_SHIFTS_TO_RANK = 5
INDEX_CAP = 150.0
WEIGHTS = (('speed_index', 0.4), ('volume_index', 0.3), ('quality_index', 0.3))

PERIOD_LABELS = {
    'this_month': 'This month', 'last_month': 'Last month', 'last_7': 'Last 7 days', 'last_30': 'Last 30 days',
    'last_90': 'Last 90 days', 'custom': 'Custom dates',
}

ROW_HEADERS = [
    'rank', 'operator', 'machines', 'shifts', 'jobs', 'impressions', 'run_hours', 'impressions_per_hour',
    'impressions_per_shift', 'waste_pct', 'downtime_hours', 'utilisation_pct', 'make_ready_min_per_job', 'score',
]
ROW_LABELS = {
    'rank': 'Rank', 'operator': 'Operator', 'machines': 'Machines', 'shifts': 'Shifts', 'jobs': 'Jobs',
    'impressions': 'Impressions', 'run_hours': 'Run hours', 'impressions_per_hour': 'Impressions / run hour',
    'impressions_per_shift': 'Impressions / shift', 'waste_pct': 'Waste %', 'downtime_hours': 'Downtime (h)',
    'utilisation_pct': 'Utilisation %', 'make_ready_min_per_job': 'Make-ready min / job', 'score': 'Score (100 = peer average)',
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
    bench = defaultdict(lambda: {'imp': 0, 'hours': 0.0, 'sheets': 0, 'waste': 0, 'shifts': set()})
    ops = {}
    for e in entries:
        m, o = e['machine_id'], e['operator_id']
        imp, hours = int(e['impressions'] or 0), float(e['run_time'] or 0) / 60
        sheets = int(e['output_sheets'] or 0) + int(e['waste_sheets'] or 0)
        b = bench[m]
        b['imp'] += imp
        b['hours'] += hours
        b['sheets'] += sheets
        b['waste'] += int(e['waste_sheets'] or 0)
        b['shifts'].add((o, e['date'], e['shift']))
        op = ops.setdefault(o, {
            'id': o, 'name': e['operator__name'], 'imp': 0, 'hours': 0.0, 'downtime': 0.0, 'mr': 0.0, 'mr_jobs': 0,
            'sheets': 0, 'waste': 0, 'shifts': set(), 'jobs': set(), 'machines': defaultdict(lambda: {'imp': 0, 'hours': 0.0, 'sheets': 0, 'waste': 0, 'shifts': set(), 'name': ''}),
            'expected_imp': 0.0, 'expected_shift_imp': 0.0, 'expected_waste': 0.0,
        })
        op['imp'] += imp
        op['hours'] += hours
        op['downtime'] += float(e['downtime_minutes'] or 0)
        if e['make_ready_time']:
            op['mr'] += float(e['make_ready_time'])
            op['mr_jobs'] += 1
        op['sheets'] += sheets
        op['waste'] += int(e['waste_sheets'] or 0)
        op['shifts'].add((e['date'], e['shift']))
        op['jobs'].add(e['job_card_id'])
        pm = op['machines'][m]
        pm['name'] = e['machine__name'] or '-'
        pm['imp'] += imp
        pm['hours'] += hours
        pm['sheets'] += sheets
        pm['waste'] += int(e['waste_sheets'] or 0)
        pm['shifts'].add((e['date'], e['shift']))

    rate = {m: (b['imp'] / b['hours']) if b['hours'] else 0 for m, b in bench.items()}
    per_shift = {m: (b['imp'] / len(b['shifts'])) if b['shifts'] else 0 for m, b in bench.items()}
    waste_rate = {m: (b['waste'] / b['sheets']) if b['sheets'] else 0 for m, b in bench.items()}

    rows = []
    for op in ops.values():
        for m, pm in op['machines'].items():
            op['expected_imp'] += pm['hours'] * rate[m]
            op['expected_shift_imp'] += len(pm['shifts']) * per_shift[m]
            op['expected_waste'] += pm['sheets'] * waste_rate[m]
        shifts = len(op['shifts'])
        speed_index = _clamp(op['imp'] / op['expected_imp'] * 100) if op['expected_imp'] else 0.0
        volume_index = _clamp(op['imp'] / op['expected_shift_imp'] * 100) if op['expected_shift_imp'] else 0.0
        if op['expected_waste'] <= 0:
            quality_index = 100.0
        elif op['waste'] == 0:
            quality_index = INDEX_CAP
        else:
            quality_index = _clamp(op['expected_waste'] / op['waste'] * 100)
        indices = {'speed_index': speed_index, 'volume_index': volume_index, 'quality_index': quality_index}
        score = sum(indices[key] * weight for key, weight in WEIGHTS)
        available = op['hours'] * 60 + op['downtime'] + op['mr']
        rows.append({
            'operator_id': op['id'], 'operator': op['name'], 'shifts': shifts, 'jobs': len(op['jobs']),
            'machines': ', '.join(sorted({pm['name'] for pm in op['machines'].values()})),
            'impressions': op['imp'], 'run_hours': round(op['hours'], 1),
            'impressions_per_hour': round(op['imp'] / op['hours']) if op['hours'] else 0,
            'impressions_per_shift': round(op['imp'] / shifts) if shifts else 0,
            'waste_pct': round(op['waste'] * 100 / op['sheets'], 2) if op['sheets'] else 0.0,
            'downtime_hours': round(op['downtime'] / 60, 1),
            'utilisation_pct': round(op['hours'] * 60 * 100 / available, 1) if available else 0.0,
            'make_ready_min_per_job': round(op['mr'] / op['mr_jobs']) if op['mr_jobs'] else 0,
            'speed_index': round(speed_index), 'volume_index': round(volume_index), 'quality_index': round(quality_index),
            'score': round(score, 1), 'ranked': shifts >= MIN_SHIFTS_TO_RANK,
        })
    ranked = sorted((r for r in rows if r['ranked']), key=lambda r: (-r['score'], -r['impressions']))
    for n, r in enumerate(ranked, start=1):
        r['rank'] = n
    unranked = sorted((r for r in rows if not r['ranked']), key=lambda r: -r['impressions'])
    for r in unranked:
        r['rank'] = ''
    machines = [{
        'name': ops_m_name, 'impressions': b['imp'], 'run_hours': round(b['hours'], 1),
        'impressions_per_hour': round(rate[m]), 'impressions_per_shift': round(per_shift[m]),
        'waste_pct': round(waste_rate[m] * 100, 2),
    } for m, b in bench.items() for ops_m_name in [next((e['machine__name'] for e in entries if e['machine_id'] == m), '-')]]
    machines.sort(key=lambda r: r['name'])
    return {'rows': ranked + unranked, 'ranked': ranked, 'machines': machines, 'entries': entries}


def champions(ranked):
    if not ranked:
        return []
    out = []
    best = max(ranked, key=lambda r: r['speed_index'])
    out.append({'title': 'Fastest on the press', 'name': best['operator'], 'detail': f"{best['impressions_per_hour']:,} impressions / run hour"})
    best = max(ranked, key=lambda r: r['impressions'])
    out.append({'title': 'Most impressions', 'name': best['operator'], 'detail': f"{best['impressions']:,} impressions"})
    clean = [r for r in ranked if r['impressions'] >= 0.5 * (sum(x['impressions'] for x in ranked) / len(ranked))]
    if clean:
        best = min(clean, key=lambda r: r['waste_pct'])
        out.append({'title': 'Cleanest prints', 'name': best['operator'], 'detail': f"{best['waste_pct']}% waste"})
    best = max(ranked, key=lambda r: r['utilisation_pct'])
    out.append({'title': 'Least lost time', 'name': best['operator'], 'detail': f"{best['utilisation_pct']}% of the time running"})
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
    return {
        'filters': {'period': period, 'date_from': start.isoformat(), 'date_to': end.isoformat(), 'machine': machine, 'shift': shift},
        'period_label': PERIOD_LABELS[period], 'start': start, 'end': end,
        'headers': ROW_HEADERS, 'header_labels': ROW_LABELS,
        'rows': result['rows'], 'export_rows': result['rows'], 'top3': ranked[:3], 'champions': champions(ranked),
        'machines': result['machines'], 'machine_names': list(Machine.objects.filter(is_active=True, production__entry_type='printing').distinct().order_by('name').values_list('name', flat=True)),
        'period_choices': list(PERIOD_LABELS.items()),
        'summary': {'operators': len(result['rows']), 'ranked': len(ranked), 'impressions': total_imp,
                    'shifts': sum(r['shifts'] for r in result['rows'])},
        'min_shifts': MIN_SHIFTS_TO_RANK,
        'detail': operator_detail(result['entries'], operator_id) if operator_id else None, 'operator_id': operator_id,
        'export_notes': [f'Period: {start:%d %b %Y} to {end:%d %b %Y}' + (f' | Machine: {machine}' if machine else '') + (f' | Shift {shift}' if shift else ''),
                         'Score: 40% speed + 30% impressions per shift + 30% waste, each against operators on the same machines (100 = peer average).'],
    }
