from __future__ import annotations

import calendar
from datetime import date

from django.db.models import ExpressionWrapper, F, FloatField, Prefetch, Sum, Value
from django.db.models.functions import Coalesce
from django.utils import timezone

from core.models import Dispatch, JobCard, Production
from planning.models import PlanningJob

from .models import KPIActionNote, KPITarget

MONTH_NAMES = [calendar.month_name[m] for m in range(1, 13)]


def _month_bounds(year, month):
    start = date(year, month, 1)
    end = date(year, month, calendar.monthrange(year, month)[1])
    return start, end


def _quarter_bounds(year, quarter):
    first_month = (quarter - 1) * 3 + 1
    start = date(year, first_month, 1)
    end_month = first_month + 2
    end = date(year, end_month, calendar.monthrange(year, end_month)[1])
    return start, end


def _order_qty_for_period(start, end):
    total = PlanningJob.objects.filter(
        is_active=True, po_approval_date__range=(start, end),
    ).aggregate(total=Sum('order_qty'))['total']
    return total or 0


def _dispatched_pcs_for_period(start, end):
    total = Dispatch.objects.filter(
        is_active=True, dispatch_date__range=(start, end),
    ).aggregate(total=Sum('dispatch_qty'))['total']
    return total or 0


def _packed_pcs_for_period(start, end):
    total = Production.objects.filter(
        is_active=True, entry_type='packing', date__range=(start, end),
    ).aggregate(total=Sum('packing_qty'))['total']
    return total or 0


def _wastage_pcs_for_period(start, end):
    waste_pcs_expr = ExpressionWrapper(
        Coalesce(F('waste_sheets'), Value(0)) * Coalesce(F('job_card__ups'), Value(1)),
        output_field=FloatField(),
    )
    printing_waste = Production.objects.filter(
        is_active=True, entry_type='printing', date__range=(start, end),
    ).aggregate(total=Sum(waste_pcs_expr))['total'] or 0.0
    sorting_waste = Production.objects.filter(
        is_active=True, entry_type='packing', date__range=(start, end),
    ).aggregate(total=Sum('sorting_waste_qty'))['total'] or 0
    return printing_waste + sorting_waste


def compute_order_fulfillment(start, end):
    order_qty = _order_qty_for_period(start, end)
    dispatched = _dispatched_pcs_for_period(start, end)
    value = round((dispatched / order_qty) * 100, 2) if order_qty else 0.0
    return value, {'order_qty': order_qty, 'dispatched_pcs': dispatched}


def compute_wastage_reduction(start, end):
    order_qty = _order_qty_for_period(start, end)
    wastage_pcs = _wastage_pcs_for_period(start, end)
    value = round((wastage_pcs / order_qty) * 100, 2) if order_qty else 0.0
    return value, {'order_qty': order_qty, 'wastage_pcs': wastage_pcs}


def compute_dispatch_alignment(start, end):
    packed = _packed_pcs_for_period(start, end)
    dispatched = _dispatched_pcs_for_period(start, end)
    value = round((dispatched / packed) * 100, 2) if packed else 0.0
    return value, {'packed_pcs': packed, 'dispatched_pcs': dispatched}


KPI_COMPUTE_FUNCS = {
    KPITarget.KPI_ORDER_FULFILLMENT: compute_order_fulfillment,
    KPITarget.KPI_WASTAGE_REDUCTION: compute_wastage_reduction,
    KPITarget.KPI_DISPATCH_ALIGNMENT: compute_dispatch_alignment,
}

KPI_LABELS = dict(KPITarget.KPI_CHOICES)


def _machine_for(job_card_machine, planning_machine):
    """Mirror JobCard.machine_name_display: the job card's own machine (which a
    production supervisor may update after planning) wins; the planner's
    originally-entered free-text machine name is only a fallback."""
    if job_card_machine:
        return job_card_machine
    return (planning_machine or '').strip()


def _order_rows(start, end):
    return [
        {
            'row_type': 'Order (Planned)', 'job_card_no': pj['jc_number'], 'reference': pj['jc_number'],
            'po_number': pj['po_number'] or '', 'date': pj['po_approval_date'], 'qty_pcs': pj['order_qty'] or 0,
            'machine': _machine_for(pj['job_card__machine_name__name'], pj['machine_name']),
        }
        for pj in PlanningJob.objects.filter(
            is_active=True, po_approval_date__range=(start, end),
        ).values('jc_number', 'po_number', 'order_qty', 'po_approval_date', 'machine_name', 'job_card__machine_name__name')
    ]


def _dispatch_rows(start, end):
    return [
        {
            'row_type': 'Dispatch (Actual)', 'job_card_no': d['job_card__job_card_no'] or '',
            'reference': d['dc_no'], 'po_number': d['job_card__PO_No'] or '',
            'date': d['dispatch_date'], 'qty_pcs': d['dispatch_qty'] or 0,
            'machine': _machine_for(d['job_card__machine_name__name'], d['job_card__planning_job__machine_name']),
        }
        for d in Dispatch.objects.filter(
            is_active=True, dispatch_date__range=(start, end),
        ).values(
            'job_card__job_card_no', 'job_card__PO_No', 'dc_no', 'dispatch_date', 'dispatch_qty',
            'job_card__machine_name__name', 'job_card__planning_job__machine_name',
        )
    ]


def _packing_rows(start, end):
    return [
        {
            'row_type': 'Packed (Production)', 'job_card_no': p['job_card__job_card_no'] or '',
            'reference': '', 'po_number': p['job_card__PO_No'] or '', 'date': p['date'], 'qty_pcs': p['packing_qty'] or 0,
            'machine': _machine_for(p['job_card__machine_name__name'], p['job_card__planning_job__machine_name']),
        }
        for p in Production.objects.filter(
            is_active=True, entry_type='packing', date__range=(start, end),
        ).values(
            'job_card__job_card_no', 'job_card__PO_No', 'date', 'packing_qty',
            'job_card__machine_name__name', 'job_card__planning_job__machine_name',
        )
    ]


def _drilldown_order_fulfillment(start, end):
    return _order_rows(start, end) + _dispatch_rows(start, end)


def _waste_rows(start, end):
    """Every printing/sorting waste entry in the period, as drill-down rows.
    Shared by the full wastage drill-down (below) and the focus/gap export."""
    rows = []
    for p in Production.objects.filter(
        is_active=True, entry_type='printing', date__range=(start, end),
    ).values(
        'job_card__job_card_no', 'job_card__PO_No', 'date', 'waste_sheets', 'job_card__ups',
        'job_card__machine_name__name', 'job_card__planning_job__machine_name',
    ):
        waste_pcs = (p['waste_sheets'] or 0) * (p['job_card__ups'] or 1)
        rows.append({
            'row_type': 'Printing Waste', 'job_card_no': p['job_card__job_card_no'] or '',
            'reference': '', 'po_number': p['job_card__PO_No'] or '', 'date': p['date'], 'qty_pcs': waste_pcs,
            'machine': _machine_for(p['job_card__machine_name__name'], p['job_card__planning_job__machine_name']),
        })
    for p in Production.objects.filter(
        is_active=True, entry_type='packing', date__range=(start, end),
    ).values(
        'job_card__job_card_no', 'job_card__PO_No', 'date', 'sorting_waste_qty',
        'job_card__machine_name__name', 'job_card__planning_job__machine_name',
    ):
        rows.append({
            'row_type': 'Sorting Waste', 'job_card_no': p['job_card__job_card_no'] or '',
            'reference': '', 'po_number': p['job_card__PO_No'] or '', 'date': p['date'],
            'qty_pcs': p['sorting_waste_qty'] or 0,
            'machine': _machine_for(p['job_card__machine_name__name'], p['job_card__planning_job__machine_name']),
        })
    return rows


def _drilldown_wastage_reduction(start, end):
    return _order_rows(start, end) + _waste_rows(start, end)


def _drilldown_dispatch_alignment(start, end):
    return _packing_rows(start, end) + _dispatch_rows(start, end)


DRILLDOWN_FUNCS = {
    KPITarget.KPI_ORDER_FULFILLMENT: _drilldown_order_fulfillment,
    KPITarget.KPI_WASTAGE_REDUCTION: _drilldown_wastage_reduction,
    KPITarget.KPI_DISPATCH_ALIGNMENT: _drilldown_dispatch_alignment,
}

DRILLDOWN_HEADERS = ['kpi', 'row_type', 'job_card_no', 'po_number', 'reference', 'date', 'qty_pcs', 'machine']
DRILLDOWN_HEADER_LABELS = {
    'kpi': 'KPI', 'row_type': 'Row Type', 'job_card_no': 'Job Card', 'po_number': 'PO/WO',
    'reference': 'Reference (DC/JC)', 'date': 'Date', 'qty_pcs': 'Qty (Pcs)', 'machine': 'Machine',
}

# Business-friendly process names for the quarterly-detail export, matching the
# vocabulary used in the hand-built reference KPI spreadsheets.
PROCESS_LABELS = {
    'Order (Planned)': 'Planning',
    'Dispatch (Actual)': 'Dispatch',
    'Packed (Production)': 'Production',
    'Printing Waste': 'Printing Waste',
    'Sorting Waste': 'Sorting Waste',
}

QUARTERLY_DETAIL_HEADERS = [
    'quarter', 'process', 'month', 'date', 'po_number', 'job_card_no', 'sku', 'qty_pcs', 'machine',
]
QUARTERLY_DETAIL_HEADER_LABELS = {
    'quarter': 'Quarter', 'process': 'Process', 'month': 'Month', 'date': 'Date', 'po_number': 'PO/WO',
    'job_card_no': 'JC', 'sku': 'SKU', 'qty_pcs': 'Order Qty (Pcs)', 'machine': 'Machine',
}


def build_kpi_quarterly_detail_rows(kpi_slug, year, quarter):
    """Every underlying row for a KPI across the whole quarter (all 3 months),
    in a layout close to the hand-built reference KPI spreadsheets, plus a
    totals row and the resulting KPI % for the quarter."""
    drilldown_func = DRILLDOWN_FUNCS[kpi_slug]
    quarter_start, quarter_end = _quarter_bounds(year, quarter)
    quarter_label = f'Q{quarter} {year}'

    rows = []
    for month_number in range((quarter - 1) * 3 + 1, (quarter - 1) * 3 + 4):
        m_start, m_end = _month_bounds(year, month_number)
        for drow in drilldown_func(m_start, m_end):
            rows.append({
                'quarter': quarter_label,
                'process': PROCESS_LABELS.get(drow['row_type'], drow['row_type']),
                'month': MONTH_NAMES[month_number - 1],
                'date': drow['date'],
                'po_number': drow.get('po_number', ''),
                'job_card_no': drow['job_card_no'],
                'sku': '',
                'qty_pcs': drow['qty_pcs'],
                'machine': drow.get('machine', ''),
            })

    total_qty = sum(r['qty_pcs'] for r in rows)
    rows.append({
        'quarter': quarter_label, 'process': 'Total', 'month': '', 'date': '', 'po_number': '',
        'job_card_no': '', 'sku': '', 'qty_pcs': total_qty, 'machine': '',
    })

    value, _detail = KPI_COMPUTE_FUNCS[kpi_slug](quarter_start, quarter_end)
    rows.append({
        'quarter': quarter_label, 'process': f'{KPI_LABELS.get(kpi_slug, kpi_slug)} %', 'month': '', 'date': '',
        'po_number': '', 'job_card_no': '', 'sku': '', 'qty_pcs': f'{value}%', 'machine': '',
    })
    return rows


# --- Improvement Focus export -------------------------------------------------
# Not a raw ledger like the drill-downs above: a ranked, job-level "what to fix"
# list for the team, showing exactly which jobs are causing a KPI's shortfall
# and where each one is stuck.

def _focus_order_fulfillment(start, end):
    rows = []
    today = timezone.localdate()
    jobs = PlanningJob.objects.filter(
        is_active=True, po_approval_date__range=(start, end),
    ).select_related('job_card', 'job_card__machine_name').prefetch_related(
        Prefetch('job_card__productions', queryset=Production.objects.filter(is_active=True)),
        Prefetch('job_card__dispatch_set', queryset=Dispatch.objects.filter(is_active=True)),
    )
    for job in jobs:
        order_qty = job.order_qty or 0
        job_card = getattr(job, 'job_card', None)

        if job_card is None:
            gap = order_qty
            if gap <= 0:
                continue
            rows.append({
                'job_card_no': job.jc_number, 'po_number': job.po_number or '', 'sku': job.sku,
                'issue': 'Not Released', 'machine': (job.machine_name or '').strip(), 'qty_pcs': gap,
                'date': '', 'days_pending': (today - job.created_at.date()).days if job.created_at else '',
            })
            continue

        ups = job_card.ups or 1
        printed_pcs = 0
        packed_pcs = 0
        last_activity_date = None
        for p in job_card.productions.all():
            if p.entry_type == 'printing':
                printed_pcs += p.output_sheets * (p.merge_allocated_ups or ups)
            elif p.entry_type == 'packing':
                packed_pcs += p.packing_qty
            if last_activity_date is None or p.date > last_activity_date:
                last_activity_date = p.date

        dispatched_pcs = 0
        for d in job_card.dispatch_set.all():
            dispatched_pcs += d.dispatch_qty
            if last_activity_date is None or d.dispatch_date > last_activity_date:
                last_activity_date = d.dispatch_date

        gap = order_qty - dispatched_pcs
        if gap <= 0:
            continue

        packing_limit_pcs = printed_pcs if job_card.is_print_job else order_qty
        pending_printing = max(order_qty - printed_pcs, 0) if job_card.is_print_job else 0
        pending_packing = max(packing_limit_pcs - packed_pcs, 0)
        pending_dispatch = max(packed_pcs - dispatched_pcs, 0)
        if pending_printing:
            issue = 'Stuck at Printing'
        elif pending_packing:
            issue = 'Stuck at Packing'
        elif pending_dispatch:
            issue = 'Awaiting Dispatch'
        else:
            issue = 'Pending Completion'

        if last_activity_date is None and job_card.created_at:
            last_activity_date = job_card.created_at.date()
        days_pending = (today - last_activity_date).days if last_activity_date else ''

        rows.append({
            'job_card_no': job_card.job_card_no, 'po_number': job_card.PO_No or '', 'sku': job_card.SKU,
            'issue': issue, 'machine': job_card.machine_name_display, 'qty_pcs': gap,
            'date': '', 'days_pending': days_pending,
        })
    return rows


def _focus_dispatch_alignment(start, end):
    rows = []
    today = timezone.localdate()
    job_cards = JobCard.objects.filter(
        is_active=True, productions__entry_type='packing', productions__date__range=(start, end),
    ).distinct().select_related('machine_name', 'planning_job').prefetch_related(
        Prefetch('productions', queryset=Production.objects.filter(is_active=True)),
        Prefetch('dispatch_set', queryset=Dispatch.objects.filter(is_active=True)),
    )
    for job_card in job_cards:
        packed_pcs = 0
        last_packing_date = None
        for p in job_card.productions.all():
            if p.entry_type != 'packing':
                continue
            packed_pcs += p.packing_qty
            if last_packing_date is None or p.date > last_packing_date:
                last_packing_date = p.date

        dispatched_pcs = sum(d.dispatch_qty for d in job_card.dispatch_set.all())
        gap = packed_pcs - dispatched_pcs
        if gap <= 0:
            continue

        days_pending = (today - last_packing_date).days if last_packing_date else ''
        rows.append({
            'job_card_no': job_card.job_card_no, 'po_number': job_card.PO_No or '', 'sku': job_card.SKU,
            'issue': 'Awaiting Dispatch', 'machine': job_card.machine_name_display, 'qty_pcs': gap,
            'date': '', 'days_pending': days_pending,
        })
    return rows


def _focus_wastage_reduction(start, end):
    rows = []
    for drow in _waste_rows(start, end):
        if not drow['qty_pcs']:
            continue
        rows.append({
            'job_card_no': drow['job_card_no'], 'po_number': drow['po_number'], 'sku': '',
            'issue': drow['row_type'], 'machine': drow['machine'], 'qty_pcs': drow['qty_pcs'],
            'date': drow['date'], 'days_pending': '',
        })
    return rows


_FOCUS_FUNCS = {
    KPITarget.KPI_ORDER_FULFILLMENT: _focus_order_fulfillment,
    KPITarget.KPI_WASTAGE_REDUCTION: _focus_wastage_reduction,
    KPITarget.KPI_DISPATCH_ALIGNMENT: _focus_dispatch_alignment,
}


def build_kpi_focus_rows(kpi_slug, start, end):
    """Ranked, job-level list of exactly what's dragging a KPI down for the
    period — worst gap/waste first — meant to be handed to the team as an
    actionable to-do list, as opposed to the raw row dumps above."""
    focus_func = _FOCUS_FUNCS.get(kpi_slug)
    if focus_func is None:
        return []
    rows = focus_func(start, end)
    rows.sort(key=lambda r: r['qty_pcs'], reverse=True)
    return rows


KPI_FOCUS_HEADERS = ['job_card_no', 'po_number', 'sku', 'issue', 'machine', 'qty_pcs', 'date', 'days_pending']
KPI_FOCUS_HEADER_LABELS = {
    'job_card_no': 'Job Card', 'po_number': 'PO/WO', 'sku': 'SKU', 'issue': 'Issue / Stuck At',
    'machine': 'Machine', 'qty_pcs': 'Qty (Pcs)', 'date': 'Date', 'days_pending': 'Days Pending',
}


def _status_for(target, value):
    """Red/Yellow/Green banding against a KPITarget's Min/Target/Max range."""
    min_v, target_v, max_v = float(target.min_value), float(target.target_value), float(target.max_value)
    if target.higher_is_better:
        if value > max_v:
            return 'yellow'  # over-shoot caution (e.g. over-dispatching ahead of production)
        if value >= target_v:
            return 'green'
        if value >= min_v:
            return 'yellow'
        return 'red'
    if value <= target_v:
        return 'green'
    if value <= max_v:
        return 'yellow'
    return 'red'


SUGGESTED_ACTIONS = {
    (KPITarget.KPI_ORDER_FULFILLMENT, 'red'): (
        'Fulfillment is {value}%, below the {min}% floor — review jobs with dispatch pending against '
        'this period\'s orders and prioritize clearing the oldest backlog first.'
    ),
    (KPITarget.KPI_ORDER_FULFILLMENT, 'yellow'): (
        'Fulfillment is {value}%, short of the {target}% target — check the Pending Work report for '
        'jobs stuck at packing/dispatch this period.'
    ),
    (KPITarget.KPI_ORDER_FULFILLMENT, 'green'): (
        'Fulfillment is {value}%, at or above target ({target}%) — no action needed.'
    ),
    (KPITarget.KPI_WASTAGE_REDUCTION, 'red'): (
        'Wastage is {value}%, above the {max}% ceiling — review setup/running/rejection waste on the '
        'highest-wastage machines and jobs this period.'
    ),
    (KPITarget.KPI_WASTAGE_REDUCTION, 'yellow'): (
        'Wastage is {value}%, above the {target}% target but within the {max}% ceiling — monitor the '
        'top wastage contributors this period.'
    ),
    (KPITarget.KPI_WASTAGE_REDUCTION, 'green'): (
        'Wastage is {value}%, at or below target ({target}%) — no action needed.'
    ),
    (KPITarget.KPI_DISPATCH_ALIGNMENT, 'red'): (
        'Alignment is {value}%, below the {min}% floor — dispatch is falling behind production; check '
        'for a dispatch backlog or blocked shipments.'
    ),
    (KPITarget.KPI_DISPATCH_ALIGNMENT, 'yellow'): (
        'Alignment is {value}%, outside the {min}–{max}% band — if above {max}%, confirm dispatch '
        'isn\'t depleting stock ahead of production; if below {target}%, check for a dispatch backlog.'
    ),
    (KPITarget.KPI_DISPATCH_ALIGNMENT, 'green'): (
        'Alignment is {value}%, within the healthy {target}–{max}% band — no action needed.'
    ),
}


def _suggested_action(kpi_slug, status, target, value):
    template = SUGGESTED_ACTIONS.get((kpi_slug, status), '')
    return template.format(
        value=value, min=float(target.min_value), target=float(target.target_value), max=float(target.max_value),
    )


def _get_kpi_target(kpi_slug, year):
    target = KPITarget.objects.filter(kpi_slug=kpi_slug, year=year).first()
    if target:
        return target
    return KPITarget.objects.filter(kpi_slug=kpi_slug, year__lte=year).order_by('-year').first()


def _period_options(today):
    years = list(range(today.year - 2, today.year + 2))
    months = list(enumerate(MONTH_NAMES, start=1))
    quarters = [1, 2, 3, 4]
    return years, months, quarters


def _trend_periods(period_type, year, key_number, count=6):
    """Return the trailing `count` (year, key_number) pairs ending at the
    selected month/quarter, oldest first."""
    periods = []
    y, k = year, key_number
    span = 12 if period_type == 'month' else 4
    for _ in range(count):
        periods.append((y, k))
        k -= 1
        if k < 1:
            k += span
            y -= 1
    return list(reversed(periods))


def build_kpi_scorecard_context(request):
    today = timezone.localdate()
    period_type = (request.GET.get('period_type') or 'month').strip().lower()
    if period_type not in ('month', 'quarter'):
        period_type = 'month'

    year = int(request.GET.get('year') or today.year)
    if period_type == 'month':
        key_number = int(request.GET.get('month') or today.month)
        key_number = min(max(key_number, 1), 12)
    else:
        key_number = int(request.GET.get('quarter') or ((today.month - 1) // 3 + 1))
        key_number = min(max(key_number, 1), 4)

    if period_type == 'month':
        start, end = _month_bounds(year, key_number)
        period_key = f'{year}-{key_number:02d}'
        period_label = f'{MONTH_NAMES[key_number - 1]} {year}'
    else:
        start, end = _quarter_bounds(year, key_number)
        period_key = f'{year}-Q{key_number}'
        period_label = f'Q{key_number} {year}'

    saved_notes = {
        note.kpi_slug: note
        for note in KPIActionNote.objects.filter(period_type=period_type, period_key=period_key)
    }

    kpis = []
    export_rows = []
    for kpi_slug, compute_func in KPI_COMPUTE_FUNCS.items():
        value, detail = compute_func(start, end)
        target = _get_kpi_target(kpi_slug, year)
        if target is None:
            continue
        status = _status_for(target, value)
        suggested = _suggested_action(kpi_slug, status, target, value)
        note = saved_notes.get(kpi_slug)
        kpis.append({
            'slug': kpi_slug,
            'label': KPI_LABELS.get(kpi_slug, kpi_slug),
            'value': value,
            'status': status,
            'detail': detail,
            'min_value': float(target.min_value),
            'target_value': float(target.target_value),
            'max_value': float(target.max_value),
            'weightage_pct': float(target.weightage_pct),
            'description': target.description,
            'suggested_action': suggested,
            'note_text': note.note if note else suggested,
            'note_saved': bool(note),
        })
        export_rows.append({
            'kpi': KPI_LABELS.get(kpi_slug, kpi_slug),
            'period': period_label,
            'value': value,
            'status': status,
            'min': float(target.min_value),
            'target': float(target.target_value),
            'max': float(target.max_value),
        })

    trend_periods = _trend_periods(period_type, year, key_number, count=6)
    trend = []
    for trend_year, trend_key in trend_periods:
        if period_type == 'month':
            t_start, t_end = _month_bounds(trend_year, trend_key)
            t_label = f'{MONTH_NAMES[trend_key - 1][:3]} {trend_year}'
        else:
            t_start, t_end = _quarter_bounds(trend_year, trend_key)
            t_label = f'Q{trend_key} {trend_year}'
        row = {'label': t_label, 'year': trend_year, 'key_number': trend_key}
        for kpi_slug, compute_func in KPI_COMPUTE_FUNCS.items():
            row[kpi_slug], _ = compute_func(t_start, t_end)
        trend.append(row)

    years, months, quarters = _period_options(today)

    # Drill-down export: clicking a KPI's percent or a trend period downloads
    # the raw job-level rows the percentage was computed from, so the numbers
    # can be checked by hand. `kpi=<slug>` downloads one KPI's supporting data
    # for the selected period; `kpi=all` downloads all three (used from the
    # trend table, where a single period-label click covers every KPI).
    kpi_param = (request.GET.get('kpi') or '').strip()
    detail_param = (request.GET.get('detail') or '').strip().lower()
    quarter_for_detail = key_number if period_type == 'quarter' else ((key_number - 1) // 3 + 1)

    if detail_param == 'quarterly' and kpi_param in DRILLDOWN_FUNCS:
        export_rows = build_kpi_quarterly_detail_rows(kpi_param, year, quarter_for_detail)
        export_headers, export_header_labels = QUARTERLY_DETAIL_HEADERS, QUARTERLY_DETAIL_HEADER_LABELS
    elif detail_param == 'focus' and kpi_param in DRILLDOWN_FUNCS:
        export_rows = build_kpi_focus_rows(kpi_param, start, end)
        export_headers, export_header_labels = KPI_FOCUS_HEADERS, KPI_FOCUS_HEADER_LABELS
    elif kpi_param == 'all':
        drilldown_rows = []
        for slug, drilldown_func in DRILLDOWN_FUNCS.items():
            for drow in drilldown_func(start, end):
                drilldown_rows.append({'kpi': KPI_LABELS.get(slug, slug), **drow})
        export_rows = drilldown_rows
        export_headers, export_header_labels = DRILLDOWN_HEADERS, DRILLDOWN_HEADER_LABELS
    elif kpi_param in DRILLDOWN_FUNCS:
        export_rows = [{'kpi': KPI_LABELS.get(kpi_param, kpi_param), **drow} for drow in DRILLDOWN_FUNCS[kpi_param](start, end)]
        export_headers, export_header_labels = DRILLDOWN_HEADERS, DRILLDOWN_HEADER_LABELS
    else:
        export_headers = ['kpi', 'period', 'value', 'status', 'min', 'target', 'max']
        export_header_labels = {
            'kpi': 'KPI', 'period': 'Period', 'value': 'Value (%)', 'status': 'Status',
            'min': 'Min', 'target': 'Target', 'max': 'Max',
        }

    return {
        'period_type': period_type,
        'year': year,
        'month': key_number if period_type == 'month' else today.month,
        'quarter': key_number if period_type == 'quarter' else ((today.month - 1) // 3 + 1),
        'period_key': period_key,
        'period_label': period_label,
        'kpis': kpis,
        'trend': trend,
        'year_options': years,
        'month_options': months,
        'quarter_options': quarters,
        'export_rows': export_rows,
        'headers': export_headers,
        'header_labels': export_header_labels,
    }
