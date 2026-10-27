"""SKU History report: everything the ERP knows about one SKU, job card by job card.

Search for a SKU and get the orders received, what was printed, packed, dispatched and wasted,
what is still owed and the pattern of repeat orders, so the next order can be planned with the
real history in hand. Wastage uses the same definitions as the Wastage Report (printing waste at
each entry's own ups, sorting waste, and the dispatch gap once a job is finalized), so the two
reports agree.
"""
from __future__ import annotations

from datetime import date
from urllib.parse import quote

from django.db.models import Count, Prefetch, Sum
from django.urls import reverse
from django.utils import timezone

from core.models import Dispatch, JobCard, Production
from planning.models import PlanningJob, SkuRecipe

ROW_HEADERS = [
    'jc_no', 'po_number', 'order_date', 'status', 'machine', 'order_qty_pcs', 'planned_pcs', 'printed_pcs',
    'printing_waste_pcs', 'packed_pcs', 'sorting_waste_pcs', 'dispatched_pcs', 'stock_pcs', 'pending_pcs',
    'dispatch_gap_pcs', 'total_wastage_pcs', 'wastage_pct', 'wastage_status', 'last_dispatch', 'lead_time_days', 'note',
]
ROW_LABELS = {
    'jc_no': 'Job Card', 'po_number': 'PO/WO', 'order_date': 'Order Date', 'status': 'Status', 'machine': 'Machine',
    'order_qty_pcs': 'Order Qty (pcs)', 'planned_pcs': 'Planned (pcs)', 'printed_pcs': 'Printed (pcs)',
    'printing_waste_pcs': 'Printing Waste (pcs)', 'packed_pcs': 'Packed (pcs)', 'sorting_waste_pcs': 'Sorting Waste (pcs)',
    'dispatched_pcs': 'Dispatched (pcs)', 'stock_pcs': 'Carried To Stock (pcs)', 'pending_pcs': 'Still Owed (pcs)',
    'dispatch_gap_pcs': 'Dispatch Gap (pcs)', 'total_wastage_pcs': 'Total Wastage (pcs)', 'wastage_pct': 'Wastage %',
    'wastage_status': 'Wastage Status', 'last_dispatch': 'Last Dispatch', 'lead_time_days': 'Order To Last Dispatch (days)',
    'note': 'Note',
}
DISPATCH_HEADERS = ['dispatch_date', 'dc_no', 'jc_no', 'po_number', 'dispatch_qty', 'dispatch_pcs']
DISPATCH_LABELS = {
    'dispatch_date': 'Dispatch Date', 'dc_no': 'DC No', 'jc_no': 'Job Card', 'po_number': 'PO/WO',
    'dispatch_qty': 'Qty (job unit)', 'dispatch_pcs': 'Qty (pcs)',
}
PRINTING_HEADERS = ['date', 'shift', 'jc_no', 'machine', 'impressions', 'output_sheets', 'waste_sheets', 'printed_pcs', 'waste_pcs', 'added_by']
PRINTING_LABELS = {
    'date': 'Date', 'shift': 'Shift', 'jc_no': 'Job Card', 'machine': 'Machine', 'impressions': 'Impressions',
    'output_sheets': 'Good Sheets', 'waste_sheets': 'Waste Sheets', 'printed_pcs': 'Printed (pcs)',
    'waste_pcs': 'Waste (pcs)', 'added_by': 'Added By',
}
PACKING_HEADERS = ['date', 'shift', 'jc_no', 'packed_pcs', 'sorting_waste', 'sorter', 'added_by']
PACKING_LABELS = {
    'date': 'Date', 'shift': 'Shift', 'jc_no': 'Job Card', 'packed_pcs': 'Packed (pcs)',
    'sorting_waste': 'Sorting Waste (pcs)', 'sorter': 'Sorter', 'added_by': 'Added By',
}
MONTH_HEADERS = ['month', 'jobs', 'ordered_pcs', 'dispatched_pcs', 'wastage_pcs']
MONTH_LABELS = {
    'month': 'Order Month', 'jobs': 'Jobs', 'ordered_pcs': 'Ordered (pcs)',
    'dispatched_pcs': 'Dispatched (pcs)', 'wastage_pcs': 'Wastage (pcs)',
}


def _clean(value):
    return (value or '').strip()


def find_matching_skus(query: str, limit: int = 60) -> list[str]:
    """Distinct SKU spellings (case-insensitive) containing the query, from job cards, planning and the SKU master."""
    found: dict[str, str] = {}
    for model, field in ((JobCard, 'SKU'), (PlanningJob, 'sku'), (SkuRecipe, 'sku')):
        values = model.objects.filter(**{f'{field}__icontains': query}).values_list(field, flat=True).distinct()[:500]
        for value in values:
            value = _clean(value)
            if value:
                found.setdefault(value.casefold(), value)
    return sorted(found.values(), key=str.casefold)[:limit]


def suggest_skus(query: str, limit: int = 12) -> list[dict]:
    """Type-ahead: SKUs containing the text, best first (those starting with it, then the most-ordered),
    each with how many job cards it has."""
    query = _clean(query)
    if len(query) < 2:
        return []
    counts: dict[str, int] = {}
    names: dict[str, str] = {}
    for row in JobCard.objects.filter(is_active=True, SKU__icontains=query).values('SKU').annotate(n=Count('id')):
        key = _clean(row['SKU']).casefold()
        if key:
            names.setdefault(key, _clean(row['SKU']))
            counts[key] = counts.get(key, 0) + row['n']
    for value in find_matching_skus(query, limit=200):          # planned-only and master-only SKUs
        key = value.casefold()
        names.setdefault(key, value)
        counts.setdefault(key, 0)
    needle = query.casefold()
    ordered = sorted(names, key=lambda k: (not k.startswith(needle), -counts[k], k))
    return [{'sku': names[k], 'jobs': counts[k]} for k in ordered[:limit]]


def _day(value):
    return value.strftime('%Y-%m-%d') if value else ''


def _pct(part, whole):
    return round(part * 100.0 / whole, 2) if whole else 0.0


def _order_date(card: JobCard | None, planning_job: PlanningJob | None):
    """When the order reached the system: PO date, else PO approval, else the card/plan creation date."""
    if card is not None and card.po_date:
        return card.po_date
    if planning_job is not None and planning_job.po_approval_date:
        return planning_job.po_approval_date
    if card is not None and card.created_at:
        return timezone.localtime(card.created_at).date()
    if planning_job is not None and planning_job.created_at:
        return timezone.localtime(planning_job.created_at).date()
    return None


def _detail_url(card: JobCard | None, planning_job: PlanningJob | None) -> str:
    planning_job_id = (card.planning_job_id if card is not None else None) or (planning_job.id if planning_job else None)
    if planning_job_id:
        return reverse('planning:job_detail', kwargs={'job_id': planning_job_id})
    if card is not None:
        return f"{reverse('job_card_records')}?q={quote(card.job_card_no)}"
    return ''


def _card_row(card: JobCard, base_dispatch_units: dict) -> dict:
    planning_job = card.planning_job
    ups = card.ups or 1
    ppu = max(int(card.pcs_per_unit or 1), 1)
    productions = list(card.productions.all())
    printing = [p for p in productions if p.entry_type == 'printing']
    packing = [p for p in productions if p.entry_type == 'packing']
    dispatches = [d for d in card.dispatch_set.all() if d.is_active]

    order_pcs = int(card.order_qty_pcs)
    planned_pcs = int(card.total_sheets_planned * ups)
    printed_pcs = sum(p.pcs_produced for p in printing)
    printing_waste_pcs = sum(p.waste_sheets * (p.merge_allocated_ups or ups) for p in printing)
    packed_pcs = sum(p.packing_qty for p in packing)
    sorting_waste_pcs = sum(p.sorting_waste_qty for p in packing)
    # A form of a multi-form book is dispatched once, on its base card, in books.
    if card.parent_job_card_id:
        dispatched_pcs = int((base_dispatch_units.get(card.parent_job_card_id) or 0) * ppu)
    else:
        dispatched_pcs = int(sum(d.dispatch_qty for d in dispatches) * ppu)
    stock_pcs = int(planning_job.stock_qty or 0) if planning_job is not None else 0

    base = card.parent_job_card if card.parent_job_card_id else None
    finalized = (
        card.status in ('completed', 'closed') or card.job_status == 'Completed'
        or bool(base and (base.status in ('completed', 'closed') or base.job_status == 'Completed'))
    )
    dispatch_gap_pcs = max(planned_pcs - dispatched_pcs - stock_pcs, 0) if planned_pcs > 0 else 0
    total_wastage_pcs = printing_waste_pcs + sorting_waste_pcs + dispatch_gap_pcs
    pending_pcs = 0 if finalized else max(order_pcs - dispatched_pcs - stock_pcs, 0)

    order_date = _order_date(card, planning_job)
    last_dispatch = max((d.dispatch_date for d in dispatches), default=None)
    lead_time = (last_dispatch - order_date).days if (finalized and last_dispatch and order_date) else ''

    if planned_pcs <= 0:
        note = 'Fulfilled from stock' if stock_pcs > 0 else 'No plan and no stock recorded'
    elif stock_pcs > 0:
        note = f'{stock_pcs:,} pcs carried to stock (not counted as wastage)'
    else:
        note = ''
    if card.form_label:
        note = f'Form: {card.form_label}' + (f'. {note}' if note else '')

    cancelled = bool(planning_job is not None and planning_job.status == 'cancelled')
    if cancelled:
        note = ('Cancelled. ' + note).strip()
    return {
        'kind': 'card',
        'jc_no': card.job_card_no,
        'url': _detail_url(card, planning_job),
        'po_number': card.PO_No or '',
        'order_date': _day(order_date),
        'status': 'Cancelled' if cancelled else card.get_status_display(),
        'machine': card.machine_name.name if card.machine_name_id else '',
        'order_qty_pcs': order_pcs,
        'planned_pcs': planned_pcs,
        'printed_pcs': int(printed_pcs),
        'printing_waste_pcs': int(printing_waste_pcs),
        'packed_pcs': int(packed_pcs),
        'sorting_waste_pcs': int(sorting_waste_pcs),
        'dispatched_pcs': dispatched_pcs,
        'stock_pcs': stock_pcs,
        'pending_pcs': pending_pcs,
        'dispatch_gap_pcs': dispatch_gap_pcs,
        'total_wastage_pcs': total_wastage_pcs,
        'wastage_pct': _pct(total_wastage_pcs, planned_pcs),
        'wastage_status': 'Finalized' if finalized else 'Tentative',
        'last_dispatch': _day(last_dispatch),
        'lead_time_days': lead_time,
        'note': note,
        '_cancelled': cancelled,
        '_finalized': finalized,
        '_order_date': order_date,
        '_dispatches': dispatches,
        '_card': card,
        '_productions': productions,
    }


def _planning_only_row(job: PlanningJob) -> dict:
    ppu = max(int(job.pcs_per_unit or 1), 1)
    order_pcs = int(job.order_qty or 0) * ppu
    cancelled = job.status == 'cancelled'
    order_date = _order_date(None, job)
    return {
        'kind': 'planning',
        'jc_no': job.jc_number,
        'url': _detail_url(None, job),
        'po_number': job.po_number or '',
        'order_date': _day(order_date),
        'status': 'Cancelled' if cancelled else f'{job.get_status_display()} (no job card yet)',
        'machine': job.machine_name or '',
        'order_qty_pcs': order_pcs,
        'planned_pcs': 0, 'printed_pcs': 0, 'printing_waste_pcs': 0, 'packed_pcs': 0, 'sorting_waste_pcs': 0,
        'dispatched_pcs': 0, 'stock_pcs': int(job.stock_qty or 0),
        'pending_pcs': 0 if cancelled else max(order_pcs - int(job.stock_qty or 0), 0),
        'dispatch_gap_pcs': 0, 'total_wastage_pcs': 0, 'wastage_pct': 0.0,
        'wastage_status': '', 'last_dispatch': '', 'lead_time_days': '',
        'note': 'Cancelled' if cancelled else 'Planned, not yet released to production',
        '_cancelled': cancelled, '_finalized': False, '_order_date': order_date, '_dispatches': [], '_card': None, '_productions': [],
    }


def _public(row: dict) -> dict:
    return {key: value for key, value in row.items() if not key.startswith('_')}


def build_sku_history_context(request) -> dict:
    from reports.services import REPORT_CATALOG

    report = next(item for item in REPORT_CATALOG if item['key'] == 'sku-history')
    query = _clean(request.GET.get('sku'))
    context = {
        'report': report,
        'query': query,
        'sku': '',
        'matches': [],
        'headers': ROW_HEADERS,
        'header_labels': ROW_LABELS,
        'filters': {'sku': query},
        'rows': [],
        'export_rows': [],
        'dispatch_rows': [],
        'monthly_rows': [],
        'printing_rows': [],
        'packing_rows': [],
        'summary': None,
        'insights': None,
        'master': None,
    }
    if not query:
        return context

    matches = find_matching_skus(query)
    exact = next((m for m in matches if m.casefold() == query.casefold()), None)
    if exact:
        sku = exact
    elif len(matches) == 1:
        sku = matches[0]
    else:
        context['matches'] = matches
        return context
    context['sku'] = sku
    context['filters'] = {'sku': sku}

    cards = list(
        JobCard.objects.filter(is_active=True, SKU__iexact=sku)
        .select_related('planning_job', 'machine_name', 'parent_job_card')
        .prefetch_related(
            Prefetch('productions', queryset=Production.objects.filter(is_active=True).select_related('machine', 'created_by', 'sorter')),
            Prefetch('dispatch_set', queryset=Dispatch.objects.filter(is_active=True)),
        )
        .order_by('created_at')
    )
    parent_ids = {c.parent_job_card_id for c in cards if c.parent_job_card_id}
    base_dispatch_units = {
        row['job_card_id']: row['total'] or 0
        for row in Dispatch.objects.filter(is_active=True, job_card_id__in=parent_ids)
        .values('job_card_id').annotate(total=Sum('dispatch_qty'))
    }
    rows = [_card_row(card, base_dispatch_units) for card in cards]
    rows += [
        _planning_only_row(job)
        for job in PlanningJob.objects.filter(is_active=True, sku__iexact=sku, job_card__isnull=True)
    ]
    rows.sort(key=lambda r: (r['_order_date'] or date.min, r['jc_no']), reverse=True)   # newest first

    live = [r for r in rows if not r['_cancelled']]
    finalized = [r for r in live if r['_finalized']]
    ordered = sum(r['order_qty_pcs'] for r in live)
    dispatched = sum(r['dispatched_pcs'] for r in live)
    printing_waste = sum(r['printing_waste_pcs'] for r in live)
    sorting_waste = sum(r['sorting_waste_pcs'] for r in live)
    final_plan = sum(r['planned_pcs'] for r in finalized)
    final_waste = sum(r['total_wastage_pcs'] for r in finalized)
    pending = sum(r['pending_pcs'] for r in live)
    stock = sum(r['stock_pcs'] for r in live)
    summary = {
        'jobs': len(live),
        'cancelled': len(rows) - len(live),
        'open_jobs': sum(1 for r in live if not r['_finalized']),
        'ordered_pcs': ordered,
        'printed_pcs': sum(r['printed_pcs'] for r in live),
        'packed_pcs': sum(r['packed_pcs'] for r in live),
        'dispatched_pcs': dispatched,
        'pending_pcs': pending,
        'stock_pcs': stock,
        'printing_waste_pcs': printing_waste,
        'sorting_waste_pcs': sorting_waste,
        'process_waste_pcs': printing_waste + sorting_waste,
        'finalized_jobs': len(finalized),
        'finalized_wastage_pcs': final_waste,
        'finalized_wastage_pct': _pct(final_waste, final_plan),
        'fulfilment_pct': _pct(dispatched, ordered),
    }

    # Repeat-order pattern, for planning the next one.
    dates = sorted({r['_order_date'] for r in live if r['_order_date']})
    gaps = [(b - a).days for a, b in zip(dates, dates[1:]) if (b - a).days > 0]
    today = timezone.localdate()
    qtys = [r['order_qty_pcs'] for r in live if r['order_qty_pcs'] > 0]
    lead_times = [r['lead_time_days'] for r in finalized if r['lead_time_days'] != '']
    insights = {
        'first_order': _day(dates[0]) if dates else '',
        'last_order': _day(dates[-1]) if dates else '',
        'days_since_last_order': (today - dates[-1]).days if dates else '',
        'avg_days_between_orders': round(sum(gaps) / len(gaps)) if gaps else '',
        'avg_order_pcs': round(sum(qtys) / len(qtys)) if qtys else 0,
        'min_order_pcs': min(qtys) if qtys else 0,
        'max_order_pcs': max(qtys) if qtys else 0,
        'avg_lead_time_days': round(sum(lead_times) / len(lead_times)) if lead_times else '',
        'avg_wastage_pct': summary['finalized_wastage_pct'] if finalized else '',
    }

    # By order month.
    months: dict[str, dict] = {}
    for r in live:
        d = r['_order_date']
        key = d.strftime('%Y-%m') if d else 'Undated'
        bucket = months.setdefault(key, {'month': d.strftime('%b %Y') if d else 'Undated', 'jobs': 0, 'ordered_pcs': 0,
                                         'dispatched_pcs': 0, 'wastage_pcs': 0})
        bucket['jobs'] += 1
        bucket['ordered_pcs'] += r['order_qty_pcs']
        bucket['dispatched_pcs'] += r['dispatched_pcs']
        bucket['wastage_pcs'] += r['total_wastage_pcs'] if r['_finalized'] else r['printing_waste_pcs'] + r['sorting_waste_pcs']
    monthly_rows = [months[key] for key in sorted(months, reverse=True)]

    dispatch_rows = []
    for r in rows:
        card = r['_card']
        if card is None:
            continue
        ppu = max(int(card.pcs_per_unit or 1), 1)
        for d in r['_dispatches']:
            dispatch_rows.append({
                'dispatch_date': _day(d.dispatch_date), 'dc_no': d.dc_no, 'jc_no': r['jc_no'], 'url': r['url'],
                'po_number': r['po_number'], 'dispatch_qty': d.dispatch_qty, 'dispatch_pcs': d.dispatch_qty * ppu,
                '_sort': d.dispatch_date,
            })
    dispatch_rows.sort(key=lambda d: d['_sort'], reverse=True)
    for d in dispatch_rows:
        d.pop('_sort')

    # Every printing and packing entry behind these job cards, newest first.
    printing_rows, packing_rows = [], []
    for r in rows:
        card = r['_card']
        if card is None:
            continue
        card_ups = card.ups or 1
        for p in r['_productions']:
            by = 'System (combined layout)' if p.merge_parent_id else (p.created_by.username if p.created_by_id else '')
            if p.entry_type == 'printing':
                printing_rows.append({
                    'date': _day(p.date), 'shift': p.get_shift_display(), 'jc_no': r['jc_no'], 'url': r['url'],
                    'machine': p.machine.name if p.machine_id else '', 'impressions': p.impressions,
                    'output_sheets': p.output_sheets, 'waste_sheets': p.waste_sheets,
                    'printed_pcs': int(p.pcs_produced), 'waste_pcs': int(p.waste_sheets * (p.merge_allocated_ups or card_ups)),
                    'added_by': by, '_sort': (p.date, p.id),
                })
            elif p.entry_type == 'packing':
                packing_rows.append({
                    'date': _day(p.date), 'shift': p.get_shift_display(), 'jc_no': r['jc_no'], 'url': r['url'],
                    'packed_pcs': p.packing_qty, 'sorting_waste': p.sorting_waste_qty,
                    'sorter': p.sorter.name if p.sorter_id else '', 'added_by': by, '_sort': (p.date, p.id),
                })
    for entry_rows in (printing_rows, packing_rows):
        entry_rows.sort(key=lambda e: e['_sort'], reverse=True)
        for e in entry_rows:
            e.pop('_sort')

    recipe = SkuRecipe.objects.filter(sku__iexact=sku).first()
    master = None
    if recipe is not None:
        master = {
            'job_name': recipe.job_name, 'product_type': recipe.product_type, 'material': recipe.material,
            'color_spec': recipe.color_spec, 'application': recipe.application, 'machine': recipe.machine_name,
        }

    public_rows = [_public(r) for r in rows]
    export_notes = [
        f'SKU: {sku}',
        f"{summary['jobs']} job(s): ordered {ordered:,} pcs, dispatched {dispatched:,} pcs, "
        f"still owed {pending:,} pcs, finalized-job wastage {summary['finalized_wastage_pct']}%",
        'Wastage = printing waste + sorting waste + dispatch gap (finalized jobs), as in the Wastage Report.',
    ]
    context.update({
        'summary': summary,
        'insights': insights,
        'master': master,
        'rows': public_rows,
        'export_rows': public_rows,
        'dispatch_rows': dispatch_rows,
        'monthly_rows': monthly_rows,
        'printing_rows': printing_rows,
        'packing_rows': packing_rows,
        'export_notes': export_notes,
        'extra_sheets': [
            {'title': 'Printing Entries', 'headers': PRINTING_HEADERS, 'labels': PRINTING_LABELS, 'rows': printing_rows},
            {'title': 'Packing Entries', 'headers': PACKING_HEADERS, 'labels': PACKING_LABELS, 'rows': packing_rows},
            {'title': 'Dispatches', 'headers': DISPATCH_HEADERS, 'labels': DISPATCH_LABELS, 'rows': dispatch_rows},
            {'title': 'By Month', 'headers': MONTH_HEADERS, 'labels': MONTH_LABELS, 'rows': monthly_rows},
        ],
    })
    return context
