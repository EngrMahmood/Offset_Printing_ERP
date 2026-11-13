"""Daily Packing & Dispatch Report: the sheet that is e-mailed every day, built from the ERP.

Same layout as the Excel report that was made by hand: a Summary sheet (packing table, dispatch table), a
'Packing Report' data sheet and a 'Dispatch Report' data sheet.

Where each figure comes from (offset jobs only; Flexo is a separate module and is entered by hand for now):
  * Packing: Production rows of type 'packing' for the date. Quantity = packing_qty, Wastage = sorting_waste_qty,
    Wastage % = wastage / quantity. Sheets Qty = quantity / ups.
  * Dispatch: active Dispatch rows for the date. Quantity = dispatch_qty in pieces, Cost = quantity x the job's piece cost
    (planning job unit cost, else the SKU master default).
  * Pcs Waste (theoretical) = the gap between the order quantity and everything dispatched up to and including that line,
    counted only when that line closes the job (the job is closed, or its dispatches have reached the order quantity
    less the allowed shortfall); a job still being delivered in parts shows 0.
  * Physical Waste = what the sorters reported in the sorting (packing) entries of that job up to that date, shown on the
    same closing dispatch line as the theoretical waste.
"""
from __future__ import annotations

import datetime
import io
from collections import defaultdict
from decimal import Decimal

from django.db.models import Sum

from core.models import Dispatch, JobCard, Production
from planning.models import SkuRecipe

# (key, label in the packing table, label in the dispatch table). Order is the order of the e-mailed sheet.
PACKING_ITEMS = [
    ('insert_card', 'Insert Card'), ('label', 'Taffeta Label'), ('fluffing', 'Fluffing'),
    ('visiting_card', 'Visiting/Employee Card'), ('letter_head', 'Letter Head'), ('sticker', 'Sticker'),
    ('stiffeners', 'Stiffeners'), ('a4_rim', 'A4 Rim'), ('report_books', 'Report Books'),
]
DISPATCH_ITEMS = [
    ('insert_card', 'Insert cards'), ('label', 'Label'), ('fluffing', 'Fluffing Instructions'), ('sticker', 'Stickers'),
    ('a4_rim', 'A-4-Paper'), ('stiffeners', 'Stiffeners'), ('report_books', 'Report Books'),
    ('visiting_card', 'Visiting/Employee Card'), ('letter_head', 'Letter Heads'),
]
OTHER_KEY = 'other'
DATA_ITEM_LABEL = {
    'insert_card': 'Insert Card', 'label': 'Label', 'fluffing': 'Fluffing', 'sticker': 'Sticker', 'a4_rim': 'A4 Rim',
    'stiffeners': 'Stiffeners', 'report_books': 'Report Books', 'visiting_card': 'Visiting/Employee Card',
    'letter_head': 'Letter Head', OTHER_KEY: 'Other',
}
# Dispatches that reach the order quantity less this share count as the end of the job.
DELIVERY_TOLERANCE = 0.05


def item_key(product_type: str) -> str:
    text = (product_type or '').strip().lower()
    if 'insert' in text:
        return 'insert_card'
    if 'label' in text and 'letter' not in text:
        return 'label'
    if 'fluff' in text:
        return 'fluffing'
    if 'sticker' in text:
        return 'sticker'
    if 'a4' in text or 'rim' in text:
        return 'a4_rim'
    if 'stiff' in text:
        return 'stiffeners'
    if 'report' in text or 'book' in text:
        return 'report_books'
    if 'visiting' in text or 'employee' in text:
        return 'visiting_card'
    if 'letter' in text:
        return 'letter_head'
    return OTHER_KEY


def _int(value) -> int:
    return int(value or 0)


def _pct(waste: int, qty: int) -> float:
    return round(waste * 100.0 / qty, 2) if qty else 0.0


def _product_types(skus) -> dict[str, str]:
    return {sku: ptype for sku, ptype in SkuRecipe.objects.filter(sku__in=set(skus)).values_list('sku', 'product_type')}


def _piece_cost(card: JobCard, recipe_cost: dict[str, Decimal]) -> float:
    planning = card.planning_job
    cost = getattr(planning, 'unit_cost', None) if planning else None
    if cost is None:
        cost = recipe_cost.get(card.SKU)
    return float(cost or 0)


def build_daily_report(report_date: datetime.date, manual: dict | None = None) -> dict:
    manual = manual or {}

    # ---------- packing ----------
    packing_rows = list(
        Production.objects.filter(is_active=True, entry_type='packing', date=report_date)
        .select_related('job_card', 'job_card__planning_job', 'job_card__material', 'sorter')
        .order_by('shift', 'sorter__name', 'job_card__job_card_no', 'id')
    )
    ptypes = _product_types(p.job_card.SKU for p in packing_rows if p.job_card_id)
    packing_data, packing_sum = [], defaultdict(lambda: [0, 0])
    for p in packing_rows:
        card = p.job_card
        if card is None:
            continue
        key = item_key(ptypes.get(card.SKU, ''))
        qty, waste = _int(p.packing_qty), _int(p.sorting_waste_qty)
        ups = _int(card.ups) or _int(getattr(card.planning_job, 'ups', 0))
        planning = card.planning_job
        packing_data.append({
            'month': report_date.strftime('%B'), 'item': DATA_ITEM_LABEL[key], 'date': report_date, 'shift': p.shift or '',
            'sorter': p.sorter.name if p.sorter_id else '', 'po': card.PO_No or '',
            'site': getattr(planning, 'destination', '') or '', 'jc': card.job_card_no, 'job_name': card.SKU or '',
            'qty': qty, 'waste': waste, 'ups': ups, 'pct': _pct(waste, qty),
            'sheets': round(qty / ups) if ups else 0, 'material': card.material.name if card.material_id else '',
            'department': getattr(planning, 'department', '') or '',
        })
        packing_sum[key][0] += qty
        packing_sum[key][1] += waste

    # ---------- dispatch ----------
    dispatch_rows = list(
        Dispatch.objects.filter(is_active=True, dispatch_date=report_date)
        .select_related('job_card', 'job_card__planning_job')
        .order_by('dc_no', 'job_card__job_card_no', 'id')
    )
    card_ids = {d.job_card_id for d in dispatch_rows}
    skus = {d.job_card.SKU for d in dispatch_rows}
    dtypes = _product_types(skus)
    recipe_cost = dict(SkuRecipe.objects.filter(sku__in=skus).values_list('sku', 'default_unit_cost'))
    shipped_to_date = {
        row['job_card_id']: _int(row['total'])
        for row in Dispatch.objects.filter(is_active=True, job_card_id__in=card_ids, dispatch_date__lte=report_date)
        .values('job_card_id').annotate(total=Sum('dispatch_qty'))
    }
    sorted_to_date = {
        row['job_card_id']: _int(row['total'])
        for row in Production.objects.filter(is_active=True, entry_type='packing', date__lte=report_date, job_card_id__in=card_ids)
        .values('job_card_id').annotate(total=Sum('sorting_waste_qty'))
    }
    dispatch_data = []
    dispatch_sum = defaultdict(lambda: [0, 0, 0.0, 0])
    waste_given = set()
    for d in dispatch_rows:
        card = d.job_card
        planning = card.planning_job
        key = item_key(dtypes.get(card.SKU, ''))
        qty = _int(d.dispatch_qty)
        order_qty = _int(card.order_qty)
        shipped = shipped_to_date.get(card.id, 0)
        closed = (card.workflow_status or '') == 'closed' or (order_qty and shipped >= order_qty * (1 - DELIVERY_TOLERANCE))
        theoretical = max(order_qty - shipped, 0) if closed and card.id not in waste_given else 0
        physical = sorted_to_date.get(card.id, 0) if closed and card.id not in waste_given else 0
        if closed:
            waste_given.add(card.id)
        piece = _piece_cost(card, recipe_cost)
        ups = _int(card.ups) or _int(getattr(planning, 'ups', 0))
        dispatch_data.append({
            '_card': card.id, 'month': report_date.strftime('%b'), 'week': f'Week {report_date.isocalendar()[1]}',
            'date': report_date, 'item': DATA_ITEM_LABEL[key], 'po': card.PO_No or '', 'dc': d.dc_no,
            'site': getattr(planning, 'destination', '') or '', 'jc': card.job_card_no, 'job_name': card.SKU or '',
            'qty': qty, 'piece_cost': piece, 'total_cost': round(qty * piece, 2), 'waste': theoretical,
            'ups': ups, 'order_sheets': round(qty / ups) if ups else 0, 'pct': _pct(theoretical, qty),
            'physical': physical, 'department': getattr(planning, 'department', '') or '',
        })
        bucket = dispatch_sum[key]
        bucket[0] += qty
        bucket[1] += theoretical
        bucket[2] += qty * piece
        bucket[3] += physical

    # ---------- summary tables (every standard item shows, even at 0) ----------
    def table(items, sums, with_cost):
        rows = []
        known = {k for k, _ in items}
        extra = [(OTHER_KEY, 'Other')] if OTHER_KEY in sums and OTHER_KEY not in known else []
        for number, (key, label) in enumerate(items + extra, 1):
            vals = sums.get(key, [0, 0, 0.0, 0])
            row = {'no': number, 'item': label, 'qty': vals[0], 'waste': vals[1], 'pct': _pct(vals[1], vals[0])}
            if with_cost:
                row['cost'] = round(vals[2])
                row['physical'] = vals[3]
            rows.append(row)
        return rows

    packing_table = table(PACKING_ITEMS, packing_sum, False)
    dispatch_table = table(DISPATCH_ITEMS, dispatch_sum, True)

    def total(rows, with_cost):
        qty, waste = sum(r['qty'] for r in rows), sum(r['waste'] for r in rows)
        out = {'qty': qty, 'waste': waste, 'pct': _pct(waste, qty)}
        if with_cost:
            out['cost'] = sum(r['cost'] for r in rows)
            out['physical'] = sum(r['physical'] for r in rows)
        return out

    packing_total, dispatch_total = total(packing_table, False), total(dispatch_table, True)
    manual_pack, manual_disp = _int(manual.get('pack')), _int(manual.get('disp'))
    return {
        'date': report_date,
        'date_label': report_date.strftime('%d %b %Y'),
        'packing_table': packing_table, 'packing_total': packing_total, 'packing_data': packing_data,
        'dispatch_table': dispatch_table, 'dispatch_total': dispatch_total,
        'dispatch_data': [{k: v for k, v in r.items() if k != '_card'} for r in dispatch_data],
        # Offset figures come from the ERP; the manual boxes hold Flexo (or anything else) entered by hand for now.
        'manual': {'pack': manual_pack, 'disp': manual_disp},
        'packing_grand': packing_total['qty'] + manual_pack,
        'dispatch_grand': dispatch_total['qty'] + manual_disp,
    }


def render_xlsx(report: dict) -> bytes:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    wb = Workbook()
    ws = wb.active
    ws.title = 'Daily packing & Dispatc Summary'
    head = PatternFill('solid', fgColor='1F3864')
    total_fill = PatternFill('solid', fgColor='D9E1F2')
    side = Side(style='thin', color='999999')
    box = Border(left=side, right=side, top=side, bottom=side)
    white = Font(bold=True, color='FFFFFF')

    def title(row, text, span):
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
        cell = ws.cell(row, 1, text)
        cell.font, cell.alignment = Font(bold=True, size=14), Alignment(horizontal='center')

    def header(row, labels):
        for col, text in enumerate(labels, 1):
            c = ws.cell(row, col, text)
            c.font, c.fill, c.border = white, head, box
            c.alignment = Alignment(horizontal='center', wrap_text=True)

    def line(row, values, fmt, bold=False, fill=None):
        for col, value in enumerate(values, 1):
            c = ws.cell(row, col, value)
            c.border = box
            if bold:
                c.font = Font(bold=True)
            if fill:
                c.fill = fill
            if col > 2 and fmt.get(col):
                c.number_format = fmt[col]

    title(1, 'UTOPIA PRINTING & PACKAGING', 5)
    ws.cell(2, 1, report['date']).number_format = 'dd-mmm-yyyy'
    ws.cell(2, 3, 'Daily Packing Report').font = Font(bold=True)
    header(3, ['S.No', 'Items', 'Quantity Produced Pcs', 'Wastage Pcs', 'Wastage (%)'])
    row = 4
    for r in report['packing_table']:
        line(row, [r['no'], r['item'], r['qty'], r['waste'], r['pct'] / 100], {3: '#,##0', 4: '#,##0', 5: '0.00%'})
        row += 1
    t = report['packing_total']
    line(row, ['Total', '', t['qty'], t['waste'], t['pct'] / 100], {3: '#,##0', 4: '#,##0', 5: '0.00%'}, True, total_fill)
    line(row + 1, ['ERP Total Production (offset)', '', t['qty']], {3: '#,##0'}, True)
    line(row + 2, ['Flexo / other items (manual)', '', report['manual']['pack']], {3: '#,##0'})
    line(row + 3, ['Grand total', '', report['packing_grand']], {3: '#,##0'}, True, total_fill)

    row += 6
    title(row, 'UTOPIA PRINTING & PACKAGING', 7)
    ws.cell(row + 1, 1, report['date']).number_format = 'dd-mmm-yyyy'
    ws.cell(row + 1, 4, 'Daily Dispatch Summary').font = Font(bold=True)
    header(row + 2, ['S.No', 'Items', 'Quantity Dispatched Pcs', 'Wastage Pcs', 'Wastage (%)', 'Cost (PKR)', 'Physical Waste'])
    row += 3
    fmt = {3: '#,##0', 4: '#,##0', 5: '0.00%', 6: '#,##0', 7: '#,##0'}
    for r in report['dispatch_table']:
        line(row, [r['no'], r['item'], r['qty'], r['waste'], r['pct'] / 100, r['cost'], r['physical']], fmt)
        row += 1
    t = report['dispatch_total']
    line(row, ['Total', '', t['qty'], t['waste'], t['pct'] / 100, t['cost'], t['physical']], fmt, True, total_fill)
    line(row + 1, ['ERP Total Dispatch (offset)', '', t['qty']], {3: '#,##0'}, True)
    line(row + 2, ['Flexo / other items (manual)', '', report['manual']['disp']], {3: '#,##0'})
    line(row + 3, ['Grand total', '', report['dispatch_grand']], {3: '#,##0'}, True, total_fill)
    for col, width in zip('ABCDEFG', (8, 30, 22, 14, 13, 14, 15)):
        ws.column_dimensions[col].width = width

    def data_sheet(name, columns, rows):
        sheet = wb.create_sheet(name)
        for col, (label, _key, _fmt) in enumerate(columns, 1):
            c = sheet.cell(1, col, label)
            c.font, c.fill = white, head
        for r_i, record in enumerate(rows, 2):
            for col, (_label, key, number_format) in enumerate(columns, 1):
                c = sheet.cell(r_i, col, record[key])
                if number_format:
                    c.number_format = number_format
        for col, (label, _key, _fmt) in enumerate(columns, 1):
            sheet.column_dimensions[sheet.cell(1, col).column_letter].width = max(11, min(44, len(label) + 4))
        sheet.freeze_panes = 'A2'

    data_sheet('Packing Report', [
        ('Month', 'month', None), ('MATERIAL', 'item', None), ('Date', 'date', 'dd-mmm-yyyy'), ('Shift', 'shift', None),
        ('Sorter 1', 'sorter', None), ('Po', 'po', None), ('Site', 'site', None), ('JC', 'jc', None),
        ('Job Name', 'job_name', None), ('Quantity', 'qty', '#,##0'), ('Wastage', 'waste', '#,##0'), ('UPS', 'ups', None),
        ('Percentage', 'pct', '0.00'), ('Sheets Qty', 'sheets', '#,##0'), ('Material', 'material', None),
        ('Department', 'department', None),
    ], report['packing_data'])
    data_sheet('Dispatch Report', [
        ('Month', 'month', None), ('Week', 'week', None), ('Date', 'date', 'dd-mmm-yyyy'), ('Item', 'item', None),
        ('Po', 'po', None), ('DR / GIN', 'dc', None), ('Site', 'site', None), ('Job Card', 'jc', None),
        ('Job Name', 'job_name', None), ('Quantity', 'qty', '#,##0'), ('Piece Cost', 'piece_cost', '0.00'),
        ('Total Cost', 'total_cost', '#,##0.00'), ('Pcs Waste', 'waste', '#,##0'), ('Ups', 'ups', None),
        ('Order Sheets Goods', 'order_sheets', '#,##0'), ('Percentage', 'pct', '0.00'),
        ('Physical Waste', 'physical', '#,##0'), ('Department', 'department', None),
    ], report['dispatch_data'])

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()
