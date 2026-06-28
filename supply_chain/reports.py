from collections import defaultdict
from datetime import datetime

from .models import StockTransaction


def _parse_optional_date(value):
    if not value:
        return None
    try:
        return datetime.strptime(value, '%Y-%m-%d').date()
    except ValueError:
        return None


def _effective_month(txn):
    if txn.month_str and txn.month_str.strip():
        return txn.month_str.strip()
    if txn.date:
        return txn.date.strftime('%B %Y')
    return 'Unknown'


def _issuance_queryset(month_filter=None, from_date=None, to_date=None):
    qs = (
        StockTransaction.objects
        .filter(transaction_type='ISSUANCE')
        .select_related('item', 'item__material')
        .order_by('date', 'id')
    )
    if month_filter:
        qs = qs.filter(month_str__iexact=month_filter)
    if from_date:
        qs = qs.filter(date__gte=from_date)
    if to_date:
        qs = qs.filter(date__lte=to_date)
    return qs


def _aggregate_consumption(month_filter=None, from_date=None, to_date=None):
    bucket = defaultdict(lambda: {'sheet_qty_pcs': 0, 'pkt_rim_qty': 0, 'consumption_value': 0})
    item_meta = {}

    for txn in _issuance_queryset(month_filter, from_date, to_date):
        month_label = _effective_month(txn)
        key = (txn.item_id, month_label)
        bucket[key]['sheet_qty_pcs'] += txn.sheet_qty_pcs
        bucket[key]['pkt_rim_qty'] += txn.pkt_rim_qty
        bucket[key]['consumption_value'] += float(txn.sheet_qty_pcs) * float(txn.item.unit_cost)
        item_meta[txn.item_id] = txn.item

    rows = []
    for (item_pk, month_label), totals in bucket.items():
        item = item_meta[item_pk]
        rows.append({
            'month': month_label,
            'item': item,
            'item_id': item.item_id or '',
            'item_type': item.material.name,
            'uom': item.uom,
            'sheet_packing_pcs': item.sheet_packing_pcs,
            'sheet_qty_pcs': totals['sheet_qty_pcs'],
            'pkt_rim_qty': totals['pkt_rim_qty'],
            'consumption_value': round(totals['consumption_value'], 2),
        })
    return rows


def build_item_wise_monthly_consumption(month_filter=None, from_date=None, to_date=None):
    rows = _aggregate_consumption(month_filter, from_date, to_date)
    rows.sort(key=lambda row: (row['item_id'] or row['item_type'], row['month']))
    return rows


def build_month_wise_item_consumption(month_filter=None, from_date=None, to_date=None):
    rows = _aggregate_consumption(month_filter, from_date, to_date)
    rows.sort(key=lambda row: (row['month'], row['item_id'] or row['item_type']))
    return rows


def parse_report_filters(request):
    month_filter = (request.GET.get('month') or '').strip() or None
    from_date = _parse_optional_date((request.GET.get('from_date') or '').strip())
    to_date = _parse_optional_date((request.GET.get('to_date') or '').strip())
    return {
        'month': month_filter or '',
        'from_date': (request.GET.get('from_date') or '').strip(),
        'to_date': (request.GET.get('to_date') or '').strip(),
        'month_filter': month_filter,
        'from_date_parsed': from_date,
        'to_date_parsed': to_date,
    }
