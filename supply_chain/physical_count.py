from decimal import Decimal

from .models import PhysicalStockCount, SupplyChainItem
from .services import build_dashboard_data


def compute_inventory_accuracy(physical_stock, system_stock):
    """(Physical Stock ÷ System Stock) × 100 per inventory KPI spec."""
    physical_stock = int(physical_stock or 0)
    system_stock = int(system_stock or 0)
    if system_stock > 0:
        return round(Decimal(physical_stock) / Decimal(system_stock) * Decimal('100'), 2)
    if physical_stock == 0:
        return Decimal('100.00')
    return Decimal('0.00')


def get_item_system_stock(item):
    row = build_dashboard_data([item])[0]
    return {
        'system_sheet_qty': row['closing'],
        'system_pkt_rim_qty': 0,
    }


def build_physical_count_rows(items=None):
    if items is None:
        items = SupplyChainItem.objects.select_related('material').order_by('item_id', 'material__name')

    rows = []
    for item in items:
        stock = get_item_system_stock(item)
        latest = item.physical_counts.order_by('-count_date', '-id').first()
        rows.append({
            'item': item,
            'system_sheet_qty': stock['system_sheet_qty'],
            'latest_count': latest,
            'latest_accuracy': latest.accuracy_percent if latest else None,
        })
    return rows


def save_physical_count(item, count_date, physical_sheet_qty, physical_pkt_rim_qty=0, notes=''):
    stock = get_item_system_stock(item)
    accuracy = compute_inventory_accuracy(physical_sheet_qty, stock['system_sheet_qty'])
    return PhysicalStockCount.objects.create(
        item=item,
        count_date=count_date,
        physical_sheet_qty=int(physical_sheet_qty or 0),
        physical_pkt_rim_qty=int(physical_pkt_rim_qty or 0),
        system_sheet_qty=stock['system_sheet_qty'],
        system_pkt_rim_qty=stock['system_pkt_rim_qty'],
        accuracy_percent=accuracy,
        notes=(notes or '').strip(),
    )


def physical_count_history(limit=200):
    return (
        PhysicalStockCount.objects
        .select_related('item', 'item__material')
        .order_by('-count_date', '-id')[:limit]
    )


def latest_accuracy_by_item():
    result = {}
    for count in PhysicalStockCount.objects.order_by('item_id', '-count_date', '-id'):
        if count.item_id not in result:
            result[count.item_id] = count
    return result
