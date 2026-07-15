from django.db.models import Max, Sum
from django.utils import timezone


def _annual_issuance_value(item):
    total_qty = (
        item.transactions
        .filter(is_active=True, is_approved=True, transaction_type='ISSUANCE')
        .aggregate(total=Sum('sheet_qty_pcs'))['total'] or 0
    )
    return float(total_qty) * float(item.unit_cost)


def assign_abc_classifications(item_values):
    """Pareto ABC: top 70% value = A, next 20% = B, last 10% = C."""
    ranked = sorted(item_values.items(), key=lambda item: item[1], reverse=True)
    total = sum(value for _, value in ranked)
    if total <= 0:
        return {item_pk: 'C' for item_pk, _ in ranked}

    cumulative = 0
    result = {}
    for item_pk, value in ranked:
        previous_share = cumulative / total
        cumulative += value
        if previous_share < 0.70:
            result[item_pk] = 'A'
        elif previous_share < 0.90:
            result[item_pk] = 'B'
        else:
            result[item_pk] = 'C'
    return result


def classify_fsn(days_since_last_issuance):
    if days_since_last_issuance is None:
        return 'Non-Moving'
    if days_since_last_issuance <= 30:
        return 'Fast'
    if days_since_last_issuance <= 90:
        return 'Slow'
    return 'Non-Moving'


def compute_reorder_point(variable_daily_demand, lead_time_days, safety_stock):
    return round((variable_daily_demand * lead_time_days) + safety_stock, 2)


def compute_inventory_turnover(total_issued, opening, closing):
    average_inventory = (opening + closing) / 2
    if average_inventory <= 0:
        return 0
    return round(total_issued / average_inventory, 2)


def _days_since(date_value, today):
    if not date_value:
        return None
    return (today - date_value).days


def _last_transaction_date(item, transaction_type):
    return (
        item.transactions
        .filter(is_active=True, is_approved=True, transaction_type=transaction_type)
        .aggregate(last=Max('date'))['last']
    )


def _last_movement_date(item):
    last_receiving = _last_transaction_date(item, 'RECEIVING')
    last_issuance = _last_transaction_date(item, 'ISSUANCE')
    dates = [value for value in (last_receiving, last_issuance) if value]
    return max(dates) if dates else None


from .physical_count import latest_accuracy_by_item


def enrich_row_with_kpis(row, abc_class, today=None, latest_counts=None):
    """Attach KPI metrics and alert flags to an existing dashboard row dict."""
    today = today or timezone.now().date()
    item = row['item']
    latest_counts = latest_counts if latest_counts is not None else latest_accuracy_by_item()
    latest_count = latest_counts.get(item.pk)

    annual_consumption_value = round(_annual_issuance_value(item), 2)
    last_issuance_date = _last_transaction_date(item, 'ISSUANCE')
    days_since_last_issuance = _days_since(last_issuance_date, today)
    last_movement_date = _last_movement_date(item)
    days_since_last_movement = _days_since(last_movement_date, today)

    reorder_point = compute_reorder_point(
        row['variable_daily_demand'],
        item.lead_time_days,
        item.safety_stock,
    )
    inventory_value = round(float(row['closing']) * float(item.unit_cost), 2)
    inventory_turnover = compute_inventory_turnover(
        row['issuance'],
        row['opening'],
        row['closing'],
    )

    safety_stock_alert = row['closing'] <= item.safety_stock
    reorder_alert = row['closing'] <= reorder_point
    dead_stock = days_since_last_movement is not None and days_since_last_movement > 180
    slow_moving = (
        days_since_last_issuance is not None
        and 30 < days_since_last_issuance <= 90
    )

    row.update({
        'abc_class': abc_class,
        'fsn_class': classify_fsn(days_since_last_issuance),
        'annual_consumption_value': annual_consumption_value,
        'reorder_point': reorder_point,
        'inventory_value': inventory_value,
        'inventory_turnover': inventory_turnover,
        'days_since_last_issuance': days_since_last_issuance,
        'days_since_last_movement': days_since_last_movement,
        'last_issuance_date': last_issuance_date,
        'safety_stock_alert': safety_stock_alert,
        'reorder': reorder_alert,
        'dead_stock': dead_stock,
        'slow_moving': slow_moving,
        'inventory_accuracy': float(latest_count.accuracy_percent) if latest_count else None,
        'physical_count_date': latest_count.count_date if latest_count else None,
    })
    return row


def build_kpi_dashboard_data(items=None):
    from .services import build_dashboard_data

    rows = build_dashboard_data(items)
    annual_values = {row['item'].pk: _annual_issuance_value(row['item']) for row in rows}
    abc_map = assign_abc_classifications(annual_values)
    latest_counts = latest_accuracy_by_item()

    enriched = []
    for row in rows:
        enriched.append(enrich_row_with_kpis(row, abc_map.get(row['item'].pk, 'C'), latest_counts=latest_counts))

    accuracy_values = [row['inventory_accuracy'] for row in enriched if row['inventory_accuracy'] is not None]
    summary = {
        'abc_a': sum(1 for row in enriched if row['abc_class'] == 'A'),
        'abc_b': sum(1 for row in enriched if row['abc_class'] == 'B'),
        'abc_c': sum(1 for row in enriched if row['abc_class'] == 'C'),
        'fsn_fast': sum(1 for row in enriched if row['fsn_class'] == 'Fast'),
        'fsn_slow': sum(1 for row in enriched if row['fsn_class'] == 'Slow'),
        'fsn_non_moving': sum(1 for row in enriched if row['fsn_class'] == 'Non-Moving'),
        'stockout': sum(1 for row in enriched if row['stockout']),
        'reorder': sum(1 for row in enriched if row['reorder']),
        'safety_stock': sum(1 for row in enriched if row['safety_stock_alert']),
        'overstock': sum(1 for row in enriched if row['overstock']),
        'dead_stock': sum(1 for row in enriched if row['dead_stock']),
        'slow_moving': sum(1 for row in enriched if row['slow_moving']),
        'total_inventory_value': round(sum(row['inventory_value'] for row in enriched), 2),
        'avg_inventory_accuracy': round(sum(accuracy_values) / len(accuracy_values), 2) if accuracy_values else None,
        'items_with_physical_count': len(accuracy_values),
    }
    return enriched, summary
