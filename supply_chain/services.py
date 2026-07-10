from django.db.models import Sum

from .models import RawMaterialSku, StockDemand, StockTransaction


def build_dashboard_data(items=None):
    """Aggregate stock metrics per raw material SKU."""
    if items is None:
        items = RawMaterialSku.objects.select_related('material').filter(is_active=True)

    dashboard_data = []
    for item in items:
        opening = item.transactions.filter(transaction_type='OPENING').aggregate(t=Sum('sheet_qty_pcs'))['t'] or 0
        receiving = item.transactions.filter(transaction_type='RECEIVING').aggregate(t=Sum('sheet_qty_pcs'))['t'] or 0
        issuance = item.transactions.filter(transaction_type='ISSUANCE').aggregate(t=Sum('sheet_qty_pcs'))['t'] or 0
        adjustment = item.transactions.filter(transaction_type='ADJUSTMENT').aggregate(t=Sum('sheet_qty_pcs'))['t'] or 0

        monthly_demand = item.demands.aggregate(t=Sum('sheet_qty_pcs'))['t'] or 0

        closing = (opening + receiving) - issuance + adjustment

        stockout = closing <= 0
        overstock = closing > item.max_stock_level

        variable_daily_demand = round(monthly_demand / 30, 2) if monthly_demand else 0
        stock_forecast_days = (
            round(closing / variable_daily_demand, 1)
            if variable_daily_demand > 0
            else None
        )

        dashboard_data.append({
            'item': item,
            'opening': opening,
            'receiving': receiving,
            'issuance': issuance,
            'adjustment': adjustment,
            'closing': closing,
            'monthly_demand': monthly_demand,
            'variable_daily_demand': variable_daily_demand,
            'stock_forecast_days': stock_forecast_days,
            'stockout': stockout,
            'overstock': overstock,
        })

    return dashboard_data


def transaction_queryset(transaction_type, month_filter=None):
    qs = (
        StockTransaction.objects
        .filter(transaction_type=transaction_type)
        .select_related('raw_material_sku', 'raw_material_sku__material', 'job_card', 'production')
        .order_by('-date', '-id')
    )
    if month_filter:
        qs = qs.filter(month_str__iexact=month_filter)
    return qs


def demand_queryset(month_filter=None):
    qs = (
        StockDemand.objects
        .select_related('raw_material_sku', 'raw_material_sku__material')
        .order_by('-month_str', '-id')
    )
    if month_filter:
        qs = qs.filter(month_str__iexact=month_filter)
    return qs
