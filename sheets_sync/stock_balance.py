"""Computed stock-balance summary, one row per RawMaterialSku.

Unlike the rest of sheets_sync (one registry entry per model, upserted
row-per-instance via SheetsRowIndex), this tab is a full recompute-and-
rewrite on every trigger: the ERP itself has no stored running balance
anywhere (confirmed by reading supply_chain/views.py — it's a pure
transaction log with no aggregation), so there is no single instance whose
post_save could drive an upsert. Instead the whole tab is derived fresh
from StockTransaction each time and the whole sheet range is overwritten.

Balance formula (confirmed with the user, since the ERP doesn't compute
this anywhere itself so there was no existing formula to copy):
    balance = opening + receiving - issuance + adjustment
Only active, approved transactions count — same filter kpis.py uses for
its issuance-cost calculation.
"""
from django.db.models import Q, Sum

STOCK_BALANCE_HEADERS = [
    'SKU', 'Material', 'UOM', 'Opening', 'Receiving', 'Issuance',
    'Adjustment', 'Balance', 'Safety Stock', 'Below Safety Stock', 'Computed At',
]


def compute_stock_balance():
    from django.utils import timezone
    from supply_chain.models import RawMaterialSku, StockTransaction

    now = timezone.now().isoformat()
    rows = []
    for sku in RawMaterialSku.objects.select_related('material').order_by('sku'):
        agg = StockTransaction.objects.filter(
            raw_material_sku=sku, is_active=True, is_approved=True,
        ).aggregate(
            opening=Sum('sheet_qty_pcs', filter=Q(transaction_type='OPENING')),
            receiving=Sum('sheet_qty_pcs', filter=Q(transaction_type='RECEIVING')),
            issuance=Sum('sheet_qty_pcs', filter=Q(transaction_type='ISSUANCE')),
            adjustment=Sum('sheet_qty_pcs', filter=Q(transaction_type='ADJUSTMENT')),
        )
        opening = agg['opening'] or 0
        receiving = agg['receiving'] or 0
        issuance = agg['issuance'] or 0
        adjustment = agg['adjustment'] or 0
        balance = opening + receiving - issuance + adjustment

        rows.append([
            sku.sku,
            str(sku.material) if sku.material_id else '',
            sku.uom,
            opening,
            receiving,
            issuance,
            adjustment,
            balance,
            sku.safety_stock,
            'YES' if balance < sku.safety_stock else '',
            now,
        ])
    return rows


def rewrite_stock_balance_tab(spreadsheet=None):
    """Recomputes and fully overwrites the Stock Balance tab. Returns row count."""
    from sheets_sync import client as sheets_client

    if spreadsheet is None:
        spreadsheet = sheets_client.open_spreadsheet()

    worksheet = sheets_client.get_or_create_worksheet(spreadsheet, 'Stock Balance', STOCK_BALANCE_HEADERS)
    rows = compute_stock_balance()

    worksheet.clear()
    worksheet.update('A1', [STOCK_BALANCE_HEADERS] + rows, value_input_option='RAW')
    return len(rows)
