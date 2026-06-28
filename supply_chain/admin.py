from django.contrib import admin

from .models import PhysicalStockCount, StockDemand, StockTransaction, SupplyChainItem


@admin.register(SupplyChainItem)
class SupplyChainItemAdmin(admin.ModelAdmin):
    list_display = ('item_id', 'material', 'uom', 'sheet_packing_pcs', 'unit_cost', 'safety_stock')
    search_fields = ('item_id', 'material__name')


@admin.register(StockTransaction)
class StockTransactionAdmin(admin.ModelAdmin):
    list_display = ('transaction_type', 'source', 'item', 'date', 'month_str', 'gin_jc', 'sheet_qty_pcs', 'pkt_rim_qty')
    list_filter = ('transaction_type', 'source', 'date')
    search_fields = ('item__item_id', 'gin_jc', 'job_card__job_card_no')


@admin.register(StockDemand)
class StockDemandAdmin(admin.ModelAdmin):
    list_display = ('item', 'month_str', 'sheet_qty_pcs', 'pkt_rim_qty')
    search_fields = ('item__item_id', 'month_str')


@admin.register(PhysicalStockCount)
class PhysicalStockCountAdmin(admin.ModelAdmin):
    list_display = (
        'count_date', 'item', 'physical_sheet_qty', 'system_sheet_qty',
        'accuracy_percent', 'notes',
    )
    list_filter = ('count_date',)
    search_fields = ('item__item_id', 'notes')
