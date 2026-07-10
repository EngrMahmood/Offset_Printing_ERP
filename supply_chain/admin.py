from django.contrib import admin

from .models import PhysicalStockCount, RawMaterialSku, StockDemand, StockTransaction, SupplyChainItem


@admin.register(RawMaterialSku)
class RawMaterialSkuAdmin(admin.ModelAdmin):
    list_display = ('sku', 'material', 'purchase_sheet_size', 'uom', 'unit_cost', 'safety_stock', 'is_active')
    search_fields = ('sku', 'material__name', 'purchase_sheet_size')
    list_filter = ('is_active',)


@admin.register(SupplyChainItem)
class SupplyChainItemAdmin(admin.ModelAdmin):
    list_display = ('item_id', 'material', 'uom')
    search_fields = ('item_id', 'material__name')


@admin.register(StockTransaction)
class StockTransactionAdmin(admin.ModelAdmin):
    list_display = ('transaction_type', 'source', 'raw_material_sku', 'date', 'month_str', 'gin_jc', 'sheet_qty_pcs', 'pkt_rim_qty')
    list_filter = ('transaction_type', 'source', 'date')
    search_fields = ('raw_material_sku__sku', 'gin_jc', 'job_card__job_card_no')


@admin.register(StockDemand)
class StockDemandAdmin(admin.ModelAdmin):
    list_display = ('raw_material_sku', 'month_str', 'sheet_qty_pcs', 'pkt_rim_qty')
    search_fields = ('raw_material_sku__sku', 'month_str')


@admin.register(PhysicalStockCount)
class PhysicalStockCountAdmin(admin.ModelAdmin):
    list_display = (
        'count_date', 'raw_material_sku', 'physical_sheet_qty', 'system_sheet_qty',
        'accuracy_percent', 'notes',
    )
    list_filter = ('count_date',)
    search_fields = ('raw_material_sku__sku', 'notes')
