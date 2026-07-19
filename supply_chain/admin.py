from django.contrib import admin

from .models import (
    ItemProcurementTimeline,
    ItemRequest,
    ItemRequestApproval,
    ItemRequestDepartment,
    ItemRequestQuote,
    ItemRequestType,
    PhysicalStockCount,
    RawMaterialSku,
    StockDemand,
    StockTransaction,
    SupplyChainItem,
)


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


@admin.register(ItemRequestType)
class ItemRequestTypeAdmin(admin.ModelAdmin):
    list_display = ('name', 'code', 'is_active')
    search_fields = ('name', 'code')


@admin.register(ItemRequestDepartment)
class ItemRequestDepartmentAdmin(admin.ModelAdmin):
    list_display = ('name', 'is_active')
    search_fields = ('name',)


@admin.register(ItemRequest)
class ItemRequestAdmin(admin.ModelAdmin):
    list_display = ('request_no', 'item_title', 'request_type', 'department', 'status', 'raised_by', 'request_date', 'is_active')
    list_filter = ('status', 'request_type', 'department', 'is_active')
    search_fields = ('request_no', 'item_title', 'part_number')


@admin.register(ItemRequestApproval)
class ItemRequestApprovalAdmin(admin.ModelAdmin):
    list_display = ('request', 'action', 'stage', 'actor', 'created_at')
    list_filter = ('action', 'stage')


@admin.register(ItemProcurementTimeline)
class ItemProcurementTimelineAdmin(admin.ModelAdmin):
    list_display = ('request', 'item_code', 'po_no', 'received_date', 'unit_price')
    search_fields = ('request__request_no', 'item_code', 'po_no')


@admin.register(ItemRequestQuote)
class ItemRequestQuoteAdmin(admin.ModelAdmin):
    list_display = ('procurement', 'supplier', 'quoted_price', 'uploaded_by', 'uploaded_at')
    search_fields = ('supplier', 'procurement__request__request_no')
