from django.contrib import admin

from .models import (
    BomTemplate, BomTemplateCondition, BomTemplateLine, ItemCategory, ProcessStep,
    QuantityBasis, RawItem, RawItemCostHistory, SkuBom, SkuBomLine, SpecAttribute,
    UnitOfMeasure,
)


@admin.register(UnitOfMeasure)
class UnitOfMeasureAdmin(admin.ModelAdmin):
    list_display = ('code', 'name', 'decimal_places', 'is_active')
    search_fields = ('code', 'name')


@admin.register(ItemCategory)
class ItemCategoryAdmin(admin.ModelAdmin):
    list_display = ('name', 'code', 'default_uom', 'sort_order', 'is_active')
    search_fields = ('name', 'code')


@admin.register(ProcessStep)
class ProcessStepAdmin(admin.ModelAdmin):
    list_display = ('label', 'code', 'production_line', 'sequence', 'is_active')
    list_filter = ('production_line', 'is_active')


class RawItemCostHistoryInline(admin.TabularInline):
    model = RawItemCostHistory
    extra = 0
    readonly_fields = ('unit_cost', 'effective_from', 'source', 'created_by', 'created_at')
    can_delete = False


@admin.register(RawItem)
class RawItemAdmin(admin.ModelAdmin):
    list_display = ('item_code', 'name', 'category', 'uom', 'item_role', 'unit_cost', 'is_active')
    list_filter = ('category', 'production_line', 'is_active')
    search_fields = ('item_code', 'name', 'specification', 'item_role')
    inlines = [RawItemCostHistoryInline]


@admin.register(SpecAttribute)
class SpecAttributeAdmin(admin.ModelAdmin):
    list_display = ('label', 'code', 'source_kind', 'value_type', 'is_match_key', 'is_active')
    list_filter = ('source_kind', 'value_type', 'is_active')
    search_fields = ('code', 'label')


@admin.register(QuantityBasis)
class QuantityBasisAdmin(admin.ModelAdmin):
    list_display = ('label', 'code', 'formula_key', 'is_active')


class BomTemplateConditionInline(admin.TabularInline):
    model = BomTemplateCondition
    extra = 1


class BomTemplateLineInline(admin.TabularInline):
    model = BomTemplateLine
    fk_name = 'template'
    extra = 1


@admin.register(BomTemplate)
class BomTemplateAdmin(admin.ModelAdmin):
    list_display = ('code', 'name', 'production_line', 'priority', 'status', 'is_active')
    list_filter = ('production_line', 'status', 'is_active')
    search_fields = ('code', 'name')
    inlines = [BomTemplateConditionInline, BomTemplateLineInline]


class SkuBomLineInline(admin.TabularInline):
    model = SkuBomLine
    extra = 0


@admin.register(SkuBom)
class SkuBomAdmin(admin.ModelAdmin):
    list_display = ('sku_recipe', 'version', 'is_current', 'status', 'total_material_cost', 'generation_mode')
    list_filter = ('status', 'generation_mode', 'is_current', 'is_active')
    search_fields = ('sku_recipe__sku',)
    inlines = [SkuBomLineInline]
