from django.contrib import admin

from .models import (
    FlexoJobCard, FlexoJobChangeRequest, FlexoJobEvent, FlexoPlanningDispatchRun, FlexoPoDocument,
    FlexoPlanningJob, FlexoPlanningPrintRun, FlexoPrintingLogEntry, FlexoSkuRecipe,
    FlexoSlittingCuttingLogEntry,
)


@admin.register(FlexoSkuRecipe)
class FlexoSkuRecipeAdmin(admin.ModelAdmin):
    list_display = ('sku', 'master_data_status', 'ups', 'roll_size', 'is_active')
    list_filter = ('master_data_status', 'is_active')
    search_fields = ('sku', 'job_name')


class FlexoPlanningDispatchRunInline(admin.TabularInline):
    model = FlexoPlanningDispatchRun
    extra = 0


class FlexoPlanningPrintRunInline(admin.TabularInline):
    model = FlexoPlanningPrintRun
    extra = 0


@admin.register(FlexoPlanningJob)
class FlexoPlanningJobAdmin(admin.ModelAdmin):
    list_display = ('jc_number', 'sku', 'po_number', 'status', 'order_qty', 'print_meters')
    list_filter = ('status', 'is_active', 'is_cancelled')
    search_fields = ('jc_number', 'sku', 'job_name', 'po_number')
    inlines = [FlexoPlanningDispatchRunInline, FlexoPlanningPrintRunInline]


class FlexoPrintingLogEntryInline(admin.TabularInline):
    model = FlexoPrintingLogEntry
    extra = 0


class FlexoSlittingCuttingLogEntryInline(admin.TabularInline):
    model = FlexoSlittingCuttingLogEntry
    extra = 0


@admin.register(FlexoJobCard)
class FlexoJobCardAdmin(admin.ModelAdmin):
    list_display = ('job_card_no', 'sku', 'status', 'order_qty')
    list_filter = ('status',)
    search_fields = ('job_card_no', 'sku', 'job_name')
    inlines = [FlexoPrintingLogEntryInline, FlexoSlittingCuttingLogEntryInline]


admin.site.register(FlexoJobChangeRequest)
admin.site.register(FlexoJobEvent)
admin.site.register(FlexoPoDocument)
