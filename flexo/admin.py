from django.contrib import admin

from .models import (
    FlexoJobCard, FlexoPlanningDispatchRun, FlexoPlanningJob,
    FlexoPrintingLogEntry, FlexoSlittingCuttingLogEntry,
)


class FlexoPlanningDispatchRunInline(admin.TabularInline):
    model = FlexoPlanningDispatchRun
    extra = 0


@admin.register(FlexoPlanningJob)
class FlexoPlanningJobAdmin(admin.ModelAdmin):
    list_display = ('jc_number', 'material_family', 'sku', 'po_number', 'status', 'order_qty', 'created_at')
    list_filter = ('material_family', 'status')
    search_fields = ('jc_number', 'sku', 'job_name', 'po_number')
    inlines = [FlexoPlanningDispatchRunInline]


class FlexoPrintingLogEntryInline(admin.TabularInline):
    model = FlexoPrintingLogEntry
    extra = 0


class FlexoSlittingCuttingLogEntryInline(admin.TabularInline):
    model = FlexoSlittingCuttingLogEntry
    extra = 0


@admin.register(FlexoJobCard)
class FlexoJobCardAdmin(admin.ModelAdmin):
    list_display = ('job_card_no', 'material_family', 'sku', 'status', 'order_qty', 'created_at')
    list_filter = ('material_family', 'status')
    search_fields = ('job_card_no', 'sku', 'job_name')
    inlines = [FlexoPrintingLogEntryInline, FlexoSlittingCuttingLogEntryInline]
