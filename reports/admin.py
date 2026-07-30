from django.contrib import admin

from .models import KPIActionNote, KPITarget, MachinePlanningJcSelection, ScheduledReport


@admin.register(KPITarget)
class KPITargetAdmin(admin.ModelAdmin):
    list_display = ('kpi_slug', 'year', 'weightage_pct', 'min_value', 'target_value', 'max_value', 'higher_is_better')
    list_filter = ('year', 'kpi_slug')


@admin.register(KPIActionNote)
class KPIActionNoteAdmin(admin.ModelAdmin):
    list_display = ('kpi_slug', 'period_type', 'period_key', 'status', 'updated_by', 'updated_at')
    list_filter = ('period_type', 'kpi_slug')
    readonly_fields = ('updated_at',)


@admin.register(ScheduledReport)
class ScheduledReportAdmin(admin.ModelAdmin):
    list_display = ('name', 'report_slug', 'frequency', 'is_active', 'next_run_at')
    list_filter = ('frequency', 'is_active')


@admin.register(MachinePlanningJcSelection)
class MachinePlanningJcSelectionAdmin(admin.ModelAdmin):
    list_display = ('jc_number', 'is_excluded', 'updated_by', 'updated_at')
    list_filter = ('is_excluded',)
