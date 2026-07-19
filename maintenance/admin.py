from django.contrib import admin

from .models import (
    FaultCategory, MachineDowntime, MaintenanceActivityLog, MaintenanceApproval, MaintenanceAttachment,
    MaintenanceRecord, MaintenanceServiceJob, MaintenanceSparePart, PreventiveMaintenancePlan,
)


@admin.register(FaultCategory)
class FaultCategoryAdmin(admin.ModelAdmin):
    list_display = ('name', 'is_active')
    search_fields = ('name',)


class MaintenanceSparePartInline(admin.TabularInline):
    model = MaintenanceSparePart
    extra = 0


class MaintenanceServiceJobInline(admin.TabularInline):
    model = MaintenanceServiceJob
    extra = 0


@admin.register(MaintenanceRecord)
class MaintenanceRecordAdmin(admin.ModelAdmin):
    list_display = ('record_no', 'machine', 'reported_date', 'maintenance_type', 'priority', 'status')
    list_filter = ('status', 'maintenance_type', 'priority', 'execution_type', 'machine')
    search_fields = ('record_no', 'fault_description')
    inlines = [MaintenanceSparePartInline, MaintenanceServiceJobInline]


@admin.register(MachineDowntime)
class MachineDowntimeAdmin(admin.ModelAdmin):
    list_display = ('machine', 'record', 'start_at', 'end_at', 'reason', 'scheduled_minutes_lost')
    list_filter = ('reason', 'machine')


@admin.register(MaintenanceActivityLog)
class MaintenanceActivityLogAdmin(admin.ModelAdmin):
    list_display = ('record', 'actor', 'action', 'from_status', 'to_status', 'created_at')
    list_filter = ('action',)


@admin.register(MaintenanceApproval)
class MaintenanceApprovalAdmin(admin.ModelAdmin):
    list_display = ('record', 'actor', 'action', 'created_at')
    list_filter = ('action',)


@admin.register(PreventiveMaintenancePlan)
class PreventiveMaintenancePlanAdmin(admin.ModelAdmin):
    list_display = ('title', 'machine', 'interval_type', 'interval_value', 'next_due_at', 'is_active')
    list_filter = ('interval_type', 'is_active', 'machine')


@admin.register(MaintenanceAttachment)
class MaintenanceAttachmentAdmin(admin.ModelAdmin):
    list_display = ('record', 'caption', 'uploaded_by', 'uploaded_at')
