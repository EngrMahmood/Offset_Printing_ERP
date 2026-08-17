from django.contrib import admin
from sheets_sync.models import SheetsSyncSetting, SheetsRowIndex, SheetsSyncLog


@admin.register(SheetsSyncSetting)
class SheetsSyncSettingAdmin(admin.ModelAdmin):
    list_display = ('enabled', 'spreadsheet_id', 'flush_interval_seconds', 'updated_at')

    def has_add_permission(self, request):
        return not SheetsSyncSetting.objects.exists()


@admin.register(SheetsSyncLog)
class SheetsSyncLogAdmin(admin.ModelAdmin):
    list_display = ('id', 'timestamp', 'tab_name', 'batch_size', 'status', 'duration_ms')
    list_filter = ('status', 'tab_name', 'timestamp')
    search_fields = ('tab_name', 'error_message')
    readonly_fields = ('timestamp', 'tab_name', 'batch_size', 'status', 'error_message', 'duration_ms')

    def has_add_permission(self, request):
        return False


@admin.register(SheetsRowIndex)
class SheetsRowIndexAdmin(admin.ModelAdmin):
    list_display = ('tab_name', 'object_pk', 'row_number', 'updated_at')
    list_filter = ('tab_name',)
    search_fields = ('object_pk',)

    def has_add_permission(self, request):
        return False
