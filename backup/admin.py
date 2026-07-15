from django.contrib import admin
from backup.models import BackupSetting, BackupHistory, RestoreHistory

@admin.register(BackupSetting)
class BackupSettingAdmin(admin.ModelAdmin):
    list_display = ('backup_enabled', 'backup_time', 'frequency', 'local_backup_folder', 'updated_at')
    
    def has_add_permission(self, request):
        # We only allow one row of settings
        return not BackupSetting.objects.exists()

@admin.register(BackupHistory)
class BackupHistoryAdmin(admin.ModelAdmin):
    list_display = ('id', 'backup_type', 'start_time', 'finish_time', 'duration_seconds', 'file_name', 'file_size', 'status')
    list_filter = ('backup_type', 'status', 'start_time')
    search_fields = ('file_name', 'backup_location', 'error_message')
    readonly_fields = ('start_time', 'finish_time', 'duration_seconds', 'file_name', 'file_size', 'backup_location', 'sha256_checksum', 'error_message')

@admin.register(RestoreHistory)
class RestoreHistoryAdmin(admin.ModelAdmin):
    list_display = ('id', 'backup', 'timestamp', 'status', 'executed_by')
    list_filter = ('status', 'timestamp')
    readonly_fields = ('backup', 'timestamp', 'status', 'error_message', 'executed_by')
