from django.contrib import admin

from bot.models import BotAutomation, BotExecution


@admin.register(BotAutomation)
class BotAutomationAdmin(admin.ModelAdmin):
    list_display = ('name', 'code', 'report_slug', 'frequency', 'send_time', 'is_active', 'last_run_at', 'last_status')
    list_filter = ('is_active', 'frequency', 'report_slug')
    search_fields = ('name', 'code', 'report_slug')
    readonly_fields = ('last_run_at', 'next_run_at', 'last_status', 'created_at', 'updated_at')


@admin.register(BotExecution)
class BotExecutionAdmin(admin.ModelAdmin):
    list_display = ('bot', 'started_at', 'trigger', 'status', 'record_count', 'recipients_to')
    list_filter = ('status', 'trigger', 'bot')
    search_fields = ('bot__name', 'rendered_subject', 'recipients_to')
    readonly_fields = tuple(
        field.name for field in BotExecution._meta.fields if field.name != 'id'
    )

    def has_add_permission(self, request):
        return False
