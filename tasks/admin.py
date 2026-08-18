from django.contrib import admin
from .models import Team, Task, TaskComment, TaskNotificationSettings, TaskNotificationLog

@admin.register(Team)
class TeamAdmin(admin.ModelAdmin):
    list_display = ('name', 'created_at')
    filter_horizontal = ('members',)
    search_fields = ('name',)

@admin.register(Task)
class TaskAdmin(admin.ModelAdmin):
    list_display = ('title', 'assignee', 'assigned_team', 'priority', 'status', 'due_date', 'score', 'completed_at')
    list_filter = ('priority', 'status', 'due_date')
    search_fields = ('title', 'description')
    readonly_fields = ('created_at', 'updated_at')

@admin.register(TaskComment)
class TaskCommentAdmin(admin.ModelAdmin):
    list_display = ('task', 'user', 'created_at')
    search_fields = ('comment',)

@admin.register(TaskNotificationSettings)
class TaskNotificationSettingsAdmin(admin.ModelAdmin):
    list_display = ('assignment_email_enabled', 'reminders_enabled', 'reminder_interval_days', 'remind_from', 'updated_at')

    def has_add_permission(self, request):
        return not TaskNotificationSettings.objects.exists()

@admin.register(TaskNotificationLog)
class TaskNotificationLogAdmin(admin.ModelAdmin):
    list_display = ('task', 'kind', 'sent_at', 'status')
    list_filter = ('kind', 'status')
    search_fields = ('task__title', 'recipients_to', 'recipients_cc', 'recipients_bcc')
    readonly_fields = [f.name for f in TaskNotificationLog._meta.fields]

    def has_add_permission(self, request):
        return False
