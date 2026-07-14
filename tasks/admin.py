from django.contrib import admin
from .models import Team, Task, TaskComment

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
