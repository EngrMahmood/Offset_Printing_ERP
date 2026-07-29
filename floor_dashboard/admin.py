from django.contrib import admin

from .models import DailyTarget


@admin.register(DailyTarget)
class DailyTargetAdmin(admin.ModelAdmin):
    list_display = ('date', 'shift', 'target_qty')
    list_filter = ('shift',)
    ordering = ('-date',)
