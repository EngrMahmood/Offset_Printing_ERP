from django.contrib import admin

from .models import PlateRequest


@admin.register(PlateRequest)
class PlateRequestAdmin(admin.ModelAdmin):
    list_display = (
        'id',
        'planning_job',
        'job_card',
        'status',
        'machine',
        'department',
        'requested_by',
        'requested_at',
    )
    list_filter = ('status', 'machine', 'department')
    search_fields = (
        'planning_job__jc_number',
        'job_card__job_card_no',
        'sku_recipe__sku',
        'set_no',
        'new_set_no',
        'vendor',
    )
    readonly_fields = (
        'requested_by',
        'requested_at',
        'sent_by',
        'sent_at',
        'received_by',
        'received_at',
        'designer',
    )
