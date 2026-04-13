from django.contrib import admin
from .models import JobCard, Production, Dispatch


class ProductionInline(admin.TabularInline):
    model = Production
    extra = 1


class DispatchInline(admin.TabularInline):
    model = Dispatch
    extra = 1


class JobCardAdmin(admin.ModelAdmin):
    list_display = (
        'job_card_no',
        'SKU',
        'order_qty',
        'total_production',
        'total_dispatch',
        'balance_qty',
        'job_status',
        'waste_percentage'
    )
    inlines = [ProductionInline, DispatchInline]

    list_filter = ('created_at',)
    search_fields = ('job_card_no', 'SKU')

    def get_queryset(self, request):
        qs = super().get_queryset(request)
        return qs


admin.site.register(JobCard, JobCardAdmin)
admin.site.register(Production)
admin.site.register(Dispatch)


from django.contrib import admin

admin.site.site_header = "Offset ERP System"
admin.site.site_title = "Offset ERP"
admin.site.index_title = "Welcome to Production Dashboard"


list_filter = ('created_at',)
search_fields = ('job_card_no', 'SKU')

def get_queryset(self, request):
    qs = super().get_queryset(request)
    return qs