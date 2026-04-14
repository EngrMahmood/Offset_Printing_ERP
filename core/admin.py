from django.contrib import admin
from django.urls import path, reverse
from django.template.response import TemplateResponse
from django.utils.html import format_html


from .models import JobCard, Production, Dispatch, Machine, Department, Material, Operator


# =========================
# INLINE MODELS
# =========================

class ProductionInline(admin.TabularInline):
    model = Production
    extra = 1


class DispatchInline(admin.TabularInline):
    model = Dispatch
    extra = 1


# =========================
# JOB CARD ADMIN (ERP CORE)
# =========================

@admin.register(JobCard)
class JobCardAdmin(admin.ModelAdmin):

    change_list_template = "admin/core/jobcard_change_list.html"

    # -------------------------
    # LIST VIEW (ONLY KPIs)
    # -------------------------
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

    list_filter = (
        'created_at',
        'status',
        'machine_name',
        'department',
    )

    search_fields = (
        'job_card_no',
        'SKU',
        'PO_No',
    )

    ordering = ('-created_at',)

    inlines = [ProductionInline, DispatchInline]

    # -------------------------
    # ADD / EDIT FORM (ALL FIELDS)
    # -------------------------
    fieldsets = (
        ("Basic Information", {
            "fields": (
                "job_card_no",
                "SKU",
                "PO_No",
                "po_date",
                "month"
            )
        }),

        ("Material Details", {
            "fields": (
                "material",
                "colour",
                "application"
            )
        }),

        ("Production Details", {
            "fields": (
                "order_qty",
                "ups",
                "wastage",
                "actual_sheet_required"
            )
        }),

        ("Printing Details", {
            "fields": (
                "print_sheet_size",
                "purchase_sheet_size",
                "purchase_sheet_ups"
            )
        }),

        ("Machine & Department", {
            "fields": (
                "machine_name",
                "department",
                "die_cutting"
            )
        }),

        ("Extra Information", {
            "fields": (
                "destination",
                "remarks",
                "status",
                "is_active"
            )
        }),
    )

    # -------------------------
    # BULK UPLOAD BUTTON
    # -------------------------
    def bulk_upload_button(self, obj):
        url = reverse('admin:jobcard_bulk_upload')

        return format_html(
            '<a class="button" style="background:#417690;color:white;padding:5px 10px;border-radius:5px;text-decoration:none;" href="{}">📥 Bulk Upload</a>',
            url
        )

    bulk_upload_button.short_description = "Bulk Upload"

    # -------------------------
    # CUSTOM ADMIN URL
    # -------------------------
    def get_urls(self):
        urls = super().get_urls()
        custom_urls = [
            path(
                'bulk-upload/',
                self.admin_site.admin_view(self.bulk_upload_view),
                name='jobcard_bulk_upload'
            ),
        ]
        return custom_urls + urls

    def bulk_upload_view(self, request):
        from .bulk_upload import process_jobcard_upload

        context = {}

        if request.method == "POST":
            file = request.FILES.get("file")
            if file:
                result = process_jobcard_upload(file)
                context = result

        return TemplateResponse(request, "admin/bulk_upload.html", context)


# =========================
# PRODUCTION ADMIN
# =========================

@admin.register(Production)
class ProductionAdmin(admin.ModelAdmin):

    list_display = (
        'job_card',
        'date',
        'shift',
        'machine',
        'output_qty',
        'waste_qty',
        'operator',
        'oee_display'
    )

    list_filter = (
        'date',
        'shift',
        'machine',
        'operator',
    )

    search_fields = (
        'job_card__job_card_no',
        'machine__name',
        'operator__name',
    )

    autocomplete_fields = ['job_card', 'machine', 'operator']

    date_hierarchy = 'date'

    def oee_display(self, obj):
        return round(obj.oee() * 100, 2)

    oee_display.short_description = "OEE %"
    
@admin.register(Operator)
class OperatorAdmin(admin.ModelAdmin):
    list_display = ['name', 'employee_code', 'is_active']
    search_fields = ['name', 'employee_code']
    list_filter = ['is_active']


@admin.register(Machine)
class MachineAdmin(admin.ModelAdmin):
    list_display = ['name', 'standard_speed', 'is_active']
    search_fields = ['name']

@admin.register(Department)
class DepartmentAdmin(admin.ModelAdmin):
    search_fields = ['name']

@admin.register(Material)
class MaterialAdmin(admin.ModelAdmin):
    search_fields = ['name']

# =========================
# DISPATCH ADMIN
# =========================

@admin.register(Dispatch)
class DispatchAdmin(admin.ModelAdmin):

    list_display = (
        'job_card',
        'order_qty',
        'dc_no',
        'dispatch_date',
        'dispatch_qty',
        'balance_check'
    )

    list_filter = ('dispatch_date',)

    search_fields = ('job_card__job_card_no','dc_no',)

    def balance_check(self, obj):
        return obj.job_card.balance_qty

    balance_check.short_description = "Remaining Balance"

    def order_qty(self, obj):
        return obj.job_card.order_qty

    order_qty.short_description = "Order Qty"


# =========================
# MASTER DATA ADMIN
# =========================


#admin.site.register(Department)
#admin.site.register(Material)


# =========================
# ERP BRANDING
# =========================

admin.site.site_header = "Offset ERP System"
admin.site.site_title = "Offset ERP"
admin.site.index_title = "Production Dashboard"