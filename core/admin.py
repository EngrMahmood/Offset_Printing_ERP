from django.contrib import admin
from django.utils.html import format_html
from django.urls import path
from django.template.response import TemplateResponse
from django.utils.safestring import mark_safe

from .models import JobCard, Production, Dispatch
from .models import Machine, Department, Material


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
# JOB CARD ADMIN (MAIN ERP SCREEN)
# =========================

@admin.register(JobCard)
class JobCardAdmin(admin.ModelAdmin):

    list_display = (
        'job_card_no',
        'SKU',
        'order_qty',
        'total_production',
        'total_dispatch',
        'balance_qty',
        'job_status',
        'waste_percentage',
        'bulk_upload_button'
    )


# =========================
    # BUTTON IN LIST VIEW
    # =========================
    def bulk_upload_button(self, obj):
        return mark_safe(
            '<a class="button" style="background:#417690;color:white;padding:5px 10px;border-radius:5px;text-decoration:none;" href="/admin/core/jobcard/bulk-upload/">📥 Bulk Upload</a>'
        )

    bulk_upload_button.short_description = "Bulk Upload"

    # =========================
    # CUSTOM URL INSIDE ADMIN
    # =========================
    def get_urls(self):
        urls = super().get_urls()
        custom_urls = [
            path(
                'bulk-upload/',
                self.admin_site.admin_view(self.bulk_upload_view),
                name='jobcard_bulk_upload'
            )
        ]
        return custom_urls + urls

    def bulk_upload_view(self, request):
        from django.shortcuts import render
        from .bulk_upload import process_jobcard_upload

        context = {}

        if request.method == "POST":
            file = request.FILES['file']
            result = process_jobcard_upload(file)
            context = result

        return TemplateResponse(request, "admin/bulk_upload.html", context)

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

    # OPTIONAL: show nice layout in form
    fieldsets = (
        ("Basic Info", {
            "fields": ("job_card_no", "SKU", "PO_No", "po_date", "month")
        }),
        ("Production Details", {
            "fields": ("order_qty", "ups", "wastage")
        }),
        ("Material Info", {
            "fields": ("material", "colour", "application")
        }),
        ("Machine Info", {
            "fields": ("machine_name", "department", "die_cutting")
        }),
        ("Other Info", {
            "fields": ("destination", "remarks", "status", "is_active")
        }),
    )


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
        'waste_qty'
    )

    list_filter = ('date', 'shift', 'machine')
    search_fields = ('job_card__job_card_no',)


# =========================
# DISPATCH ADMIN
# =========================

@admin.register(Dispatch)
class DispatchAdmin(admin.ModelAdmin):
    list_display = (
        'job_card',
        'dispatch_date',
        'dispatch_qty'
    )

    list_filter = ('dispatch_date',)


# =========================
# MASTER TABLES
# =========================

admin.site.register(Machine)
admin.site.register(Department)
admin.site.register(Material)


# =========================
# ADMIN HEADER (ERP BRANDING)
# =========================

admin.site.site_header = "Offset ERP System"
admin.site.site_title = "Offset ERP"
admin.site.index_title = "Production Dashboard"