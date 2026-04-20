from django.contrib import admin
from django.contrib.auth.admin import UserAdmin as BaseUserAdmin
from django.contrib.auth import get_user_model
from django.urls import path, reverse
from django.template.response import TemplateResponse
from django.utils.html import format_html
from django.utils.safestring import mark_safe


from .models import JobCard, Production, ProductionDowntime, Dispatch, Machine, Department, Material, Operator, UserProfile, ChangeLog, EditOverrideRequest

User = get_user_model()


# =========================
# INLINE MODELS
# =========================

class ProductionInline(admin.TabularInline):
    model = Production
    extra = 1


class DispatchInline(admin.TabularInline):
    model = Dispatch
    extra = 1


class ProductionDowntimeInline(admin.TabularInline):
    model = ProductionDowntime
    extra = 1


@admin.register(ChangeLog)
class ChangeLogAdmin(admin.ModelAdmin):
    list_display = ('created_at', 'entity_type', 'record_label', 'action', 'changed_by')
    list_filter = ('entity_type', 'action', 'created_at')
    search_fields = ('record_label', 'changed_by__username', 'change_reason')
    readonly_fields = ('entity_type', 'record_id', 'record_label', 'action', 'changed_by', 'change_reason', 'field_changes', 'created_at')

    def has_add_permission(self, request):
        return False

    def has_change_permission(self, request, obj=None):
        return False


@admin.register(EditOverrideRequest)
class EditOverrideRequestAdmin(admin.ModelAdmin):
    list_display = ('created_at', 'entity_type', 'record_label', 'requested_by', 'status', 'reviewed_by', 'expires_at')
    list_filter = ('status', 'entity_type', 'created_at')
    search_fields = ('record_label', 'requested_by__username', 'reason')
    readonly_fields = ('entity_type', 'record_id', 'record_label', 'requested_by', 'reason', 'status',
                       'reviewed_by', 'review_note', 'created_at', 'reviewed_at', 'expires_at')

    def has_add_permission(self, request):
        return False

    def has_change_permission(self, request, obj=None):
        return False


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
        'total_impressions_required',
        'total_sheets_planned',
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
                "total_impressions_required",
                "wastage"
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

    inlines = [ProductionDowntimeInline]

    list_display = (
    'job_card',
    'date',
    'shift',
    'machine',
    'output_sheets',
    'waste_sheets',
    'waste_reason',
    'pcs_produced',
    'impressions',
    'oee'
)

    list_filter = (
        'date',
        'shift',
        'machine',
        'operator',
        'waste_reason',
        'downtime_category',
    )

    search_fields = (
        'job_card__job_card_no',
        'machine__name',
        'operator__name',
    )

    autocomplete_fields = ['job_card', 'machine', 'operator']

    date_hierarchy = 'date'

    fieldsets = (
        ("Production Details", {
            "fields": (
                "job_card",
                "date",
                "shift",
                "machine",
                "operator"
            )
        }),
        ("Output & Waste", {
            "fields": (
                "output_sheets",
                "waste_sheets",
                "waste_reason",
                "impressions"
            )
        }),
        ("Time Tracking", {
            "fields": (
                "planned_time",
                "run_time",
                "downtime",
                "downtime_category",
                "setup_time"
            )
        }),
    )
    
@admin.register(Operator)
class OperatorAdmin(admin.ModelAdmin):
    list_display = ['name', 'employee_code', 'is_active']
    search_fields = ['name', 'employee_code']
    list_filter = ['is_active']


@admin.register(Machine)
class MachineAdmin(admin.ModelAdmin):
    list_display = ['name', 'standard_impressions_per_hour', 'is_active']
    search_fields = ['name']

@admin.register(Department)
class DepartmentAdmin(admin.ModelAdmin):
    search_fields = ['name']


# =========================
# USER PROFILE & RBAC ADMIN
# =========================

@admin.register(UserProfile)
class UserProfileAdmin(admin.ModelAdmin):
    """Manage user roles and permissions"""
    list_display = ['username', 'email', 'role_display', 'created_at']
    list_filter = ['role', 'created_at']
    search_fields = ['user__username', 'user__email']
    readonly_fields = ['user', 'created_at', 'updated_at']
    
    def username(self, obj):
        return obj.user.username
    username.short_description = 'Username'
    
    def email(self, obj):
        return obj.user.email
    email.short_description = 'Email'
    
    def role_display(self, obj):
        colors = {
            'admin': 'darkred',
            'manager': 'darkblue',
            'planner': 'darkblue',
            'production': 'darkgreen',
            'operator': 'darkgreen',
            'dispatch': 'darkorange',
            'finance': 'purple',
            'qc': 'darkred',
            'storekeeper': 'darkslategray',
        }
        color = colors.get(obj.role, 'gray')
        html = f'<span style="color:{color};font-weight:bold;">{obj.get_role_display()}</span>'
        return mark_safe(html)
    role_display.short_description = 'Role'


# Extend User admin to include UserProfile
class CustomUserAdmin(BaseUserAdmin):
    """Extended User admin with role assignment"""
    pass


# Unregister default User admin and register custom one
admin.site.unregister(User)
admin.site.register(User, CustomUserAdmin)

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
        'balance_check',
        'balance_qty_percentage'
    )

    list_filter = ('dispatch_date',)

    search_fields = ('job_card__job_card_no','dc_no',)

    
    def balance_qty_percentage(self, obj):
     if obj.job_card.order_qty == 0:
        return "0%"

     balance = obj.job_card.balance_qty
     percent = (balance / obj.job_card.order_qty) * 100

     return f"{round(percent, 2)}%"
    balance_qty_percentage.short_description = "Balance %"

    def balance_check(self, obj):
        return obj.job_card.balance_qty

    balance_check.short_description = "DC Balance"


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