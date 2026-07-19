from django.contrib import admin
from django.contrib.auth.admin import UserAdmin as BaseUserAdmin
from django.contrib.auth import get_user_model
from django.urls import path, reverse
from django.template.response import TemplateResponse
from django.utils.html import format_html
from django.utils.safestring import mark_safe


from .models import JobCard, Production, ProductionDowntime, Dispatch, Machine, Department, DeliveryLocation, PrintColor, ProductType, ApplicationType, Material, Operator, Supervisor, Sorter, UserProfile, ChangeLog, EditOverrideRequest, Vendor, Notification, NotificationEvent, NotificationRule, WorkflowTransition, NotificationRuleAuditLog, PasswordResetRequest

User = get_user_model()


@admin.register(Sorter)
class SorterAdmin(admin.ModelAdmin):
    list_display = ('name', 'employee_code', 'is_active')
    search_fields = ('name', 'employee_code')
    list_filter = ('is_active',)


@admin.register(Vendor)
class VendorAdmin(admin.ModelAdmin):
    list_display = ('name', 'is_active', 'created_at')
    search_fields = ('name',)


@admin.register(PrintColor)
class PrintColorAdmin(admin.ModelAdmin):
    list_display = ('name', 'is_active', 'sort_order', 'created_at')
    list_filter = ('is_active',)
    search_fields = ('name',)
    ordering = ('sort_order', 'name')


@admin.register(Notification)
class NotificationAdmin(admin.ModelAdmin):
    list_display = ('title', 'user', 'event_type', 'is_read', 'created_at')
    list_filter = ('event_type', 'is_read', 'created_at')
    search_fields = ('title', 'message', 'user__username')
    readonly_fields = ('created_at',)


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
        'operator',
        'supervisor',
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
        'supervisor',
        'waste_reason',
        'downtime_category',
    )

    search_fields = (
        'job_card__job_card_no',
        'machine__name',
        'operator__name',
        'supervisor__name',
    )

    autocomplete_fields = ['job_card', 'machine', 'operator', 'supervisor']

    date_hierarchy = 'date'

    fieldsets = (
        ("Production Details", {
            "fields": (
                "job_card",
                "date",
                "shift",
                "machine",
                "operator",
                "supervisor",
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
                "downtime_minutes",
                "downtime_category",
                "make_ready_time"
            )
        }),
    )
    
@admin.register(Operator)
class OperatorAdmin(admin.ModelAdmin):
    list_display = ['name', 'employee_code', 'is_active']
    search_fields = ['name', 'employee_code']
    list_filter = ['is_active']


@admin.register(Supervisor)
class SupervisorAdmin(admin.ModelAdmin):
    list_display = ['name', 'employee_code', 'is_active']
    search_fields = ['name', 'employee_code']
    list_filter = ['is_active']


@admin.register(Machine)
class MachineAdmin(admin.ModelAdmin):
    list_display = [
        'name', 'machine_type', 'machine_group_code', 'default_colors', 'operational_colors',
        'standard_impressions_per_hour', 'plate_life_impressions', 'is_active',
    ]
    list_filter = ['machine_type', 'machine_group_code', 'is_active']
    search_fields = ['name', 'machine_group_code']
    fieldsets = (
        (None, {'fields': ('name', 'machine_type', 'machine_group_code', 'is_active')}),
        ('Colour (offset printing only)', {'fields': ('default_colors', 'operational_colors')}),
        ('Print size range (mm)', {
            'fields': (
                ('min_print_length_mm', 'min_print_width_mm'),
                ('max_print_length_mm', 'max_print_width_mm'),
            ),
        }),
        ('Speed & setup', {
            'fields': ('standard_impressions_per_hour', 'standard_setup_minutes_per_color', 'plate_life_impressions'),
        }),
    )

@admin.register(Department)
class DepartmentAdmin(admin.ModelAdmin):
    search_fields = ['name']


@admin.register(DeliveryLocation)
class DeliveryLocationAdmin(admin.ModelAdmin):
    search_fields = ['name']


@admin.register(ProductType)
class ProductTypeAdmin(admin.ModelAdmin):
    search_fields = ['name']


@admin.register(ApplicationType)
class ApplicationTypeAdmin(admin.ModelAdmin):
    search_fields = ['name']


# =========================
# USER PROFILE & RBAC ADMIN
# =========================

@admin.register(UserProfile)
class UserProfileAdmin(admin.ModelAdmin):
    """Manage user roles and permissions"""
    list_display = ['username', 'email', 'role_display', 'department', 'manager', 'supervisor', 'sku_review_status', 'created_at']
    list_filter = ['role', 'department', 'manager', 'supervisor', 'can_view_sku_master_review', 'can_approve_sku_master', 'created_at']
    search_fields = ['user__username', 'user__email']
    readonly_fields = ['user', 'created_at', 'updated_at']
    fieldsets = (
        ('User Info', {
            'fields': ('user', 'role', 'department', 'manager', 'supervisor')
        }),
        ('SKU Master Review Permissions', {
            'fields': ('can_view_sku_master_review', 'can_approve_sku_master'),
            'description': 'Grant custom access to SKU master review functions'
        }),
        ('Timestamps', {
            'fields': ('created_at', 'updated_at'),
            'classes': ('collapse',)
        }),
    )
    
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
    
    def sku_review_status(self, obj):
        """Display SKU master review permission status"""
        if obj.can_approve_sku_master:
            return '✓ View & Approve'
        elif obj.can_view_sku_master_review:
            return '✓ View Only'
        return '✗ No Access'
    sku_review_status.short_description = 'SKU Review Access'


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


# =========================
# RULE-BASED NOTIFICATION ADMIN
# =========================

@admin.register(NotificationEvent)
class NotificationEventAdmin(admin.ModelAdmin):
    list_display = ('code', 'name', 'module', 'is_active', 'created_at')
    list_filter = ('module', 'is_active', 'created_at')
    search_fields = ('code', 'name', 'description')
    fieldsets = (
        ('Event Info', {
            'fields': ('code', 'name', 'description', 'module', 'is_active')
        }),
        ('Templates', {
            'fields': ('title_template', 'message_template', 'link_template'),
            'description': 'Dynamic template strings using Django template language. Context has `instance` and `actor`.'
        }),
    )


@admin.register(NotificationRule)
class NotificationRuleAdmin(admin.ModelAdmin):
    list_display = ('event', 'enabled', 'recipient_type', 'role', 'user', 'department', 'priority', 'in_app_enabled')
    list_filter = ('event', 'enabled', 'recipient_type', 'priority')
    search_fields = ('event__code', 'event__name', 'role', 'user__username', 'department__name')
    fieldsets = (
        ('Rule Setting', {
            'fields': ('event', 'enabled', 'priority')
        }),
        ('Recipient Definition', {
            'fields': (
                'recipient_type', 'role', 'user', 'department',
                'send_to_creator', 'send_to_manager', 'send_to_supervisor', 'send_to_next_stage'
            )
        }),
        ('Constraints & Channels', {
            'fields': ('exclude_actor', 'email_enabled', 'sms_enabled', 'in_app_enabled')
        }),
    )

    def save_model(self, request, obj, form, change):
        old_obj = None
        if change:
            try:
                old_obj = NotificationRule.objects.get(pk=obj.pk)
            except NotificationRule.DoesNotExist:
                pass
        super().save_model(request, obj, form, change)
        
        # log change
        from core.notifications import log_rule_change
        action = 'update' if change else 'create'
        log_rule_change(request.user, obj, action, old_obj)

    def delete_model(self, request, obj):
        # log deletion
        from core.notifications import log_rule_change
        log_rule_change(request.user, obj, 'delete')
        super().delete_model(request, obj)


@admin.register(WorkflowTransition)
class WorkflowTransitionAdmin(admin.ModelAdmin):
    list_display = ('module', 'current_stage', 'action', 'next_stage', 'notify_role')
    list_filter = ('module', 'notify_role')
    search_fields = ('module', 'current_stage', 'action', 'next_stage', 'notify_role')


@admin.register(NotificationRuleAuditLog)
class NotificationRuleAuditLogAdmin(admin.ModelAdmin):
    list_display = ('timestamp', 'changed_by', 'action', 'rule_event', 'rule_recipient')
    list_filter = ('action', 'timestamp', 'changed_by')
    search_fields = ('changed_by__username', 'action')
    readonly_fields = ('rule', 'changed_by', 'action', 'old_values', 'new_values', 'timestamp')

    def has_add_permission(self, request):
        return False

    def has_change_permission(self, request, obj=None):
        return False

    def rule_event(self, obj):
        if obj.rule and obj.rule.event:
            return obj.rule.event.name
        # Fallback to old_values or new_values
        event_id = obj.old_values.get('event') or obj.new_values.get('event')
        if event_id:
            try:
                return NotificationEvent.objects.get(pk=int(event_id)).name
            except Exception:
                return f"Event ID: {event_id}"
        return '-'
    rule_event.short_description = 'Event'

    def rule_recipient(self, obj):
        if obj.rule:
            return f"{obj.rule.get_recipient_type_display()}: {obj.rule.role or obj.rule.user or obj.rule.department or ''}"
        # Fallback to old_values
        rec_type = obj.old_values.get('recipient_type')
        if rec_type:
            target = obj.old_values.get('role') or obj.old_values.get('user') or obj.old_values.get('department') or ''
            return f"{rec_type}: {target}"
        return '-'
    rule_recipient.short_description = 'Recipient'


@admin.register(PasswordResetRequest)
class PasswordResetRequestAdmin(admin.ModelAdmin):
    list_display = ('username_or_email', 'created_at', 'resolved', 'resolved_at')
    list_filter = ('resolved', 'created_at')
    search_fields = ('username_or_email',)
    actions = ['mark_resolved']

    def mark_resolved(self, request, queryset):
        from django.utils import timezone
        queryset.update(resolved=True, resolved_at=timezone.now())
    mark_resolved.short_description = "Mark selected requests as resolved"