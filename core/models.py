from django.db import models, transaction, IntegrityError
from django.core.exceptions import ValidationError
from django.db.models import Sum
from django.contrib.auth import get_user_model


# =========================
# MASTER TABLES
# =========================

class Machine(models.Model):
    name = models.CharField(max_length=100, unique=True)
    standard_impressions_per_hour = models.FloatField(default=4000, help_text="Standard printing speed in impressions per hour")
    standard_setup_minutes_per_color = models.FloatField(
        default=15,
        help_text="Default setup/make-ready minutes per color for planning"
    )

    is_active = models.BooleanField(default=True)

    def __str__(self):
        return self.name


class Department(models.Model):
    name = models.CharField(max_length=100, unique=True)

    def __str__(self):
        return self.name


class Material(models.Model):
    name = models.CharField(max_length=100, unique=True)

    def __str__(self):
        return self.name


class Operator(models.Model):
    name = models.CharField(max_length=100)
    employee_code = models.CharField(max_length=50, null=True, blank=True)
    is_active = models.BooleanField(default=True)

    def __str__(self):
        return self.name


class SequenceCounter(models.Model):
    """Generic counters for business document serials (e.g., JC numbers)."""

    key = models.CharField(max_length=50, unique=True)
    last_value = models.PositiveIntegerField(default=0)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return f"{self.key}: {self.last_value}"

# =========================
# JOB CARD
# =========================

class JobCard(models.Model):
    job_card_no = models.CharField(max_length=50, unique=True)

    month = models.CharField(max_length=20, null=True, blank=True)
    po_date = models.DateField(null=True, blank=True)
    PO_No = models.CharField(max_length=50, null=True, blank=True)

    SKU = models.CharField(max_length=100)

    material = models.ForeignKey(Material, on_delete=models.SET_NULL, null=True, blank=True)

    colour = models.CharField(max_length=20, null=True, blank=True, help_text="Supports values like 4, 1+1, 2+0")
    application = models.CharField(max_length=100, null=True, blank=True)

    order_qty = models.IntegerField()

    total_impressions_required = models.IntegerField(
        null=True, 
        blank=True,
        help_text="Total impressions required for this job (manually entered based on machine config - 1/2/5 color, front-back, etc.)"
    )

    estimated_run_time_minutes = models.FloatField(
        null=True,
        blank=True,
        help_text="Auto-estimated run time in minutes from impressions and machine speed"
    )
    estimated_setup_time_minutes = models.FloatField(
        null=True,
        blank=True,
        help_text="Auto-estimated setup time in minutes from colors and machine setup rate"
    )
    estimated_total_time_minutes = models.FloatField(
        null=True,
        blank=True,
        help_text="Auto-estimated total planned time in minutes (run + setup)"
    )
    production_tolerance_percent = models.FloatField(
        default=5,
        help_text="Allowed extra production over planned sheets in percent"
    )

    ups = models.IntegerField(null=True, blank=True)
    print_sheet_size = models.CharField(max_length=50, null=True, blank=True)

    wastage = models.IntegerField(default=0,help_text="in Sheets")

    purchase_sheet_size = models.CharField(max_length=50, null=True, blank=True)
    purchase_sheet_ups = models.IntegerField(null=True, blank=True)

    remarks = models.TextField(null=True, blank=True)

    destination = models.CharField(max_length=100, null=True, blank=True)

    machine_name = models.ForeignKey('Machine', on_delete=models.SET_NULL, null=True, blank=True)
    department = models.ForeignKey(Department, on_delete=models.SET_NULL, null=True, blank=True)

    die_cutting = models.CharField(max_length=100, null=True, blank=True)

    is_print_job = models.BooleanField(
        default=True,
        help_text="Uncheck for Cut & Pack jobs (no printing, dispatch directly against order qty)"
    )

    created_by = models.ForeignKey(
        'auth.User',
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='jobcards_created',
        editable=False,
    )

    short_close_closed_qty = models.PositiveIntegerField(
        default=0,
        help_text="Quantity manager has explicitly short-closed from pending completion gap"
    )
    short_close_wastage_qty = models.PositiveIntegerField(
        default=0,
        help_text="Short-close quantity moved to wastage bucket by manager decision"
    )
    short_close_closed_by = models.ForeignKey(
        'auth.User',
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='jobcards_short_closed',
        editable=False,
    )
    short_close_closed_at = models.DateTimeField(null=True, blank=True)
    short_close_close_reason = models.TextField(null=True, blank=True)

    is_active = models.BooleanField(default=True)
    created_at = models.DateTimeField(auto_now_add=True)

    status = models.CharField(max_length=20, default='Open')

    def __str__(self):
        return self.job_card_no

    # ===== ERP PROPERTIES =====

    @property
    def required_sheets(self):
        if self.ups:
            return self.order_qty / self.ups
        return 0
    
    @property
    def total_sheets_planned(self):
        return int (self.required_sheets + self.wastage)

    @property
    def tolerance_sheets(self):
        return int(round((self.total_sheets_planned * (self.production_tolerance_percent or 0)) / 100))

    @property
    def total_sheets_allowed_with_tolerance(self):
        return self.total_sheets_planned + self.tolerance_sheets

    @property
    def total_impressions_allowed_with_tolerance(self):
        tolerance = float(self.production_tolerance_percent or 0) / 100
        return int(round(self.total_impressions_required * (1 + tolerance)))

    @property
    def extra_sheets_used(self):
        total_consumed = self.productions.filter(is_active=True).aggregate(
            total_output=Sum('output_sheets'),
            total_waste=Sum('waste_sheets'),
        )
        consumed = (total_consumed['total_output'] or 0) + (total_consumed['total_waste'] or 0)
        return max(consumed - self.total_sheets_planned, 0)
    
    @property
    def total_production(self):
        return self.productions.filter(is_active=True).aggregate(total=Sum('output_sheets'))['total'] or 0

    @property
    def total_dispatch(self):
        return self.dispatch_set.filter(is_active=True).aggregate(total=Sum('dispatch_qty'))['total'] or 0

    @property
    def total_waste(self):
        return self.productions.filter(is_active=True).aggregate(total=Sum('waste_sheets'))['total'] or 0

    @property
    def balance_qty(self):
        return self.order_qty - self.total_dispatch

    @property
    def dispatch_completion_percent(self):
        if self.order_qty <= 0:
            return 0
        return round((self.total_dispatch / self.order_qty) * 100, 2)

    @property
    def short_close_qty(self):
        if self.job_status == "Completed" and self.total_dispatch < self.order_qty:
            gap = self.order_qty - self.total_dispatch
            return max(gap - (self.short_close_closed_qty or 0), 0)
        return 0

    @property
    def waste_percentage(self):
        if self.total_production == 0:
            return 0
        return round((self.total_waste / self.total_production) * 100, 2)

    @property
    def job_status(self):
        if not self.is_active:
            return "Archived"
        if self.order_qty == 0:
            return "Open"

        dispatch_ratio = self.dispatch_completion_percent

        if dispatch_ratio >= 95:
            return "Completed"
        elif self.total_production > 0:
            return "In Progress"
        return "Open"


# =========================
# PRODUCTION
# =========================

User = get_user_model()


class Production(models.Model):

    SHIFT_CHOICES = [
        ('A', 'Shift A'),
        ('B', 'Shift B'),
        
            ]

    job_card = models.ForeignKey('JobCard', on_delete=models.CASCADE, related_name='productions')

    date = models.DateField()
    shift = models.CharField(max_length=1, choices=SHIFT_CHOICES)

    machine = models.ForeignKey('Machine', on_delete=models.PROTECT)

    output_sheets = models.PositiveIntegerField()
    waste_sheets = models.PositiveIntegerField(default=0)

    WASTE_CHOICES = [
        ('paper_jam', 'Paper Jam (Affects OEE Quality)'),
        ('color_issue', 'Color/Registration Issue (Affects OEE Quality)'),
        ('material_defect', 'Material Defect (External - Excluded from OEE)'),
        ('operator_error', 'Operator Error (Affects OEE Quality)'),
        ('machine_issue', 'Machine Issue (Affects OEE Quality)'),
        ('other', 'Other (Affects OEE Quality)'),
    ]

    waste_reason = models.CharField(
        max_length=20,
        choices=WASTE_CHOICES,
        null=True,
        blank=True,
        help_text="Primary reason for waste"
    )

    impressions = models.PositiveIntegerField(help_text="Total impressions produced (sheets × passes)")

    planned_time = models.FloatField(help_text="in minutes")
    run_time = models.FloatField(help_text="in minutes")
    downtime = models.FloatField(default=0,help_text="in minutes")
    setup_time = models.FloatField(default=0,help_text="in minutes")

    DOWNTIME_CHOICES = [
        ('setup', 'Setup Time (Planned - Excluded from OEE)'),
        ('maintenance', 'Maintenance (Planned - Excluded from OEE)'),
        ('breakdown', 'Machine Breakdown (Unplanned - Affects OEE)'),
        ('material', 'Material Issue (External - Excluded from OEE)'),
        ('operator', 'Operator Issue (Affects OEE)'),
        ('other', 'Other (Affects OEE)'),
    ]

    downtime_category = models.CharField(
        max_length=20,
        choices=DOWNTIME_CHOICES,
        null=True,
        blank=True,
        help_text="Category of downtime"
    )

    ideal_run_rate = models.FloatField(null=True, blank=True)

    operator = models.ForeignKey('Operator',on_delete=models.SET_NULL,null=True,blank=True,limit_choices_to={'is_active': True})

    created_by = models.ForeignKey(
        'auth.User',
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='productions_created',
        editable=False,
    )

    is_active = models.BooleanField(default=True)
    created_at = models.DateTimeField(auto_now_add=True)

    #Calculated Fields

    @property
    def pcs_produced(self):
        if self.job_card.ups:
            return self.output_sheets * self.job_card.ups
        return 0
    
    @property
    def good_sheets(self):
        return self.output_sheets
    
    @property
    def total_sheets(self):
        return self.output_sheets + self.waste_sheets

    def clean(self):
        errors = {}

        existing = Production.objects.filter(job_card=self.job_card, is_active=True)\
            .exclude(id=self.id)\
            .aggregate(
                total_output=Sum('output_sheets'),
                total_waste=Sum('waste_sheets')
        )

        existing_output = existing['total_output'] or 0
        existing_waste = existing['total_waste'] or 0

        total_existing_consumption = existing_output + existing_waste
        current_consumption = (self.output_sheets or 0) + (self.waste_sheets or 0)

    # 🔴 MAIN VALIDATION (FIXED)
        if total_existing_consumption + current_consumption > self.job_card.total_sheets_allowed_with_tolerance:
            errors['output_sheets'] = (
                "Total sheets (production + waste) exceed allowed sheets with tolerance! "
                f"Allowed: {self.job_card.total_sheets_allowed_with_tolerance}"
            )

        # Impressions validation
        if self.impressions <= 0:
            errors['impressions'] = "Impressions must be greater than 0"
        if self.impressions < self.output_sheets:
            errors['impressions'] = "Impressions should be at least equal to output sheets"

        existing_impressions = Production.objects.filter(
            job_card=self.job_card,
            is_active=True,
        ).exclude(id=self.id).aggregate(total_impressions=Sum('impressions'))
        total_existing_impressions = existing_impressions['total_impressions'] or 0
        total_impressions = total_existing_impressions + (self.impressions or 0)
        allowed_impressions = self.job_card.total_impressions_allowed_with_tolerance
        if total_impressions > allowed_impressions:
            errors['impressions'] = (
                "Total impressions exceed allowed tolerance. "
                f"Allowed: {allowed_impressions}"
            )

    # ⏱ TIME VALIDATIONS
        # Overruns are allowed. Run time can exceed planned allocation for a session.

    # ⚙️ RATE VALIDATION
        if self.ideal_run_rate is not None and self.ideal_run_rate <= 0:
            errors['ideal_run_rate'] = "Must be > 0"

    # 🔥 VERY IMPORTANT (OUTSIDE ALL BLOCKS)
        if errors:
         raise ValidationError(errors)

    def save(self, *args, **kwargs):
        if not self.ideal_run_rate and self.machine:
            self.ideal_run_rate = self.machine.standard_impressions_per_hour

        self.full_clean()

        try:
            with transaction.atomic():
                super().save(*args, **kwargs)
        except IntegrityError:
            raise ValidationError("DB error while saving Production")

    # ===== OEE =====

    @property
    def expected_impressions(self):
        """Expected impressions based on machine capacity and run time"""
        if not self.machine or not self.machine.standard_impressions_per_hour:
            return 0
        run_time_hours = self.run_time / 60
        return self.machine.standard_impressions_per_hour * run_time_hours

    @property
    def availability(self):
        if self.planned_time == 0:
            return 0
        unplanned_downtime = self.unplanned_downtime_minutes
        return (self.planned_time - unplanned_downtime) / self.planned_time

    @property
    def unplanned_downtime_minutes(self):
        """Unplanned downtime minutes used in OEE availability logic."""
        unplanned_categories = {'breakdown', 'operator', 'other'}
        detail_rows = list(self.downtime_entries.all())
        if detail_rows:
            return float(sum(
                row.minutes for row in detail_rows
                if row.category in unplanned_categories
            ))
        return float(self.downtime or 0) if self.downtime_category in unplanned_categories else 0.0

    @property
    def downtime_breakdown_text(self):
        """Readable downtime split for UI/reporting, with fallback for legacy rows."""
        detail_rows = list(self.downtime_entries.all())
        if detail_rows:
            labels = dict(self.DOWNTIME_CHOICES)
            return ', '.join(
                f"{labels.get(row.category, row.category)}: {row.minutes:g}m"
                for row in detail_rows
            )
        if self.downtime and self.downtime_category:
            labels = dict(self.DOWNTIME_CHOICES)
            return f"{labels.get(self.downtime_category, self.downtime_category)}: {float(self.downtime):g}m"
        return '-'

    @property
    def performance(self):
        if not self.machine or not self.machine.standard_impressions_per_hour:
            return 0
        
        run_time_hours = self.run_time/60

        expected_impressions = self.machine.standard_impressions_per_hour * run_time_hours
        if expected_impressions == 0:
            return 0

        return self.impressions / expected_impressions
    


    @property
    def quality(self):
        if self.total_sheets == 0:
            return 0
        # OEE Quality excludes wastes not caused by production process
        # Only count wastes that affect machine/process quality
        quality_affecting_wastes = ['paper_jam', 'color_issue', 'operator_error', 'machine_issue', 'other']
        quality_waste = self.waste_sheets if self.waste_reason in quality_affecting_wastes else 0
        good_sheets = self.output_sheets  # assuming output_sheets are good
        total_quality_sheets = good_sheets + quality_waste
        if total_quality_sheets == 0:
            return 0
        return good_sheets / total_quality_sheets
    

    @property
    def oee(self):
        return round((self.availability * self.performance * self.quality),2)
    


    @property
    def overrun_minutes(self):
        """Minutes by which total time exceeded planned time (0 if on schedule)."""
        total = (self.run_time or 0) + (self.downtime or 0) + (self.setup_time or 0)
        return max(0, total - (self.planned_time or 0))

    @property
    def actual_total_time_minutes(self):
        """Actual consumed time for this production entry."""
        return (self.run_time or 0) + (self.downtime or 0) + (self.setup_time or 0)

    @property
    def planned_variance_minutes(self):
        """Actual minus planned. Positive means overrun, negative means underrun."""
        return self.actual_total_time_minutes - (self.planned_time or 0)

    def operator_efficiency(self):
        if self.run_time and self.ideal_run_rate:
            run_time_hours = self.run_time/60
            expected_impressions = self.ideal_run_rate * run_time_hours
            if expected_impressions == 0:
                return 0
            return (self.impressions / expected_impressions) * 100
        return 0

    def __str__(self):
        return f"{self.job_card.job_card_no} - {self.date}"


class ProductionDowntime(models.Model):
    """Detailed downtime rows to capture multiple reasons per production entry."""

    production = models.ForeignKey(
        Production,
        on_delete=models.CASCADE,
        related_name='downtime_entries'
    )
    category = models.CharField(max_length=20, choices=Production.DOWNTIME_CHOICES)
    minutes = models.FloatField(help_text="Downtime minutes for this category")
    note = models.CharField(max_length=200, null=True, blank=True)

    class Meta:
        ordering = ['id']

    def __str__(self):
        return f"{self.production} - {self.get_category_display()} ({self.minutes:g}m)"

# ========================= 
# DISPATCH 
# =========================


class Dispatch(models.Model):

    job_card = models.ForeignKey(JobCard, on_delete=models.CASCADE)

    dc_no = models.CharField(
        max_length=50,
        null=True,
        blank=True,
        help_text="Dispatch Challan Number (can be shared across multiple Job Cards)"

    )

    dispatch_date = models.DateField()

    dispatch_qty = models.IntegerField(default=0)

    created_by = models.ForeignKey(
        'auth.User',
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='dispatches_created',
        editable=False,
    )

    is_active = models.BooleanField(default=True)


    # =========================
    # VALIDATION ONLY
    # =========================
    def clean(self):
        errors = {}

        if self.dispatch_qty <= 0:
            errors['dispatch_qty'] = "Dispatch must be greater than 0"

        existing_dispatch = Dispatch.objects.filter(job_card=self.job_card, is_active=True)\
            .exclude(id=self.id)\
            .aggregate(total=Sum('dispatch_qty'))['total'] or 0

        total_after = existing_dispatch + (self.dispatch_qty or 0)

        if self.job_card.is_print_job:
            # Print jobs: dispatch must not exceed total produced pieces
            total_production = sum(
                p.pcs_produced for p in self.job_card.productions.filter(is_active=True)
            )
            if total_after > total_production:
                errors['dispatch_qty'] = "Dispatch cannot exceed total produced quantity!"
        else:
            # Cut & Pack jobs: dispatch directly against order qty (no production entry needed)
            if total_after > self.job_card.order_qty:
                errors['dispatch_qty'] = (
                    f"Dispatch ({total_after}) cannot exceed order qty ({self.job_card.order_qty}) "
                    f"for a Cut & Pack job!"
                )

        if errors:
            raise ValidationError(errors)

    # =========================
    # SAVE LOGIC
    # =========================
    def save(self, *args, **kwargs):
        self.full_clean()   # runs clean() + field validation
        super().save(*args, **kwargs)

    def __str__(self):
        return f"{self.job_card.job_card_no} - {self.dispatch_date}"


class ChangeLog(models.Model):
    ENTITY_CHOICES = [
        ('job_card', 'Job Card'),
        ('production', 'Production'),
        ('dispatch', 'Dispatch'),
    ]

    ACTION_CHOICES = [
        ('create', 'Created'),
        ('update', 'Updated'),
        ('delete', 'Deleted'),
        ('restore', 'Restored'),
    ]

    entity_type = models.CharField(max_length=20, choices=ENTITY_CHOICES)
    record_id = models.PositiveIntegerField()
    record_label = models.CharField(max_length=200)
    action = models.CharField(max_length=20, choices=ACTION_CHOICES)
    changed_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='change_logs')
    change_reason = models.TextField(blank=True)
    field_changes = models.JSONField(default=dict, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['-created_at']
        indexes = [
            models.Index(fields=['entity_type', 'record_id']),
            models.Index(fields=['created_at']),
        ]

    def __str__(self):
        return f"{self.get_entity_type_display()} {self.record_label} - {self.get_action_display()}"


# =========================
# USER ROLES & PERMISSIONS
# =========================

class UserProfile(models.Model):
    """Extended user model with role-based access control"""
    
    ROLE_CHOICES = [
        ('admin', 'Admin — Full system access & configuration'),
        ('manager', 'Manager — Overall oversight (jobs, production, dispatch, reports)'),
        ('planner', 'Planner — Create & manage job cards, view analytics'),
        ('production', 'Production Supervisor — Manage production entries & team'),
        ('operator', 'Machine Operator — Production entry only'),
        ('dispatch', 'Dispatch Coordinator — Dispatch approval & tracking'),
        ('qc', 'QC Inspector — Quality checks & approvals'),
        ('storekeeper', 'Store Keeper — Material & inventory management'),
        ('finance', 'Finance Viewer — Read-only analytics & reports'),
    ]
    
    user = models.OneToOneField(User, on_delete=models.CASCADE, related_name='profile')
    role = models.CharField(max_length=20, choices=ROLE_CHOICES, default='operator')
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    
    class Meta:
        verbose_name = "User Profile"
        verbose_name_plural = "User Profiles"
    
    def __str__(self):
        return f"{self.user.username} — {self.get_role_display()}"
    
    # Permission helpers
    def can_edit_jobcard(self):
        """Can create/edit job cards"""
        return self.role in ('admin', 'manager', 'planner')
    
    def can_edit_production(self):
        """Can log production data"""
        return self.role in ('admin', 'manager', 'production', 'operator')
    
    def can_approve_dispatch(self):
        """Can approve/edit dispatch"""
        return self.role in ('admin', 'manager', 'dispatch')
    
    def can_view_analytics(self):
        """Can view dashboard and analytics"""
        return self.role in ('admin', 'manager', 'planner', 'production', 'dispatch', 'finance')
    
    def can_manage_masters(self):
        """Can manage machines, operators, materials, departments"""
        return self.role in ('admin', 'manager')
    
    def can_approve_qc(self):
        """Can perform QC checks"""
        return self.role in ('admin', 'qc')
    
    def can_manage_operators(self):
        """Can assign operators to shifts/jobs"""
        return self.role in ('admin', 'manager', 'production')

    def can_archive_records(self):
        """Can archive and restore operational records"""
        return self.role in ('admin', 'manager')
    
    def can_view_reports(self):
        """Can view financial/operational reports"""
        return self.role in ('admin', 'manager', 'finance')


# =========================
# EDIT OVERRIDE REQUESTS
# =========================

class EditOverrideRequest(models.Model):
    ENTITY_CHOICES = [
        ('production', 'Production'),
        ('dispatch', 'Dispatch'),
    ]
    STATUS_CHOICES = [
        ('pending', 'Pending Review'),
        ('approved', 'Approved'),
        ('rejected', 'Rejected'),
    ]

    entity_type = models.CharField(max_length=20, choices=ENTITY_CHOICES)
    record_id = models.PositiveIntegerField()
    record_label = models.CharField(max_length=200)
    requested_by = models.ForeignKey(
        User, on_delete=models.CASCADE, related_name='override_requests'
    )
    reason = models.TextField()
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='pending')
    reviewed_by = models.ForeignKey(
        User, on_delete=models.SET_NULL, null=True, blank=True, related_name='reviewed_overrides'
    )
    review_note = models.TextField(blank=True)
    created_at = models.DateTimeField(auto_now_add=True)
    reviewed_at = models.DateTimeField(null=True, blank=True)
    expires_at = models.DateTimeField(null=True, blank=True)

    class Meta:
        ordering = ['-created_at']
        indexes = [
            models.Index(fields=['entity_type', 'record_id', 'requested_by']),
            models.Index(fields=['status', 'created_at']),
        ]

    def __str__(self):
        return (
            f"{self.get_entity_type_display()} #{self.record_id}"
            f" — {self.get_status_display()} — {self.requested_by}"
        )

    @property
    def is_valid_for_edit(self):
        from django.utils import timezone as _tz
        return (
            self.status == 'approved'
            and self.expires_at is not None
            and self.expires_at > _tz.now()
        )


# =========================
# SHIFT CONFIG
# =========================

class ShiftConfig(models.Model):
    """Net available hours per shift per day of week (week-wise config)."""

    DAY_CHOICES = [
        (0, 'Monday'),
        (1, 'Tuesday'),
        (2, 'Wednesday'),
        (3, 'Thursday'),
        (4, 'Friday'),
        (5, 'Saturday'),
        (6, 'Sunday'),
    ]

    SHIFT_CHOICES = [
        ('A', 'Shift A'),
        ('B', 'Shift B'),
    ]

    day_of_week = models.IntegerField(choices=DAY_CHOICES)
    shift = models.CharField(max_length=1, choices=SHIFT_CHOICES)
    effective_from = models.DateField(null=True, blank=True)
    effective_to = models.DateField(null=True, blank=True)
    net_hours = models.FloatField(
        default=11.0,
        help_text="Net available production hours after breaks (e.g. 11 for 12hr shift - 60min break)"
    )

    class Meta:
        unique_together = ('day_of_week', 'shift', 'effective_from', 'effective_to')
        ordering = ['day_of_week', 'shift']

    def __str__(self):
        return f"{self.get_day_of_week_display()} — Shift {self.shift}: {self.net_hours}h"


class MachineWorkSchedule(models.Model):
    """Defines which machines are OFF on which day+shift (default: all machines work every day)."""

    DAY_CHOICES = ShiftConfig.DAY_CHOICES
    SHIFT_CHOICES = ShiftConfig.SHIFT_CHOICES

    machine = models.ForeignKey('Machine', on_delete=models.CASCADE, related_name='work_schedule')
    day_of_week = models.IntegerField(choices=DAY_CHOICES)
    shift = models.CharField(max_length=1, choices=SHIFT_CHOICES)
    effective_from = models.DateField(null=True, blank=True)
    effective_to = models.DateField(null=True, blank=True)
    is_working = models.BooleanField(
        default=False,
        help_text="Uncheck = machine is OFF on this day+shift (e.g. GTO off Friday, on Sunday)"
    )

    class Meta:
        unique_together = ('machine', 'day_of_week', 'shift', 'effective_from', 'effective_to')
        ordering = ['machine', 'day_of_week', 'shift']

    def __str__(self):
        status = 'Working' if self.is_working else 'OFF'
        return f"{self.machine.name} — {self.get_day_of_week_display()} Shift {self.shift}: {status}"