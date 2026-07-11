import math
import re

from django.db import models, transaction, IntegrityError
from django.core.exceptions import ValidationError
from django.db.models import Sum
from django.contrib.auth import get_user_model

from production.services import OEECalculator


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
    plate_life_impressions = models.PositiveIntegerField(
        default=25000,
        help_text="Impressions one plate set can run before a replacement set is needed",
    )

    is_active = models.BooleanField(default=True)

    def __str__(self):
        return self.name


class Department(models.Model):
    name = models.CharField(max_length=100, unique=True)

    def __str__(self):
        return self.name


class DeliveryLocation(models.Model):
    name = models.CharField(max_length=120, unique=True)

    class Meta:
        ordering = ['name']
        verbose_name = 'Delivery Location'
        verbose_name_plural = 'Delivery Locations'

    def __str__(self):
        return self.name


class ProductType(models.Model):
    name = models.CharField(max_length=100, unique=True)

    class Meta:
        ordering = ['name']
        verbose_name = 'Product Type'
        verbose_name_plural = 'Product Types'

    def __str__(self):
        return self.name


class PrintColor(models.Model):
    """Production print-colour patterns (1, 2, 4, 1+1, 2+1, …). Admin-managed master."""

    name = models.CharField(
        max_length=20,
        unique=True,
        help_text='Production colour pattern, e.g. 1, 2, 4, 1+1, 2+1',
    )
    is_active = models.BooleanField(default=True)
    sort_order = models.PositiveSmallIntegerField(default=0)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['sort_order', 'name']
        verbose_name = 'Print Color'
        verbose_name_plural = 'Print Colors'

    def __str__(self):
        return self.name

    @property
    def total_units(self):
        """Numeric colour units for setup/pass logic (1+1 → 2, 4 → 4)."""
        import re
        raw = (self.name or '').strip()
        match = re.fullmatch(r'(\d+)\s*\+\s*(\d+)', raw)
        if match:
            return int(match.group(1)) + int(match.group(2))
        match = re.fullmatch(r'(\d+)', raw)
        if match:
            return int(match.group(1))
        return 0


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


class Supervisor(models.Model):
    name = models.CharField(max_length=100)
    employee_code = models.CharField(max_length=50, null=True, blank=True)
    is_active = models.BooleanField(default=True)

    def __str__(self):
        return self.name


class Sorter(models.Model):
    name = models.CharField(max_length=100)
    employee_code = models.CharField(max_length=50, null=True, blank=True)
    is_active = models.BooleanField(default=True)

    def __str__(self):
        if self.employee_code:
            return f'{self.name} ({self.employee_code})'
        return self.name


class Vendor(models.Model):
    name = models.CharField(max_length=120, unique=True)
    is_active = models.BooleanField(default=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['name']
        verbose_name = 'Vendor'
        verbose_name_plural = 'Vendors'

    def __str__(self):
        return self.name


class SequenceCounter(models.Model):
    """Generic counters for business document serials (e.g., JC numbers)."""

    key = models.CharField(max_length=50, unique=True)
    last_value = models.PositiveIntegerField(default=0)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return f"{self.key}: {self.last_value}"


JOB_CARD_STATUS_CHOICES = [
    ('draft', 'Draft'),
    ('pending_data', 'Pending Data'),
    ('planning_approved', 'Planning Approved'),
    ('pending_qc', 'Pending QC'),
    ('qc_approved', 'QC Approved'),
    ('qc_rejected', 'QC Rejected'),
    ('pending_pm_approval', 'Pending Production Manager Approval'),
    ('production_approved', 'Production Approved'),
    ('pm_rejected', 'PM Rejected'),
    ('released', 'Released'),
    ('in_production', 'In Production'),
    ('completed', 'Completed'),
    ('closed', 'Closed'),
]

JOB_CARD_STATUS_ALIASES = {
    'open': 'pending_data',
    'pending': 'pending_data',
    'in progress': 'in_production',
    'in_progress': 'in_production',
    'approved': 'production_approved',
    'printed': 'released',
    'done': 'completed',
    'finished': 'completed',
    'archived': 'closed',
}

JOB_CARD_PLANNING_EDITABLE_STATUSES = {'draft', 'pending_data', 'qc_rejected'}
JOB_CARD_PLANNING_APPROVAL_STATUSES = {
    'planning_approved',
    'pending_qc',
    'qc_approved',
    'pending_pm_approval',
    'production_approved',
    'pm_rejected',
    'released',
    'in_production',
    'completed',
    'closed',
}
JOB_CARD_EXECUTION_STATUSES = {'in_production', 'completed', 'closed'}
JOB_CARD_PRODUCTION_START_STATUSES = {'released', 'in_production'}
JOB_CARD_DISPATCHABLE_STATUSES = {'in_production', 'completed', 'closed'}
JOB_CARD_PRINTABLE_STATUSES = {'production_approved', 'released', 'in_production', 'completed', 'closed'}
JOB_CARD_PLANNING_REQUIRED_FIELDS = (
    ('po_date', 'PO Received Date'),
    ('total_sheet_quantity', 'Total Sheet Quantity'),
    ('total_colors', 'Number of Colors'),
    ('wastage', 'Wastage'),
    ('machine_name', 'Machine Name'),
)

# =========================
# JOB CARD
# =========================

class JobCard(models.Model):
    job_card_no = models.CharField(max_length=50, unique=True)
    planning_job = models.OneToOneField(
        'planning.PlanningJob',
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='job_card',
    )

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
    plate_set_no = models.CharField(max_length=120, null=True, blank=True)

    wastage = models.IntegerField(default=0, help_text="in Sheets")
    total_sheet_quantity = models.PositiveIntegerField(null=True, blank=True)
    total_colors = models.PositiveIntegerField(null=True, blank=True)

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

    is_active = models.BooleanField(default=True, db_index=True)
    created_at = models.DateTimeField(auto_now_add=True, db_index=True)
    updated_at = models.DateTimeField(auto_now=True)

    status = models.CharField(max_length=40, choices=JOB_CARD_STATUS_CHOICES, default='pending_data', blank=True, db_index=True)

    def __str__(self):
        return self.job_card_no

    @property
    def workflow_status(self):
        raw_status = (self.status or '').strip().lower()
        return JOB_CARD_STATUS_ALIASES.get(raw_status, raw_status or 'pending_data')

    @property
    def workflow_status_label(self):
        labels = dict(JOB_CARD_STATUS_CHOICES)
        normalized_status = self.workflow_status
        return labels.get(normalized_status, normalized_status.replace('_', ' ').title())

    @property
    def machine_name_display(self):
        if self.machine_name_id:
            return str(self.machine_name)
        if self.planning_job and str(self.planning_job.machine_name or '').strip():
            return self.planning_job.machine_name
        return ''

    @property
    def total_sheet_quantity_display(self):
        if self.total_sheet_quantity is not None:
            return self.total_sheet_quantity
        if self.planning_job and self.planning_job.total_sheet_quantity is not None:
            return self.planning_job.total_sheet_quantity
        return None

    @property
    def total_colors_display(self):
        if self.total_colors is not None:
            return self.total_colors
        if self.planning_job and self.planning_job.number_of_colors is not None:
            return self.planning_job.number_of_colors
        return None

    @property
    def is_planning_editable(self):
        return self.workflow_status in JOB_CARD_PLANNING_EDITABLE_STATUSES

    @property
    def latest_rejection_reason(self):
        log = ChangeLog.objects.filter(
            entity_type='job_card',
            record_id=self.pk,
            action='reject',
        ).order_by('-created_at').first()
        if not log:
            return ''
        if log.change_reason:
            return log.change_reason
        note = log.field_changes.get('note') if isinstance(log.field_changes, dict) else None
        if isinstance(note, dict):
            return note.get('to') or ''
        return ''

    @property
    def po_received_date(self):
        if self.po_date:
            return self.po_date
        if self.created_at:
            return self.created_at.date()
        return None

    def planning_missing_fields(self):
        if self.workflow_status in JOB_CARD_PLANNING_EDITABLE_STATUSES:
            return []

        missing_fields = []
        for field_name, label in JOB_CARD_PLANNING_REQUIRED_FIELDS:
            if field_name == 'machine_name':
                if getattr(self, 'machine_name_id', None):
                    continue
                if self.planning_job and str(self.planning_job.machine_name or '').strip():
                    continue
                missing_fields.append(label)
                continue

            if field_name == 'total_sheet_quantity':
                if getattr(self, 'total_sheet_quantity', None) is not None:
                    continue
                if self.planning_job and self.planning_job.total_sheet_quantity is not None:
                    continue
                missing_fields.append(label)
                continue

            if field_name == 'total_colors':
                if getattr(self, 'total_colors', None) is not None:
                    continue
                if self.planning_job and self.planning_job.number_of_colors is not None:
                    continue
                missing_fields.append(label)
                continue

            if field_name in {'po_date', 'wastage'}:
                if getattr(self, field_name, None) is None:
                    missing_fields.append(label)
                continue

            value = getattr(self, field_name, None)
            if not str(value or '').strip():
                missing_fields.append(label)
        return missing_fields

    @property
    def wip_status_name(self):
        if hasattr(self, 'production_wip_status') and self.production_wip_status.status:
            return self.production_wip_status.status.name
        return 'Not Set'

    def planning_validation_errors(self):
        missing_fields = self.planning_missing_fields()
        if not missing_fields:
            return {}

        errors = {}
        for field_name, label in JOB_CARD_PLANNING_REQUIRED_FIELDS:
            if label not in missing_fields:
                continue
            errors[field_name] = f'{label} is required before the Job Card can move past planning approval.'
        return errors

    @property
    def number_of_colors(self):
        if self.total_colors is not None:
            return self.total_colors

        raw_value = (self.colour or '').strip()
        if not raw_value:
            return None

        plus_match = re.fullmatch(r'(\d+)\s*\+\s*(\d+)', raw_value)
        if plus_match:
            return int(plus_match.group(1)) + int(plus_match.group(2))

        single_match = re.fullmatch(r'(\d+)\s*(?:colou?r(?:s)?)?', raw_value, re.IGNORECASE)
        if single_match:
            return int(single_match.group(1))

        numbers = re.findall(r'\d+', raw_value)
        if len(numbers) == 1:
            return int(numbers[0])
        if len(numbers) == 2:
            return int(numbers[0]) + int(numbers[1])
        return None

    def can_print_job_card(self):
        return self.workflow_status in JOB_CARD_PRINTABLE_STATUSES and not self.planning_missing_fields()

    def clean(self):
        super().clean()
        planning_errors = self.planning_validation_errors()
        if planning_errors:
            raise ValidationError(planning_errors)

    def save(self, *args, **kwargs):
        update_fields = kwargs.get('update_fields')
        if update_fields is not None:
            update_fields = set(update_fields)

        self.status = self.workflow_status
        if update_fields is not None:
            update_fields.add('status')

        self.full_clean()

        if update_fields is not None:
            kwargs['update_fields'] = list(update_fields)
        return super().save(*args, **kwargs)

    # ===== ERP PROPERTIES =====

    @property
    def required_sheets(self):
        if self.ups:
            return self.order_qty / self.ups
        return 0
    
    @property
    def total_sheets_planned(self):
        if self.total_sheet_quantity is not None:
            return self.total_sheet_quantity
        return int(self.required_sheets + self.wastage)

    @property
    def tolerance_sheets(self):
        return int(round((self.total_sheets_planned * (self.production_tolerance_percent or 0)) / 100))

    @property
    def total_sheets_allowed_with_tolerance(self):
        return self.total_sheets_planned + self.tolerance_sheets

    @property
    def impression_pass_multiplier(self):
        """Compute the number of passes that should be applied when calculating impressions."""
        if self.planning_job and self.planning_job.print_passes:
            return 1

        if self.planning_job:
            front_pass = int(self.planning_job.front_pass or 0)
            back_pass = int(self.planning_job.back_pass or 0)
            if front_pass > 0 or back_pass > 0:
                return max(1, front_pass + back_pass)

        colour = (self.colour or '').strip().lower()
        if colour:
            match = re.fullmatch(r'(\d+)\s*\+\s*(\d+)', colour)
            if match:
                front = int(match.group(1))
                back = int(match.group(2))
                return max(1, front + back)
            match = re.search(r'(\d+)', colour)
            if match:
                return max(1, int(match.group(1)))

        return 1

    @property
    def total_impressions_allowed_with_tolerance(self):
        tolerance = float(self.production_tolerance_percent or 0) / 100
        return int(round(self.total_impressions_required * self.impression_pass_multiplier * (1 + tolerance)))

    @property
    def extra_sheets_used(self):
        total_consumed = self.printing_productions.aggregate(
            total_output=Sum('output_sheets'),
            total_waste=Sum('waste_sheets'),
        )
        consumed = (total_consumed['total_output'] or 0) + (total_consumed['total_waste'] or 0)
        return max(consumed - self.total_sheets_planned, 0)
    
    @property
    def total_production(self):
        return self.printing_productions.aggregate(total=Sum('output_sheets'))['total'] or 0

    @property
    def printing_productions(self):
        return self.productions.filter(is_active=True, entry_type='printing')

    @property
    def packing_productions(self):
        return self.productions.filter(is_active=True, entry_type='packing')

    @property
    def total_printed_pcs(self):
        return sum(p.pcs_produced for p in self.printing_productions)

    @property
    def total_production_pcs(self):
        return self.total_printed_pcs

    @property
    def total_packed_pcs(self):
        return self.packing_productions.aggregate(total=Sum('packing_qty'))['total'] or 0

    @property
    def total_sorting_waste_pcs(self):
        return self.packing_productions.aggregate(total=Sum('sorting_waste_qty'))['total'] or 0

    @property
    def total_packing_used_pcs(self):
        return int(self.total_packed_pcs or 0) + int(self.total_sorting_waste_pcs or 0)

    @property
    def packing_limit_pcs(self):
        if self.is_print_job:
            return int(self.total_printed_pcs or 0)
        return int(self.order_qty or 0)

    @property
    def remaining_packing_allowance_pcs(self):
        return max(0, self.packing_limit_pcs - self.total_packing_used_pcs)

    @property
    def process_type_label(self):
        return 'Print + Pack' if self.is_print_job else 'Cut & Pack'

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

    ENTRY_TYPE_CHOICES = [
        ('printing', 'Printing'),
        ('packing', 'Packing'),
    ]

    entry_type = models.CharField(
        max_length=20,
        choices=ENTRY_TYPE_CHOICES,
        default='printing',
        db_index=True,
    )

    job_card = models.ForeignKey('JobCard', on_delete=models.CASCADE, related_name='productions')

    date = models.DateField(db_index=True)
    shift = models.CharField(max_length=1, choices=SHIFT_CHOICES)

    machine = models.ForeignKey('Machine', on_delete=models.PROTECT, null=True, blank=True)

    output_sheets = models.PositiveIntegerField(default=0)
    waste_sheets = models.PositiveIntegerField(default=0)
    intermediate_pass = models.BooleanField(default=False, help_text="Mark as an intermediate print pass with no final usable output")
    print_pass_number = models.PositiveSmallIntegerField(
        null=True,
        blank=True,
        help_text='Which print pass this run belongs to (1..N from planning).',
    )

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

    impressions = models.PositiveIntegerField(
        default=0,
        help_text="Total impressions produced (sheets × passes)",
    )

    packing_qty = models.PositiveIntegerField(
        default=0,
        help_text="Good pieces packed (dispatchable)",
    )
    sorting_waste_qty = models.PositiveIntegerField(
        default=0,
        help_text="Pieces rejected during sorting",
    )
    sorter = models.ForeignKey(
        'Sorter',
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        limit_choices_to={'is_active': True},
    )

    planned_time = models.FloatField(default=0, help_text="in minutes")
    run_time = models.FloatField(default=0, help_text="in minutes")
    downtime_minutes = models.FloatField(default=0, help_text="in minutes")
    make_ready_time = models.FloatField(default=0, help_text="in minutes")

    DOWNTIME_CHOICES = [
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

    downtime_category_other = models.CharField(max_length=255, null=True, blank=True)
    waste_reason_other = models.CharField(max_length=255, null=True, blank=True)

    counter_start = models.PositiveIntegerField(default=0)
    counter_end = models.PositiveIntegerField(default=0)
    start_time = models.TimeField(null=True, blank=True)
    end_time = models.TimeField(null=True, blank=True)

    PRODUCTION_STATUS_CHOICES = [
        ('in_progress', 'In Progress'),
        ('completed', 'Completed'),
        ('hold', 'Hold'),
        ('cancelled', 'Cancelled'),
    ]
    status = models.CharField(max_length=20, choices=PRODUCTION_STATUS_CHOICES, default='in_progress')

    supervisor = models.ForeignKey(
        'Supervisor',
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        limit_choices_to={'is_active': True}
    )

    remark_notes = models.TextField(null=True, blank=True)
    change_reason = models.TextField(null=True, blank=True)

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

    is_active = models.BooleanField(default=True, db_index=True)
    created_at = models.DateTimeField(auto_now_add=True)

    #Calculated Fields

    @property
    def pcs_produced(self):
        if self.entry_type != 'printing':
            return 0
        if self.job_card.ups:
            return self.output_sheets * self.job_card.ups
        return 0
    
    @property
    def good_sheets(self):
        return self.output_sheets
    
    @property
    def total_sheets(self):
        return self.output_sheets + self.waste_sheets

    def minimum_impressions_for_output(self):
        """Calculate the minimum impressions required for the current output based on pass count."""
        if not self.job_card or self.output_sheets <= 0:
            return 0

        passes = 1
        if getattr(self.job_card, 'planning_job', None):
            front_pass = int(self.job_card.planning_job.front_pass or 0)
            back_pass = int(self.job_card.planning_job.back_pass or 0)
            if front_pass > 0 and back_pass > 0:
                passes = 2
        else:
            colour = (self.job_card.colour or '').strip()
            match = re.fullmatch(r'(\d+)\s*\+\s*(\d+)', colour)
            if match:
                front = int(match.group(1))
                back = int(match.group(2))
                if front > 0 and back > 0:
                    passes = 2

        return self.output_sheets * passes

    def clean(self):
        errors = {}

        if self.job_card and self.job_card.workflow_status not in JOB_CARD_PRODUCTION_START_STATUSES:
            errors['job_card'] = (
                'Production can only start after the Job Card has been released for execution.'
            )

        if self.entry_type == 'packing':
            packing_qty = int(self.packing_qty or 0)
            sorting_waste = int(self.sorting_waste_qty or 0)
            if packing_qty < 0 or sorting_waste < 0:
                errors['packing_qty'] = 'Packing and sorting waste quantities cannot be negative.'
            if packing_qty == 0 and sorting_waste == 0:
                errors['packing_qty'] = 'Enter packing qty and/or sorting waste qty.'
            if not self.sorter_id:
                errors['sorter'] = 'Sorter is required for packing entry.'
            if not self.shift:
                errors['shift'] = 'Shift is required.'

            existing_used = Production.objects.filter(
                job_card=self.job_card,
                is_active=True,
                entry_type='packing',
            ).exclude(id=self.id).aggregate(
                packed=Sum('packing_qty'),
                waste=Sum('sorting_waste_qty'),
            )
            already_used = int(existing_used['packed'] or 0) + int(existing_used['waste'] or 0)
            limit = self.job_card.packing_limit_pcs if self.job_card else 0
            if already_used + packing_qty + sorting_waste > limit:
                errors['packing_qty'] = (
                    f'Packing qty + sorting waste cannot exceed allowed limit ({limit:,} pcs). '
                    f'Already logged: {already_used:,}; remaining: {max(0, limit - already_used):,}.'
                )
            if errors:
                raise ValidationError(errors)
            return

        existing = Production.objects.filter(
            job_card=self.job_card,
            is_active=True,
            entry_type='printing',
        ).exclude(id=self.id).aggregate(
            total_output=Sum('output_sheets'),
            total_waste=Sum('waste_sheets'),
        )

        existing_output = existing['total_output'] or 0
        existing_waste = existing['total_waste'] or 0

        total_existing_consumption = existing_output + existing_waste
        current_consumption = (self.output_sheets or 0) + (self.waste_sheets or 0)

        if total_existing_consumption + current_consumption > self.job_card.total_sheets_allowed_with_tolerance:
            errors['output_sheets'] = (
                "Total sheets (production + waste) exceed allowed sheets with tolerance! "
                f"Allowed: {self.job_card.total_sheets_allowed_with_tolerance}"
            )

        if self.print_pass_number and self.job_card:
            from production.printing_pass_helpers import get_job_card_pass_count
            total_passes = get_job_card_pass_count(self.job_card)
            if self.print_pass_number < 1 or self.print_pass_number > total_passes:
                errors['print_pass_number'] = f'Print pass must be between 1 and {total_passes}.'
            elif self.print_pass_number < total_passes and (self.output_sheets or 0) > 0:
                errors['output_sheets'] = 'Good sheets are only allowed on the final print pass.'
            elif self.print_pass_number >= total_passes and (self.output_sheets or 0) <= 0:
                errors['output_sheets'] = 'Good sheets are required on the final print pass.'
            self.intermediate_pass = self.print_pass_number < total_passes

        # Impressions validation
        if self.impressions < 0:
            errors['impressions'] = "Impressions must be greater than or equal to 0"

        existing_impressions = Production.objects.filter(
            job_card=self.job_card,
            is_active=True,
            entry_type='printing',
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
        return OEECalculator.availability(self.run_time, self.downtime_minutes)

    @property
    def press_utilization(self):
        return OEECalculator.press_utilization(self.make_ready_time, self.run_time, self.downtime_minutes)

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
        return float(self.downtime_minutes or 0) if self.downtime_category in unplanned_categories else 0.0

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
        if self.downtime_minutes and self.downtime_category:
            labels = dict(self.DOWNTIME_CHOICES)
            return f"{labels.get(self.downtime_category, self.downtime_category)}: {float(self.downtime_minutes):g}m"
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
        total = (self.run_time or 0) + (self.downtime_minutes or 0) + (self.make_ready_time or 0)
        return max(0, total - (self.planned_time or 0))

    @property
    def actual_total_time_minutes(self):
        """Actual consumed time for this production entry."""
        return (self.run_time or 0) + (self.downtime_minutes or 0) + (self.make_ready_time or 0)

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


class ProductionWipStatus(models.Model):
    name = models.CharField(max_length=100, unique=True)
    is_active = models.BooleanField(default=True)
    created_by = models.ForeignKey(
        User,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='production_wip_statuses_created',
        editable=False,
    )
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['name']
        verbose_name = 'Production WIP Status'
        verbose_name_plural = 'Production WIP Statuses'

    def __str__(self):
        return self.name


class JobCardWipStatus(models.Model):
    job_card = models.OneToOneField(
        JobCard,
        on_delete=models.CASCADE,
        related_name='production_wip_status',
    )
    status = models.ForeignKey(
        ProductionWipStatus,
        on_delete=models.PROTECT,
        related_name='job_card_wip_statuses',
    )
    updated_by = models.ForeignKey(
        User,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='job_card_wip_status_updates',
        editable=False,
    )
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = 'Job Card WIP Status'
        verbose_name_plural = 'Job Card WIP Statuses'

    def __str__(self):
        return f"{self.job_card.job_card_no} => {self.status.name}"


# ========================= 
# DISPATCH 
# =========================


class Dispatch(models.Model):

    job_card = models.ForeignKey(JobCard, on_delete=models.CASCADE)

    dc_no = models.CharField(
        max_length=50,
        help_text="Dispatch Challan / DR number (required; can be shared across multiple Job Cards and SKUs)",
    )

    dispatch_date = models.DateField(db_index=True)

    dispatch_qty = models.IntegerField(default=0)

    created_by = models.ForeignKey(
        'auth.User',
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='dispatches_created',
        editable=False,
    )

    is_active = models.BooleanField(default=True, db_index=True)


    # =========================
    # VALIDATION ONLY
    # =========================
    def clean(self):
        errors = {}

        if not str(self.dc_no or '').strip():
            errors['dc_no'] = 'DC / DR number is required.'

        dc_no = str(self.dc_no or '').strip()
        if dc_no and self.job_card_id:
            duplicate_qs = Dispatch.objects.filter(
                is_active=True,
                dc_no__iexact=dc_no,
                job_card_id=self.job_card_id,
            )
            if self.pk:
                duplicate_qs = duplicate_qs.exclude(pk=self.pk)
            if duplicate_qs.exists():
                errors['dc_no'] = (
                    f'DC No "{dc_no}" is already used for Job Card {self.job_card.job_card_no}. '
                    'Use a different DC number or edit the existing entry.'
                )

        if self.job_card and self.job_card.workflow_status not in JOB_CARD_DISPATCHABLE_STATUSES:
            errors['job_card'] = (
                'Dispatch can only be created after the Job Card has entered production execution.'
            )

        if self.dispatch_qty <= 0:
            errors['dispatch_qty'] = "Dispatch must be greater than 0"

        existing_dispatch = Dispatch.objects.filter(job_card=self.job_card, is_active=True)\
            .exclude(id=self.id)\
            .aggregate(total=Sum('dispatch_qty'))['total'] or 0

        total_after = existing_dispatch + (self.dispatch_qty or 0)

        if self.job_card.is_print_job:
            total_production = self.job_card.total_packed_pcs
            if total_after > total_production:
                if self.job_card.ups in (None, 0) and self.job_card.packing_productions.exists():
                    errors['dispatch_qty'] = (
                        "Dispatch cannot be validated because no packed quantity is recorded yet."
                    )
                elif self.job_card.ups in (None, 0) and self.job_card.printing_productions.exists():
                    errors['dispatch_qty'] = (
                        "Dispatch cannot be validated because print job UPS is missing. "
                        "Please set UPS on the job card before dispatching."
                    )
                else:
                    errors['dispatch_qty'] = "Dispatch cannot exceed total packed quantity!"
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
        ('plate_request', 'Plate Request'),
    ]

    ACTION_CHOICES = [
        ('create', 'Created'),
        ('update', 'Updated'),
        ('delete', 'Deleted'),
        ('restore', 'Restored'),
        ('submit', 'Submitted'),
        ('approve', 'Approved'),
        ('reject', 'Rejected'),
        ('release', 'Released'),
        ('start_production', 'Start Production'),
        ('complete', 'Completed'),
        ('close', 'Closed'),
        ('send_plate', 'Sent to Vendor'),
        ('receive_plate', 'Received from Vendor'),
        ('mark_available', 'Available for Production'),
        ('archive', 'Archived'),
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
# IN-APP NOTIFICATIONS
# =========================

class Notification(models.Model):
    """Per-user in-app notification (navbar bell + optional toast)."""

    user = models.ForeignKey(User, on_delete=models.CASCADE, related_name='notifications')
    event_type = models.CharField(max_length=80, db_index=True)
    title = models.CharField(max_length=200)
    message = models.TextField(blank=True)
    link = models.CharField(max_length=500, blank=True)
    entity_type = models.CharField(max_length=60, blank=True)
    entity_id = models.PositiveIntegerField(null=True, blank=True)
    is_read = models.BooleanField(default=False, db_index=True)
    created_at = models.DateTimeField(auto_now_add=True, db_index=True)
    created_by = models.ForeignKey(
        User,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='notifications_created',
    )

    class Meta:
        ordering = ['-created_at', '-id']
        indexes = [
            models.Index(fields=['user', 'is_read', '-created_at']),
        ]

    def __str__(self):
        return f'{self.user_id}: {self.title}'


# =========================
# USER ROLES & PERMISSIONS
# =========================

class UserProfile(models.Model):
    """Extended user model with role-based access control"""
    
    ROLE_CHOICES = [
        ('admin', 'Admin — Full system access & configuration'),
        ('manager', 'Manager — Overall oversight (jobs, production, dispatch, reports)'),
        ('planner', 'Planner — Create & manage job cards, view analytics'),
        ('production_manager', 'Production Manager — Final approval and release'),
        ('production', 'Production Supervisor — Manage production entries & team'),
        ('graphics_designer', 'Graphics Designer — Manage plate request workflow'),
        ('operator', 'Machine Operator — Production entry only'),
        ('dispatch', 'Dispatch Coordinator — Dispatch approval & tracking'),
        ('qc', 'QC Inspector — Quality checks & approvals'),
        ('storekeeper', 'Store Keeper — Material & inventory management'),
        ('finance', 'Finance Viewer — Read-only analytics & reports'),
        ('supply_chain', 'Supply Chain — Manage supply chain dashboard & stock'),
    ]
    
    user = models.OneToOneField(User, on_delete=models.CASCADE, related_name='profile')
    role = models.CharField(max_length=20, choices=ROLE_CHOICES, default='operator')
    department = models.ForeignKey(Department, on_delete=models.SET_NULL, null=True, blank=True, related_name='user_profiles')
    manager = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='subordinates_as_manager')
    supervisor = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='subordinates_as_supervisor')
    
    # Custom permissions for granular access control
    can_view_sku_master_review = models.BooleanField(
        default=False,
        help_text="Allow this user to view and manage SKU master review queue"
    )
    can_approve_sku_master = models.BooleanField(
        default=False,
        help_text="Allow this user to approve/reject SKUs in master review"
    )

    created_at = models.DateTimeField(auto_now_add=True)

    def can_view_plate_queue(self):
        return self.normalized_role in ('admin', 'manager', 'planner', 'graphics_designer')

    def can_create_plate_request(self):
        return self.normalized_role in ('admin', 'manager', 'planner', 'graphics_designer')

    def can_send_plate(self):
        return self.normalized_role in ('admin', 'manager', 'planner', 'graphics_designer')

    def can_receive_plate(self):
        return self.normalized_role in ('admin', 'manager', 'planner', 'graphics_designer')

    def can_archive_plate(self):
        return self.normalized_role in ('admin', 'manager', 'planner', 'graphics_designer')
    updated_at = models.DateTimeField(auto_now=True)
    
    class Meta:
        verbose_name = "User Profile"
        verbose_name_plural = "User Profiles"
    
    def __str__(self):
        return f"{self.user.username} — {self.get_role_display()}"
    
    @property
    def normalized_role(self):
        return (self.role or '').strip().lower()
    
    # Permission helpers
    def can_edit_jobcard(self):
        """Can create/edit job cards"""
        return self.normalized_role in ('admin', 'manager', 'planner', 'production_manager')

    def can_approve_planning(self):
        """Can approve the planning queue."""
        return self.normalized_role in ('admin', 'manager', 'planner')
    
    def can_edit_production(self):
        """Can log production data"""
        return self.normalized_role in ('admin', 'manager', 'production_manager', 'production', 'operator')

    def can_start_production(self):
        """Can move a released JobCard into production execution."""
        return self.normalized_role in ('admin', 'manager', 'production_manager', 'production')
    
    def can_approve_dispatch(self):
        """Can approve/edit dispatch"""
        return self.normalized_role in ('admin', 'manager', 'dispatch')
    
    def can_view_analytics(self):
        """Can view dashboard and analytics"""
        return self.normalized_role in ('admin', 'manager', 'planner', 'production', 'dispatch', 'finance')
    
    def can_manage_masters(self):
        """Can manage machines, operators, materials, departments"""
        return self.normalized_role in ('admin', 'manager', 'production_manager')
    
    def can_approve_qc(self):
        """Can perform QC checks"""
        return self.normalized_role in ('admin', 'qc', 'manager')

    def can_approve_pm(self):
        """Can perform production manager approval."""
        return self.normalized_role in ('admin', 'manager', 'production_manager')

    def can_view_planning_queue(self):
        return self.normalized_role in ('admin', 'manager', 'planner', 'production_manager', 'qc')

    def can_view_qc_queue(self):
        return self.normalized_role in ('admin', 'manager', 'qc', 'production_manager')

    def can_view_pm_queue(self):
        return self.normalized_role in ('admin', 'manager', 'production_manager')

    def can_plan(self):
        """Can create, edit, and manage planning jobs (planner role)."""
        return self.normalized_role in ('admin', 'manager', 'planner')

    def can_view_approval_queue(self):
        """Can access the approval queue page (view-only or with actions)."""
        return self.normalized_role in ('admin', 'manager', 'planner', 'qc', 'production_manager')
    
    def can_view_sku_master_review_queue(self):
        """Can view SKU master review queue - role-based or custom permission."""
        if self.can_view_sku_master_review:
            return True
        return self.normalized_role in ('admin', 'qc', 'manager', 'planner')
    
    def can_approve_sku_master_review(self):
        """Can approve/reject SKUs in master review - role-based or custom permission."""
        if self.can_approve_sku_master:
            return True
        return self.normalized_role in ('admin', 'qc', 'manager')
    
    def can_manage_operators(self):
        """Can assign operators to shifts/jobs"""
        return self.normalized_role in ('admin', 'manager', 'production_manager', 'production')

    def can_archive_records(self):
        """Can archive and restore operational records"""
        return self.normalized_role in ('admin', 'manager', 'production_manager')
    
    def can_view_job_summary(self):
        """Can view the Jobs Summary dashboard."""
        return self.normalized_role in (
            'admin', 'manager', 'planner', 'production_manager',
            'production', 'qc', 'dispatch'
        )

    def can_view_reports(self):
        """Can view financial/operational reports"""
        return self.normalized_role in ('admin', 'manager', 'finance')


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


# =========================
# RULE-BASED NOTIFICATIONS
# =========================

class NotificationEvent(models.Model):
    code = models.CharField(max_length=80, unique=True, db_index=True)
    name = models.CharField(max_length=200)
    description = models.TextField(blank=True)
    module = models.CharField(max_length=100, blank=True)
    is_active = models.BooleanField(default=True)
    title_template = models.CharField(max_length=200, blank=True)
    message_template = models.TextField(blank=True)
    link_template = models.CharField(max_length=500, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.name} ({self.code})"


class NotificationRule(models.Model):
    PRIORITY_CHOICES = [
        ('low', 'Low'),
        ('medium', 'Medium'),
        ('high', 'High'),
    ]
    RECIPIENT_TYPE_CHOICES = [
        ('role', 'Role'),
        ('user', 'Specific User'),
        ('department', 'Department'),
        ('creator', 'Creator'),
        ('manager', 'Manager'),
        ('supervisor', 'Supervisor'),
        ('next_stage', 'Next Workflow Stage'),
    ]

    event = models.ForeignKey(NotificationEvent, on_delete=models.CASCADE, related_name='rules')
    enabled = models.BooleanField(default=True)
    recipient_type = models.CharField(
        max_length=50,
        choices=RECIPIENT_TYPE_CHOICES,
        default='role'
    )
    role = models.CharField(
        max_length=50,
        blank=True,
        null=True,
        choices=UserProfile.ROLE_CHOICES
    )
    user = models.ForeignKey(
        User,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='notification_rules'
    )
    department = models.ForeignKey(
        Department,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='notification_rules'
    )
    send_to_creator = models.BooleanField(default=False)
    send_to_manager = models.BooleanField(default=False)
    send_to_supervisor = models.BooleanField(default=False)
    send_to_next_stage = models.BooleanField(default=False)
    exclude_actor = models.BooleanField(default=True)
    priority = models.CharField(
        max_length=20,
        choices=PRIORITY_CHOICES,
        default='medium'
    )
    email_enabled = models.BooleanField(default=False)
    sms_enabled = models.BooleanField(default=False)
    in_app_enabled = models.BooleanField(default=True)
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"Rule for {self.event.code} ({self.get_recipient_type_display()})"


class WorkflowTransition(models.Model):
    module = models.CharField(max_length=100)
    current_stage = models.CharField(max_length=100)
    action = models.CharField(max_length=100)
    next_stage = models.CharField(max_length=100)
    notify_role = models.CharField(max_length=50, blank=True)

    class Meta:
        unique_together = ('module', 'current_stage', 'action', 'next_stage')

    def __str__(self):
        return f"{self.module}: {self.current_stage} -> {self.action} -> {self.next_stage} (Notify: {self.notify_role})"


class NotificationRuleAuditLog(models.Model):
    rule = models.ForeignKey(
        NotificationRule,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='audit_logs'
    )
    changed_by = models.ForeignKey(
        User,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='notification_rule_audits'
    )
    action = models.CharField(max_length=50)  # 'create', 'update', 'delete'
    old_values = models.JSONField(default=dict, blank=True)
    new_values = models.JSONField(default=dict, blank=True)
    timestamp = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['-timestamp']

    def __str__(self):
        return f"Rule {self.rule_id or 'Deleted'} {self.action} by {self.changed_by} at {self.timestamp}"


class PasswordResetRequest(models.Model):
    username_or_email = models.CharField(max_length=254)
    created_at = models.DateTimeField(auto_now_add=True)
    resolved = models.BooleanField(default=False)
    resolved_at = models.DateTimeField(null=True, blank=True)

    class Meta:
        ordering = ['-created_at']

    def __str__(self):
        return f"Reset request for {self.username_or_email} at {self.created_at}"