from decimal import Decimal

from django.conf import settings
from django.core.exceptions import ValidationError
from django.db import models

from .calculations import compute_print_plan

# Vocabulary mirrors planning.PLANNING_STATUS_CHOICES but is defined here so the
# Flexo module has no import dependency on the offset `planning` app.
FLEXO_STATUS_CHOICES = [
    ('draft', 'Draft'),
    ('pending_qc', 'Pending QC'),
    ('qc_approved', 'QC Approved'),
    ('released', 'Released'),
    ('in_production', 'In Production'),
    ('completed', 'Completed'),
    ('cancelled', 'Cancelled'),
]

# Jobs at or before these statuses can be cancelled directly by a planner.
# Later ones need Production Manager approval via FlexoJobChangeRequest.
PRE_RELEASE_STATUSES = {'draft', 'pending_qc', 'qc_approved'}
LOCKED_STATUSES = {'qc_approved', 'released', 'in_production', 'completed'}

CANCEL_REASON_CHOICES = [
    ('customer_cancelled', 'Customer Cancelled the SKU'),
    ('po_amended', 'PO Amended / Quantity Removed'),
    ('duplicate_entry', 'Duplicate Planning Entry'),
    ('superseded', 'Superseded by Another Job'),
    ('other', 'Other'),
]
JOB_CANCEL_REQUEST_TYPE = 'cancel_job'

PRIORITY_CHOICES = [
    (0, 'Low'),
    (1, 'Normal'),
    (2, 'Medium'),
    (3, 'High'),
    (4, 'Urgent'),
]

PURCHASE_MATERIAL_ORIGIN_CHOICES = [
    ('', 'Select Origin'),
    ('local', 'Local'),
    ('import', 'Import'),
]

REPEAT_FLAG_CHOICES = [
    ('', 'Select'),
    ('New', 'New'),
    ('Repeat', 'Repeat'),
]

SHIFT_CHOICES = [
    ('', 'Select'),
    ('A', 'A'),
    ('B', 'B'),
]

UNIT_CHOICES = [
    ('piece', 'Piece'),
    ('roll', 'Roll'),
]
UNIT_LABELS = {'piece': 'pcs', 'roll': 'rolls'}

MASTER_STATUS_CHOICES = [
    ('draft', 'Draft'),
    ('pending_review', 'Pending Review'),
    ('reviewed', 'Reviewed'),
    ('approved', 'Approved'),
]

_USER = settings.AUTH_USER_MODEL


def _user_fk():
    return models.ForeignKey(_USER, on_delete=models.SET_NULL, null=True, blank=True, related_name='+')


def _dec(max_digits=14, places=3, **kw):
    return models.DecimalField(max_digits=max_digits, decimal_places=places, null=True, blank=True, **kw)


class FlexoSkuRecipe(models.Model):
    """Flexo SKU master — mirrors planning.SkuRecipe's approval workflow with
    roll/web specs instead of sheet/plate specs."""

    sku = models.CharField(max_length=255, unique=True)
    job_name = models.CharField(max_length=255, blank=True)
    material = models.CharField(max_length=100, blank=True)
    color_spec = models.CharField(max_length=150, blank=True)

    size_w_mm = _dec(10, 3)
    size_h_mm = _dec(10, 3)
    ups = _dec(6, 2)
    roll_size = _dec(10, 3)          # web width mm ("Print Roll Size")
    meter_per_roll = _dec(10, 3, help_text='Length of one finished roll in meters.')
    pcs_per_roll = models.PositiveIntegerField(
        null=True, blank=True,
        help_text='Finished pcs per roll (e.g. 5000). Blank = derived from piece length (no label gap).')
    wastage_pct = models.DecimalField(max_digits=5, decimal_places=2, default=Decimal('5'))
    unit_type = models.CharField(
        'Order unit', max_length=10, choices=UNIT_CHOICES, default='piece',
        help_text='Unit the customer orders and receives in: Roll (label POs) or Piece (ribbon WOs).')

    die_code = models.CharField('Die Code / AWC', max_length=80, blank=True)
    teeth = models.CharField(max_length=40, blank=True)

    machine = models.ForeignKey('core.Machine', on_delete=models.SET_NULL, null=True, blank=True, related_name='+')
    department = models.ForeignKey('core.Department', on_delete=models.SET_NULL, null=True, blank=True, related_name='+')
    purchase_material_origin = models.CharField(max_length=10, choices=PURCHASE_MATERIAL_ORIGIN_CHOICES, blank=True)
    default_unit_cost = _dec(12, 4)
    daily_demand = _dec(10, 3)
    remarks = models.CharField(max_length=255, blank=True)

    master_data_status = models.CharField(max_length=20, choices=MASTER_STATUS_CHOICES, default='draft', db_index=True)
    reviewed_by = _user_fk()
    reviewed_at = models.DateTimeField(null=True, blank=True)
    approved_by = _user_fk()
    approved_at = models.DateTimeField(null=True, blank=True)
    rejection_comment = models.TextField(blank=True)
    last_rejected_by = _user_fk()
    last_rejected_at = models.DateTimeField(null=True, blank=True)

    is_active = models.BooleanField(default=True, db_index=True)
    archived_by = _user_fk()
    archived_at = models.DateTimeField(null=True, blank=True)
    archive_reason = models.TextField(blank=True)
    created_by = _user_fk()
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['sku']
        verbose_name = 'Flexo SKU Recipe'

    def __str__(self):
        return self.sku

    @property
    def is_approved(self):
        return self.master_data_status == 'approved'

    def missing_required_fields(self):
        required = [
            ('job_name', 'Job Name'), ('material', 'Material'), ('size_w_mm', 'Size W (mm)'),
            ('size_h_mm', 'Size H (mm)'), ('ups', 'Ups'), ('roll_size', 'Roll Size'),
            ('meter_per_roll', 'Meter Per Roll'), ('machine', 'Machine'),
        ]
        missing = []
        for field, label in required:
            value = getattr(self, field, None)
            if value is None or value == '':
                missing.append(label)
        return missing

    def warning_fields(self):
        """Optional until die work is done (like Plate Set No. in offset)."""
        return [] if str(self.die_code or '').strip() else ['Die Code / AWC']


class FlexoPlanningJob(models.Model):
    """The Flexo planning job (one row per SKU/order). Order/stock/dispatch are in the order unit (pcs or rolls); print qty is derived in METERS."""

    PRIORITY_CHOICES = PRIORITY_CHOICES

    sku_recipe = models.ForeignKey(FlexoSkuRecipe, on_delete=models.SET_NULL, null=True, blank=True, related_name='jobs')

    jc_number = models.CharField(max_length=50, unique=True)
    po_number = models.CharField(max_length=80, blank=True)
    po_received_date = models.DateField(null=True, blank=True)
    plan_date = models.DateField(null=True, blank=True)
    plan_month = models.CharField(max_length=20, blank=True)
    delivery_date = models.DateField(null=True, blank=True)

    sku = models.CharField(max_length=255, blank=True)
    job_name = models.CharField(max_length=255, blank=True)
    repeat_flag = models.CharField(max_length=10, choices=REPEAT_FLAG_CHOICES, blank=True)
    material = models.CharField(max_length=100, blank=True)
    color_spec = models.CharField(max_length=150, blank=True)

    size_w_mm = _dec(10, 3)
    size_h_mm = _dec(10, 3)
    ups = _dec(6, 2)
    roll_size = _dec(10, 3)
    meter_per_roll = _dec(10, 3)
    pcs_per_roll = models.PositiveIntegerField(null=True, blank=True)
    wastage_pct = models.DecimalField(max_digits=5, decimal_places=2, default=Decimal('5'))
    unit_type = models.CharField('Order unit', max_length=10, choices=UNIT_CHOICES, default='piece')

    die_code = models.CharField('Die Code / AWC', max_length=80, blank=True)
    teeth = models.CharField(max_length=40, blank=True)

    # Quantities: order/stock/dispatch in the order unit (pcs or rolls), print in METERS.
    order_qty = models.PositiveIntegerField(default=0)
    stock_qty = models.PositiveIntegerField(default=0)
    net_print_qty = models.PositiveIntegerField(default=0)
    order_roll_qty = _dec(14, 3)      # finished rolls (= net qty when ordered in rolls)
    base_meters = _dec(14, 3)
    wastage_meters = _dec(14, 3)
    print_meters = _dec(14, 3)        # "Actual Roll Qty Meter" = base + wastage

    roll_qty_received = _dec(14, 3)
    actual_wastage_mtr = _dec(14, 3)
    received_qty_mtr = _dec(14, 3)
    remarks = models.CharField(max_length=255, blank=True)

    status = models.CharField(max_length=20, choices=FLEXO_STATUS_CHOICES, default='draft', db_index=True)
    priority = models.IntegerField(choices=PRIORITY_CHOICES, default=1, db_index=True)
    issued_to_production = models.BooleanField(default=False)
    rejection_reason = models.TextField(blank=True)

    rejected_qty = models.PositiveIntegerField(default=0)
    balance_qty = models.IntegerField(default=0)
    destination = models.CharField(max_length=100, blank=True)
    unit_cost = _dec(12, 4)
    amount = _dec(14, 2)

    machine_name = models.ForeignKey('core.Machine', on_delete=models.SET_NULL, null=True, blank=True, related_name='+')
    department = models.ForeignKey('core.Department', on_delete=models.SET_NULL, null=True, blank=True, related_name='+')
    purchase_material_origin = models.CharField(max_length=10, choices=PURCHASE_MATERIAL_ORIGIN_CHOICES, blank=True)
    daily_demand = _dec(10, 3)

    is_active = models.BooleanField(default=True, db_index=True)
    is_on_hold = models.BooleanField(default=False)
    hold_reason = models.TextField(blank=True)
    hold_by = _user_fk()
    hold_at = models.DateTimeField(null=True, blank=True)

    archived_by = _user_fk()
    archived_at = models.DateTimeField(null=True, blank=True)
    archive_reason = models.TextField(blank=True)
    restored_by = _user_fk()
    restored_at = models.DateTimeField(null=True, blank=True)
    restore_reason = models.TextField(blank=True)

    is_cancelled = models.BooleanField(default=False, db_index=True)
    cancel_reason_code = models.CharField(max_length=40, choices=CANCEL_REASON_CHOICES, blank=True)
    cancel_reason = models.TextField(blank=True)
    cancelled_by = _user_fk()
    cancelled_at = models.DateTimeField(null=True, blank=True)

    job_card_version = models.PositiveIntegerField(default=1)
    created_by = _user_fk()
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    last_edited_by = _user_fk()
    last_edited_at = models.DateTimeField(null=True, blank=True)

    class Meta:
        ordering = ['-created_at', '-id']
        verbose_name = 'Flexo Planning Job'

    def __str__(self):
        return f"{self.jc_number} | {self.sku}" if self.sku else self.jc_number

    # ---- helpers -----------------------------------------------------
    @property
    def unit_label(self):
        return UNIT_LABELS.get(self.unit_type, 'pcs')

    @property
    def effective_status(self):
        return 'cancelled' if self.is_cancelled else self.status

    @property
    def is_locked(self):
        return self.effective_status in LOCKED_STATUSES

    @property
    def total_dispatched(self):
        total = self.dispatch_runs.aggregate(total=models.Sum('delivered_qty'))['total']
        return total or 0

    @property
    def total_printed_meters(self):
        total = self.print_runs.aggregate(total=models.Sum('print_meters'))['total']
        return total or Decimal('0')

    @property
    def can_cancel_directly(self):
        return self.effective_status in PRE_RELEASE_STATUSES

    def cancellation_blockers(self):
        blockers = []
        if self.is_cancelled:
            return ['This job is already cancelled.']
        card = getattr(self, 'job_card', None)
        if card and card.status in {'in_production', 'completed'}:
            blockers.append(
                f'Job Card {card.job_card_no} is already {card.status.replace("_", " ")}; it cannot be cancelled.'
            )
        if self.print_runs.exists():
            blockers.append('Production has already been booked against this job.')
        return blockers

    def recalculate(self):
        plan = compute_print_plan(
            order_qty=self.order_qty, stock_qty=self.stock_qty, ups=self.ups,
            meter_per_roll=self.meter_per_roll, pcs_per_roll=self.pcs_per_roll,
            size_h_mm=self.size_h_mm, wastage_pct=self.wastage_pct, order_unit=self.unit_type,
        )
        self.net_print_qty = plan.net_qty
        self.order_roll_qty = plan.order_rolls.quantize(Decimal('0.001'))
        self.base_meters = plan.base_meters.quantize(Decimal('0.001'))
        self.wastage_meters = plan.wastage_meters.quantize(Decimal('0.001'))
        self.print_meters = plan.print_meters.quantize(Decimal('0.001'))
        return plan

    def clean(self):
        if self.stock_qty and self.order_qty and self.stock_qty > self.order_qty:
            raise ValidationError({'stock_qty': 'Stock qty cannot exceed the order qty.'})

    def save(self, *args, **kwargs):
        update_fields = kwargs.get('update_fields')
        self.recalculate()
        if self.pk:
            self.balance_qty = (self.order_qty or 0) - self.total_dispatched
        else:
            self.balance_qty = self.order_qty or 0
        if update_fields is not None:
            kwargs['update_fields'] = set(update_fields) | {
                'net_print_qty', 'order_roll_qty', 'base_meters', 'wastage_meters',
                'print_meters', 'balance_qty',
            }
        super().save(*args, **kwargs)


class FlexoPlanningPrintRun(models.Model):
    """Meters printed per run against a planning job (mirrors PlanningPrintRun)."""

    planning_job = models.ForeignKey(FlexoPlanningJob, on_delete=models.CASCADE, related_name='print_runs')
    run_index = models.PositiveSmallIntegerField()
    print_date = models.DateField(null=True, blank=True)
    print_meters = _dec(14, 3)
    wastage_meters = _dec(14, 3)

    class Meta:
        unique_together = ('planning_job', 'run_index')
        ordering = ['run_index']


class FlexoPlanningDispatchRun(models.Model):
    """One partial delivery in PCS (mirrors planning.PlanningDispatchRun)."""

    flexo_planning_job = models.ForeignKey(FlexoPlanningJob, on_delete=models.CASCADE, related_name='dispatch_runs')
    dispatch_index = models.PositiveSmallIntegerField()
    delivery_date = models.DateField(null=True, blank=True)
    dc_no = models.CharField(max_length=80, blank=True)
    delivered_qty = models.PositiveIntegerField(null=True, blank=True)
    packing = models.CharField(max_length=80, blank=True)
    site = models.CharField(max_length=80, blank=True)

    class Meta:
        unique_together = ('flexo_planning_job', 'dispatch_index')
        ordering = ['dispatch_index']

    def __str__(self):
        return f"{self.flexo_planning_job.jc_number} #{self.dispatch_index}"


class FlexoJobChangeRequest(models.Model):
    """PM-approved reopen/cancel request on a locked job (mirrors JobCardChangeRequest)."""

    STATUS_CHOICES = [
        ('pending', 'Pending PM Approval'),
        ('approved', 'Approved'),
        ('rejected', 'Rejected'),
    ]

    planning_job = models.ForeignKey(FlexoPlanningJob, on_delete=models.CASCADE, related_name='change_requests')
    request_type = models.CharField(max_length=50, default='reopen_to_draft')
    reason = models.TextField()
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='pending')
    requested_by = _user_fk()
    requested_at = models.DateTimeField(auto_now_add=True)
    approved_by = _user_fk()
    approved_at = models.DateTimeField(null=True, blank=True)
    rejection_reason = models.TextField(blank=True)

    class Meta:
        ordering = ['-requested_at']

    def __str__(self):
        return f"{self.request_type} {self.planning_job.jc_number} ({self.status})"

    @property
    def is_cancellation(self):
        return self.request_type == JOB_CANCEL_REQUEST_TYPE

    def _split_cancel_reason(self):
        import re
        match = re.match(r'^\[([a-z_]+)\]\s*(.*)$', (self.reason or '').strip(), re.DOTALL)
        if not match:
            return '', (self.reason or '').strip()
        return match.group(1), match.group(2).strip()

    @property
    def cancel_reason_code(self):
        return self._split_cancel_reason()[0]

    @property
    def cancel_reason_text(self):
        return self._split_cancel_reason()[1]


class FlexoJobEvent(models.Model):
    """Lightweight audit trail for workflow transitions on a job."""

    planning_job = models.ForeignKey(FlexoPlanningJob, on_delete=models.CASCADE, related_name='events')
    action = models.CharField(max_length=40)
    from_status = models.CharField(max_length=20, blank=True)
    to_status = models.CharField(max_length=20, blank=True)
    reason = models.TextField(blank=True)
    actor = _user_fk()
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['-created_at', '-id']


class FlexoJobCard(models.Model):
    """Printable Flexo PO Job Card — one per planning job (created when the job
    is sent to QC), mirrors the concept of core.JobCard."""

    planning_job = models.OneToOneField(FlexoPlanningJob, on_delete=models.CASCADE, related_name='job_card')
    job_card_no = models.CharField(max_length=50, unique=True)

    sku = models.CharField(max_length=255, blank=True)
    job_name = models.CharField(max_length=255, blank=True)
    po_number = models.CharField(max_length=80, blank=True)
    po_received_date = models.DateField(null=True, blank=True)
    order_qty = models.PositiveIntegerField(default=0)
    unit_type = models.CharField(max_length=10, choices=UNIT_CHOICES, default='piece')
    delivery_location = models.CharField(max_length=100, blank=True)
    department = models.ForeignKey('core.Department', on_delete=models.SET_NULL, null=True, blank=True, related_name='+')
    days_plan = models.PositiveIntegerField(null=True, blank=True)
    daily_demand = _dec(10, 3)

    material_type = models.CharField(max_length=100, blank=True)
    size_w_mm = _dec(10, 3)
    size_h_mm = _dec(10, 3)
    print_roll_size = _dec(10, 3)
    roll_meter_qty = _dec(14, 3)
    ups = _dec(6, 2)
    machine = models.ForeignKey('core.Machine', on_delete=models.SET_NULL, null=True, blank=True, related_name='+')
    wastage = _dec(14, 3)
    color = models.CharField(max_length=150, blank=True)
    special_instructions = models.TextField(blank=True)
    mtr_per_roll = models.CharField(max_length=40, blank=True)
    die_code = models.CharField('Die Code / AWC', max_length=80, blank=True)

    prepared_by = models.CharField(max_length=100, blank=True)
    approved_by = models.CharField(max_length=100, blank=True)

    stock_check = models.CharField(max_length=100, blank=True)
    dispatched_by = models.CharField(max_length=100, blank=True)
    extra_production = models.CharField(max_length=100, blank=True)
    actual_waste = models.CharField(max_length=100, blank=True)
    inventory_controller = models.CharField(max_length=100, blank=True)
    receiving_signature = models.CharField(max_length=100, blank=True)

    status = models.CharField(max_length=20, choices=FLEXO_STATUS_CHOICES, default='draft', db_index=True)
    is_active = models.BooleanField(default=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['-created_at']
        verbose_name = 'Flexo Job Card'

    def __str__(self):
        return self.job_card_no

    @property
    def unit_label(self):
        return UNIT_LABELS.get(self.unit_type, 'pcs')


class FlexoPrintingLogEntry(models.Model):
    """PRINTING row on the Merge Job Card — all in METERS."""

    job_card = models.ForeignKey(FlexoJobCard, on_delete=models.CASCADE, related_name='printing_log_entries')
    date = models.DateField(null=True, blank=True)
    machine = models.ForeignKey('core.Machine', on_delete=models.SET_NULL, null=True, blank=True, related_name='+')
    operator = models.ForeignKey('core.Operator', on_delete=models.SET_NULL, null=True, blank=True, related_name='+')
    shift = models.CharField(max_length=1, choices=SHIFT_CHOICES, blank=True)
    given_meter_qty = _dec(14, 3)
    wastage_meters = _dec(14, 3)
    finished_meter_qty = _dec(14, 3)
    signature_name = models.CharField(max_length=100, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['date', 'id']

    def __str__(self):
        return f"{self.job_card.job_card_no} printing {self.date or ''}"


class FlexoSlittingCuttingLogEntry(models.Model):
    """SLITTING / CUTTING row on the Merge Job Card: meters in, PCS out."""

    job_card = models.ForeignKey(FlexoJobCard, on_delete=models.CASCADE, related_name='slitting_cutting_log_entries')
    date = models.DateField(null=True, blank=True)
    machine = models.ForeignKey('core.Machine', on_delete=models.SET_NULL, null=True, blank=True, related_name='+')
    operator = models.ForeignKey('core.Operator', on_delete=models.SET_NULL, null=True, blank=True, related_name='+')
    shift = models.CharField(max_length=1, choices=SHIFT_CHOICES, blank=True)
    given_meter_qty = _dec(14, 3)
    wastage_pcs = models.PositiveIntegerField(null=True, blank=True)
    finished_pcs = models.PositiveIntegerField(null=True, blank=True)
    signature_name = models.CharField(max_length=100, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['date', 'id']
        verbose_name = 'Flexo Slitting/Cutting Log Entry'

    def __str__(self):
        return f"{self.job_card.job_card_no} slitting/cutting {self.date or ''}"


class FlexoPoDocument(models.Model):
    """An uploaded PO/WO whose lines were extracted for Flexo planning. Only the
    extracted payload is kept (the PDF itself is not stored)."""

    STATUS_CHOICES = [('pending', 'Pending'), ('processed', 'Processed')]

    po_number = models.CharField(max_length=80, unique=True)
    document_type = models.CharField(max_length=10, blank=True)
    payload = models.JSONField(default=dict)
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='pending', db_index=True)
    uploaded_by = _user_fk()
    uploaded_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['-uploaded_at']
        verbose_name = 'Flexo PO/WO Document'

    def __str__(self):
        return self.po_number

    @property
    def items(self):
        return (self.payload or {}).get('items') or []
