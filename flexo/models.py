from django.conf import settings
from django.db import models


# Mirrors planning.PLANNING_STATUS_CHOICES's vocabulary so the two modules stay
# consistent, but is defined independently here (not imported from `planning`)
# so nothing in the offset app is touched by this module.
FLEXO_STATUS_CHOICES = [
    ('draft', 'Draft'),
    ('pending_qc', 'Pending QC'),
    ('qc_approved', 'QC Approved'),
    ('released', 'Released'),
    ('in_production', 'In Production'),
    ('completed', 'Completed'),
    ('cancelled', 'Cancelled'),
]

# Mirrors planning.PURCHASE_MATERIAL_ORIGIN_CHOICES verbatim (same field, same
# values in the source spreadsheet) — redefined here rather than imported so
# this app has no hard dependency on `planning`.
PURCHASE_MATERIAL_ORIGIN_CHOICES = [
    ('', 'Select Origin'),
    ('local', 'Local'),
    ('import', 'Import'),
]

MATERIAL_FAMILY_CHOICES = [
    ('fl', 'FL (Flexo Label)'),
    ('satin', 'SATIN (Satin Ribbon)'),
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


class FlexoPlanningJob(models.Model):
    """The Flexo "Planning Sheet" entry — one row per SKU/order, covering both
    FL (label) and SATIN (ribbon) jobs via `material_family`. Mirrors the
    concept (not the code) of planning.PlanningJob for the offset module."""

    material_family = models.CharField(max_length=10, choices=MATERIAL_FAMILY_CHOICES)

    jc_number = models.CharField(max_length=50, unique=True)
    po_number = models.CharField(max_length=80, blank=True)
    po_received_date = models.DateField(null=True, blank=True)
    plan_month = models.CharField(max_length=20, blank=True)

    sku = models.CharField(max_length=255, blank=True)
    job_name = models.CharField(max_length=255, blank=True)
    repeat_flag = models.CharField(max_length=10, choices=REPEAT_FLAG_CHOICES, blank=True)

    material = models.CharField(max_length=100, blank=True)
    color_spec = models.CharField(max_length=150, blank=True)

    size_w_mm = models.DecimalField(max_digits=10, decimal_places=3, null=True, blank=True)
    # FL-only
    size_h_mm = models.DecimalField(max_digits=10, decimal_places=3, null=True, blank=True)
    die_code = models.CharField(max_length=80, blank=True)
    # SATIN-only
    size_h_inches = models.DecimalField(max_digits=10, decimal_places=3, null=True, blank=True)
    awc_no = models.CharField(max_length=80, blank=True)
    teeth = models.CharField(max_length=40, blank=True)
    order_roll_qty = models.DecimalField(max_digits=14, decimal_places=3, null=True, blank=True)
    printing_cylinder = models.CharField(max_length=40, blank=True)

    order_qty = models.DecimalField(max_digits=14, decimal_places=3, null=True, blank=True)
    ups = models.DecimalField(max_digits=6, decimal_places=2, null=True, blank=True)
    roll_size = models.DecimalField(max_digits=10, decimal_places=3, null=True, blank=True)
    meter_per_roll = models.DecimalField(max_digits=10, decimal_places=3, null=True, blank=True)
    roll_qty_meter = models.DecimalField(max_digits=14, decimal_places=3, null=True, blank=True)
    wastage_qty_meter = models.DecimalField(max_digits=14, decimal_places=3, null=True, blank=True)
    actual_roll_qty_meter = models.DecimalField(max_digits=14, decimal_places=3, null=True, blank=True)
    remarks = models.CharField(max_length=255, blank=True)

    roll_qty_received = models.DecimalField(max_digits=14, decimal_places=3, null=True, blank=True)
    actual_wastage_mtr = models.DecimalField(max_digits=14, decimal_places=3, null=True, blank=True)
    received_qty_mtr = models.DecimalField(max_digits=14, decimal_places=3, null=True, blank=True)

    status = models.CharField(max_length=20, choices=FLEXO_STATUS_CHOICES, default='draft', blank=True, db_index=True)

    rejected_qty = models.DecimalField(max_digits=14, decimal_places=3, default=0, blank=True)
    balance_qty = models.DecimalField(max_digits=14, decimal_places=3, default=0, blank=True)
    destination = models.CharField(max_length=100, blank=True)
    cost = models.DecimalField(max_digits=12, decimal_places=4, null=True, blank=True)
    amount = models.DecimalField(max_digits=14, decimal_places=2, null=True, blank=True)

    machine_name = models.ForeignKey('core.Machine', on_delete=models.SET_NULL, null=True, blank=True, related_name='+')
    purchase_material_origin = models.CharField(max_length=10, choices=PURCHASE_MATERIAL_ORIGIN_CHOICES, blank=True)
    stock_qty = models.DecimalField(max_digits=14, decimal_places=3, null=True, blank=True)
    daily_demand = models.DecimalField(max_digits=10, decimal_places=3, null=True, blank=True)
    department = models.ForeignKey('core.Department', on_delete=models.SET_NULL, null=True, blank=True, related_name='+')

    is_active = models.BooleanField(default=True)
    created_by = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, blank=True, related_name='+')
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['-created_at']
        verbose_name = 'Flexo Planning Job'

    def __str__(self):
        return self.jc_number

    @property
    def total_dispatched(self):
        total = self.dispatch_runs.aggregate(total=models.Sum('delivered_qty'))['total']
        return total or 0

    @property
    def is_fl(self):
        return self.material_family == 'fl'

    @property
    def is_satin(self):
        return self.material_family == 'satin'


class FlexoPlanningDispatchRun(models.Model):
    """Mirrors planning.PlanningDispatchRun field-for-field — one row per
    partial delivery against a FlexoPlanningJob, replacing the spreadsheet's
    fixed 6 "Date Delivery 0N / DC 0N / Delivered Quantity 0N" column groups
    with an unbounded related table."""

    flexo_planning_job = models.ForeignKey(FlexoPlanningJob, on_delete=models.CASCADE, related_name='dispatch_runs')
    dispatch_index = models.PositiveSmallIntegerField()
    delivery_date = models.DateField(null=True, blank=True)
    dc_no = models.CharField(max_length=80, blank=True)
    delivered_qty = models.DecimalField(max_digits=14, decimal_places=3, null=True, blank=True)

    class Meta:
        unique_together = ('flexo_planning_job', 'dispatch_index')
        ordering = ['dispatch_index']

    def __str__(self):
        return f"{self.flexo_planning_job.jc_number} #{self.dispatch_index}"


class FlexoJobCard(models.Model):
    """The printable Flexo PO Job Card — one per FlexoPlanningJob, used on the
    shop floor. Mirrors the concept of core.JobCard for the offset module."""

    planning_job = models.OneToOneField(FlexoPlanningJob, on_delete=models.CASCADE, related_name='job_card')
    job_card_no = models.CharField(max_length=50, unique=True)
    material_family = models.CharField(max_length=10, choices=MATERIAL_FAMILY_CHOICES)

    # Order Information
    sku = models.CharField(max_length=255, blank=True)
    job_name = models.CharField(max_length=255, blank=True)
    po_number = models.CharField(max_length=80, blank=True)
    po_received_date = models.DateField(null=True, blank=True)
    order_qty = models.DecimalField(max_digits=14, decimal_places=3, null=True, blank=True)
    delivery_location = models.CharField(max_length=100, blank=True)
    department = models.ForeignKey('core.Department', on_delete=models.SET_NULL, null=True, blank=True, related_name='+')
    days_plan = models.PositiveIntegerField(null=True, blank=True)
    daily_demand = models.DecimalField(max_digits=10, decimal_places=3, null=True, blank=True)

    # Material and Work Process
    material_type = models.CharField(max_length=100, blank=True)
    size_w_mm = models.DecimalField(max_digits=10, decimal_places=3, null=True, blank=True)
    size_h_mm = models.DecimalField(max_digits=10, decimal_places=3, null=True, blank=True)
    size_h_inches = models.DecimalField(max_digits=10, decimal_places=3, null=True, blank=True)
    print_roll_size = models.DecimalField(max_digits=10, decimal_places=3, null=True, blank=True)
    roll_meter_qty = models.DecimalField(max_digits=14, decimal_places=3, null=True, blank=True)
    ups = models.DecimalField(max_digits=6, decimal_places=2, null=True, blank=True)
    machine = models.ForeignKey('core.Machine', on_delete=models.SET_NULL, null=True, blank=True, related_name='+')
    wastage = models.DecimalField(max_digits=14, decimal_places=3, null=True, blank=True)
    color = models.CharField(max_length=150, blank=True)
    special_instructions = models.TextField(blank=True)
    mtr_per_roll = models.CharField(max_length=40, blank=True)

    # FL-only
    die_code = models.CharField(max_length=80, blank=True)
    # SATIN-only
    awc_no = models.CharField(max_length=80, blank=True)
    printing_cylinder = models.CharField(max_length=40, blank=True)

    # Sign-off (FL uses prepared_by/approved_by; SATIN uses all four)
    prepared_by = models.CharField(max_length=100, blank=True)
    checked_by = models.CharField(max_length=100, blank=True)
    artwork_developed_by = models.CharField(max_length=100, blank=True)
    approved_by = models.CharField(max_length=100, blank=True)

    status = models.CharField(max_length=20, choices=FLEXO_STATUS_CHOICES, default='draft', blank=True, db_index=True)

    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['-created_at']
        verbose_name = 'Flexo Job Card'

    def __str__(self):
        return self.job_card_no

    @property
    def is_satin(self):
        return self.material_family == 'satin'


class FlexoPrintingLogEntry(models.Model):
    """One row in the Job Card's Printing log grid. FL logs Jumbo/Baby Roll
    Qty; SATIN logs Given/Finished Meter Qty — both columns exist here and the
    template only shows the pair relevant to the card's material_family."""

    job_card = models.ForeignKey(FlexoJobCard, on_delete=models.CASCADE, related_name='printing_log_entries')
    date = models.DateField(null=True, blank=True)
    machine = models.ForeignKey('core.Machine', on_delete=models.SET_NULL, null=True, blank=True, related_name='+')
    operator = models.ForeignKey('core.Operator', on_delete=models.SET_NULL, null=True, blank=True, related_name='+')
    shift = models.CharField(max_length=1, choices=SHIFT_CHOICES, blank=True)

    # FL
    jumbo_roll_qty = models.DecimalField(max_digits=14, decimal_places=3, null=True, blank=True)
    baby_roll_qty = models.DecimalField(max_digits=14, decimal_places=3, null=True, blank=True)
    # SATIN
    given_meter_qty = models.DecimalField(max_digits=14, decimal_places=3, null=True, blank=True)
    finished_meter_qty = models.DecimalField(max_digits=14, decimal_places=3, null=True, blank=True)

    wastage = models.DecimalField(max_digits=14, decimal_places=3, null=True, blank=True)
    signature_name = models.CharField(max_length=100, blank=True)

    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['date', 'id']

    def __str__(self):
        return f"{self.job_card.job_card_no} printing {self.date or ''}"


class FlexoSlittingCuttingLogEntry(models.Model):
    """SATIN-only Slitting/Cutting log — the ribbon-converting step (slit to
    width / cut to length) that runs after printing. FL job cards have no
    equivalent stage."""

    job_card = models.ForeignKey(FlexoJobCard, on_delete=models.CASCADE, related_name='slitting_cutting_log_entries')
    date = models.DateField(null=True, blank=True)
    operator = models.ForeignKey('core.Operator', on_delete=models.SET_NULL, null=True, blank=True, related_name='+')
    cut_qty = models.DecimalField(max_digits=14, decimal_places=3, null=True, blank=True)
    cut_length_m = models.DecimalField(max_digits=14, decimal_places=3, null=True, blank=True)
    wastage = models.DecimalField(max_digits=14, decimal_places=3, null=True, blank=True)
    signature_name = models.CharField(max_length=100, blank=True)

    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['date', 'id']
        verbose_name = 'Flexo Slitting/Cutting Log Entry'

    def __str__(self):
        return f"{self.job_card.job_card_no} slitting/cutting {self.date or ''}"
