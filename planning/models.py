from django.conf import settings
from django.db import models


class PlanningJob(models.Model):
    jc_number = models.CharField(max_length=50, unique=True)
    plan_month = models.CharField(max_length=20, blank=True)
    plan_date = models.DateField(null=True, blank=True)

    po_number = models.CharField(max_length=120, blank=True)
    sku = models.CharField(max_length=255, blank=True)
    job_name = models.CharField(max_length=255, blank=True)
    repeat_flag = models.CharField(max_length=50, blank=True)

    material = models.CharField(max_length=120, blank=True)
    color_spec = models.CharField(max_length=60, blank=True)
    application = models.CharField(max_length=120, blank=True)

    size_w_mm = models.DecimalField(max_digits=10, decimal_places=2, null=True, blank=True)
    size_h_mm = models.DecimalField(max_digits=10, decimal_places=2, null=True, blank=True)
    size_w_inch = models.DecimalField(max_digits=10, decimal_places=2, null=True, blank=True)
    size_h_inch = models.DecimalField(max_digits=10, decimal_places=2, null=True, blank=True)

    order_qty = models.PositiveIntegerField(null=True, blank=True)
    print_pcs = models.PositiveIntegerField(null=True, blank=True)
    ups = models.PositiveIntegerField(null=True, blank=True)

    print_sheet_size = models.CharField(max_length=80, blank=True)
    print_sheets = models.PositiveIntegerField(null=True, blank=True)
    wastage_sheets = models.PositiveIntegerField(null=True, blank=True)
    actual_sheet_required = models.PositiveIntegerField(null=True, blank=True)

    purchase_sheet_size = models.CharField(max_length=80, blank=True)
    purchase_sheet_ups = models.PositiveIntegerField(null=True, blank=True)
    purchase_sheet_required = models.PositiveIntegerField(null=True, blank=True)

    pkt_value = models.DecimalField(max_digits=12, decimal_places=2, null=True, blank=True)
    remarks = models.TextField(blank=True)
    requirement = models.TextField(blank=True)

    front_colors = models.PositiveIntegerField(null=True, blank=True)
    back_colors = models.PositiveIntegerField(null=True, blank=True)
    total_colors = models.PositiveIntegerField(null=True, blank=True)
    total_mr_time_minutes = models.PositiveIntegerField(null=True, blank=True)

    front_pass = models.PositiveIntegerField(null=True, blank=True)
    back_pass = models.PositiveIntegerField(null=True, blank=True)
    planned_total_impressions = models.PositiveIntegerField(null=True, blank=True)

    mi_quantity = models.PositiveIntegerField(null=True, blank=True)
    mi_balance = models.PositiveIntegerField(null=True, blank=True)

    remaining_sheet = models.PositiveIntegerField(null=True, blank=True)
    status = models.CharField(max_length=40, blank=True)
    pr_reference = models.CharField(max_length=120, blank=True)

    rejected_qty = models.PositiveIntegerField(null=True, blank=True)
    balance_qty = models.PositiveIntegerField(null=True, blank=True)
    destination = models.CharField(max_length=120, blank=True)

    unit_cost = models.DecimalField(max_digits=12, decimal_places=2, null=True, blank=True)
    stock_bag = models.DecimalField(max_digits=12, decimal_places=2, null=True, blank=True)

    machine_name = models.CharField(max_length=120, blank=True)
    purchase_material = models.CharField(max_length=120, blank=True)
    stock_qty = models.DecimalField(max_digits=12, decimal_places=2, null=True, blank=True)
    daily_demand = models.DecimalField(max_digits=12, decimal_places=2, null=True, blank=True)

    department = models.CharField(max_length=120, blank=True)
    plate_set_no = models.CharField(max_length=120, blank=True)
    awc_no = models.CharField(max_length=120, blank=True)
    aging_days = models.PositiveIntegerField(null=True, blank=True)
    die_cutting = models.CharField(max_length=120, blank=True)

    issued_to_production = models.BooleanField(default=False)
    job_card_version = models.PositiveIntegerField(default=1)
    has_edits_since_creation = models.BooleanField(default=False)
    edited_fields_list = models.JSONField(default=list, blank=True)

    created_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='planning_jobs_created',
    )
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    last_edited_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='planning_jobs_edited',
    )
    last_edited_at = models.DateTimeField(null=True, blank=True)

    class Meta:
        ordering = ['-plan_date', '-id']

    def __str__(self):
        return f"{self.jc_number} | {self.sku}" if self.sku else self.jc_number


class PlanningPrintRun(models.Model):
    planning_job = models.ForeignKey(PlanningJob, on_delete=models.CASCADE, related_name='print_runs')
    run_index = models.PositiveSmallIntegerField()
    print_date = models.DateField(null=True, blank=True)
    print_qty = models.PositiveIntegerField(null=True, blank=True)
    wastage_qty = models.PositiveIntegerField(null=True, blank=True)

    class Meta:
        unique_together = ('planning_job', 'run_index')
        ordering = ['run_index']


class PlanningDispatchRun(models.Model):
    planning_job = models.ForeignKey(PlanningJob, on_delete=models.CASCADE, related_name='dispatch_runs')
    dispatch_index = models.PositiveSmallIntegerField()
    delivery_date = models.DateField(null=True, blank=True)
    dc_no = models.CharField(max_length=80, blank=True)
    delivered_qty = models.PositiveIntegerField(null=True, blank=True)

    class Meta:
        unique_together = ('planning_job', 'dispatch_index')
        ordering = ['dispatch_index']


class PoDocument(models.Model):
    STATUS_CHOICES = [
        ('pending', 'Pending'),
        ('processed', 'Processed'),
        ('failed', 'Failed'),
    ]

    planning_job = models.ForeignKey(
        PlanningJob,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='po_documents',
    )
    po_file = models.FileField(upload_to='planning/po_docs/')
    extracted_payload = models.JSONField(null=True, blank=True)
    extraction_status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='pending')
    uploaded_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='uploaded_po_documents',
    )
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['-created_at']


class SkuRecipe(models.Model):
    lamination_front_and_back = models.BooleanField(default=False, help_text='Lamination is applied on both front and back')
    MASTER_DATA_STATUS_CHOICES = [
        ('draft', 'Draft'),
        ('pending_review', 'Pending Review'),
        ('reviewed', 'Pending Approval (Manager)'),
        ('approved', 'Approved'),
    ]

    sku = models.CharField(max_length=255, unique=True)
    job_name = models.CharField(max_length=255, blank=True)

    material = models.CharField(max_length=120, blank=True)
    color_spec = models.CharField(max_length=60, blank=True)
    application = models.CharField(max_length=120, blank=True)

    size_w_mm = models.DecimalField(max_digits=10, decimal_places=2, null=True, blank=True)
    size_h_mm = models.DecimalField(max_digits=10, decimal_places=2, null=True, blank=True)
    ups = models.PositiveIntegerField(null=True, blank=True)

    print_sheet_size = models.CharField(max_length=80, blank=True)
    purchase_sheet_size = models.CharField(max_length=80, blank=True)
    purchase_sheet_ups = models.PositiveIntegerField(null=True, blank=True)
    purchase_material = models.CharField(max_length=120, blank=True)

    machine_name = models.CharField(max_length=120, blank=True)
    department = models.CharField(max_length=120, blank=True)
    default_unit_cost = models.DecimalField(max_digits=12, decimal_places=2, null=True, blank=True)
    daily_demand = models.DecimalField(max_digits=12, decimal_places=2, null=True, blank=True)
    awc_no = models.CharField(max_length=120, blank=True)
    plate_set_no = models.CharField(max_length=120, blank=True)
    die_cutting = models.CharField(max_length=120, blank=True)
    notes = models.TextField(blank=True)
    is_active = models.BooleanField(default=True)
    archive_reason = models.TextField(blank=True)
    archived_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='sku_recipes_archived',
    )
    archived_at = models.DateTimeField(null=True, blank=True)
    master_data_status = models.CharField(
        max_length=20,
        choices=MASTER_DATA_STATUS_CHOICES,
        default='draft',
    )
    reviewed_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='sku_recipes_reviewed',
    )
    reviewed_at = models.DateTimeField(null=True, blank=True)
    approved_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='sku_recipes_approved',
    )
    approved_at = models.DateTimeField(null=True, blank=True)

    rejection_comment = models.TextField(blank=True)
    last_rejected_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='sku_recipes_rejected',
    )
    last_rejected_at = models.DateTimeField(null=True, blank=True)

    created_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='sku_recipes_created',
    )
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['sku']

    def __str__(self):
        return self.sku
