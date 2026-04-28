import math
import re

from django.conf import settings
from django.core.exceptions import ValidationError
from django.db import models


PLANNING_STATUS_CHOICES = [
    ('draft', 'Draft'),
    ('pending_qc', 'Pending QC'),
    ('qc_approved', 'QC Approved'),
    ('released', 'Released'),
    ('in_production', 'In Production'),
    ('completed', 'Completed'),
]

PLANNING_STATUS_ALIASES = {
    'open': 'draft',
    'pending': 'draft',
    'reviewed': 'pending_qc',
    'approved': 'qc_approved',
    'closed': 'completed',
}

PLANNING_QC_GATE_STATUSES = {
    'pending_qc',
    'qc_approved',
    'released',
    'in_production',
    'completed',
}


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
    status = models.CharField(max_length=40, choices=PLANNING_STATUS_CHOICES, default='draft', blank=True)
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
    is_active = models.BooleanField(default=True)
    is_on_hold = models.BooleanField(default=False)
    hold_reason = models.TextField(blank=True)
    hold_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='planning_jobs_held',
    )
    hold_at = models.DateTimeField(null=True, blank=True)
    archived_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='planning_jobs_archived',
    )
    archived_at = models.DateTimeField(null=True, blank=True)
    archive_reason = models.TextField(blank=True)
    restored_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='planning_jobs_restored',
    )
    restored_at = models.DateTimeField(null=True, blank=True)
    restore_reason = models.TextField(blank=True)
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

    @property
    def workflow_status(self):
        raw_status = (self.status or '').strip().lower()
        return PLANNING_STATUS_ALIASES.get(raw_status, raw_status or 'draft')

    @property
    def workflow_status_label(self):
        normalized_status = self.workflow_status
        status_labels = dict(PLANNING_STATUS_CHOICES)
        return status_labels.get(normalized_status, normalized_status.replace('_', ' ').title())

    @property
    def total_sheet_quantity(self):
        return self.calculated_sheets_required

    @property
    def po_received_date(self):
        po_document = self.po_documents.order_by('created_at').first() if hasattr(self, 'po_documents') else None
        if po_document and po_document.created_at:
            return po_document.created_at.date()
        if self.created_at:
            return self.created_at.date()
        return self.plan_date

    @property
    def calculated_sheets_required(self):
        """Auto-calculate total sheets required from order qty, UPS and wastage."""
        if self.order_qty is not None and self.ups:
            return math.ceil(self.order_qty / self.ups) + (self.wastage_sheets or 0)
        if self.print_sheets is not None:
            return self.print_sheets + (self.wastage_sheets or 0)
        if self.actual_sheet_required is not None:
            return self.actual_sheet_required
        return None

    @property
    def number_of_colors(self):
        if self.total_colors is not None:
            return self.total_colors
        if self.front_colors is not None or self.back_colors is not None:
            return (self.front_colors or 0) + (self.back_colors or 0)

        raw_value = (self.color_spec or '').strip()
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

    def qc_validation_errors(self):
        if self.workflow_status not in PLANNING_QC_GATE_STATUSES:
            return {}

        errors = {}
        if not str(self.plate_set_no or '').strip():
            errors['plate_set_no'] = 'Plate Set is required before QC approval.'
        if self.wastage_sheets is None:
            errors['wastage_sheets'] = 'Wastage is required before QC approval.'
        if not str(self.machine_name or '').strip():
            errors['machine_name'] = 'Machine Name is required before QC approval.'
        if not str(self.remarks or '').strip():
            errors['remarks'] = 'Remarks are required before QC approval.'
        return errors

    def qc_missing_fields(self):
        errors = self.qc_validation_errors()
        return [field.replace('_', ' ').title() for field in errors.keys()]

    def clean(self):
        super().clean()
        errors = self.qc_validation_errors()
        if errors:
            raise ValidationError(errors)

    def save(self, *args, **kwargs):
        update_fields = kwargs.get('update_fields')
        if update_fields is not None:
            update_fields = set(update_fields)

        self.status = self.workflow_status
        if update_fields is not None:
            update_fields.add('status')

        calculated_total_sheet_quantity = self.calculated_sheets_required
        if calculated_total_sheet_quantity is not None:
            self.actual_sheet_required = calculated_total_sheet_quantity
            if update_fields is not None:
                update_fields.add('actual_sheet_required')

        calculated_number_of_colors = self.number_of_colors
        if calculated_number_of_colors is not None:
            self.total_colors = calculated_number_of_colors
            if update_fields is not None:
                update_fields.add('total_colors')

        self.full_clean()
        if update_fields is not None:
            kwargs['update_fields'] = list(update_fields)
        result = super().save(*args, **kwargs)

        if self.sku and self.order_qty is not None:
            try:
                from core.services import ensure_job_card_from_planning_job

                ensure_job_card_from_planning_job(self, actor=self.last_edited_by or self.created_by)
            except Exception:
                raise

        return result


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
