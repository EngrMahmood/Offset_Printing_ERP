import math
import re
from decimal import Decimal, ROUND_HALF_UP

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

PURCHASE_MATERIAL_ORIGIN_CHOICES = [
    ('', 'Select Origin'),
    ('local', 'Local'),
    ('import', 'Import'),
]

PLANNING_STAGE_CHOICES = [
    ('', 'Not Set'),
    ('jc_ready', 'JC Ready'),
    ('new_plate_making', 'New Plate Making'),
    ('repeat_plate_making', 'Repeat Plate Making'),
    ('plate_received', 'Plate Received'),
    ('planning_done', 'Planning Done'),
]

PLANNING_STAGE_DONE = 'planning_done'



class PlanningJob(models.Model):
    jc_number = models.CharField(max_length=50, unique=True)
    plan_month = models.CharField(max_length=20, blank=True)
    plan_date = models.DateField(null=True, blank=True)
    po_approval_date = models.DateField(null=True, blank=True)
    delivery_date = models.DateField(null=True, blank=True)

    po_number = models.CharField(max_length=120, blank=True)
    sku = models.CharField(max_length=255, blank=True)
    job_name = models.CharField(max_length=255, blank=True)
    repeat_flag = models.CharField(max_length=50, blank=True)

    material = models.CharField(max_length=120, blank=True)
    color_spec = models.CharField(max_length=60, blank=True)
    application = models.CharField(max_length=120, blank=True)

    size_w_mm = models.IntegerField(null=True, blank=True)
    size_h_mm = models.IntegerField(null=True, blank=True)
    size_w_inch = models.DecimalField(max_digits=10, decimal_places=2, null=True, blank=True)
    size_h_inch = models.DecimalField(max_digits=10, decimal_places=2, null=True, blank=True)

    order_qty = models.PositiveIntegerField(null=True, blank=True)
    print_pcs = models.PositiveIntegerField(null=True, blank=True)
    ups = models.IntegerField(null=True, blank=True)

    print_sheet_size = models.CharField(max_length=80, blank=True)
    print_sheets = models.PositiveIntegerField(null=True, blank=True)
    wastage_sheets = models.PositiveIntegerField(null=True, blank=True)
    actual_sheet_required = models.PositiveIntegerField(null=True, blank=True)

    purchase_sheet_size = models.CharField(max_length=80, blank=True)
    purchase_sheet_ups = models.IntegerField(null=True, blank=True)
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
    JOB_PROCESS_TYPE_CHOICES = [
        ('print_and_pack', 'Print + Pack'),
        ('cut_and_pack', 'Cut & Pack (no printing)'),
    ]
    job_process_type = models.CharField(
        max_length=20,
        choices=JOB_PROCESS_TYPE_CHOICES,
        default='print_and_pack',
        db_index=True,
    )
    print_passes = models.PositiveSmallIntegerField(
        null=True,
        blank=True,
        help_text='Number of press passes (1, 2, 3, or 4). Total impressions = print sheets × passes.',
    )
    planned_total_impressions = models.PositiveIntegerField(null=True, blank=True)

    mi_quantity = models.PositiveIntegerField(null=True, blank=True)
    mi_balance = models.PositiveIntegerField(null=True, blank=True)

    remaining_sheet = models.PositiveIntegerField(null=True, blank=True)
    status = models.CharField(max_length=40, choices=PLANNING_STATUS_CHOICES, default='draft', blank=True)
    planning_stage = models.CharField(max_length=40, choices=PLANNING_STAGE_CHOICES, default='', blank=True)
    planning_stage_changed_at = models.DateTimeField(null=True, blank=True)
    planning_stage_changed_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='planning_jobs_stage_changed',
    )
    pr_reference = models.CharField(max_length=120, blank=True)
    change_request_pending = models.BooleanField(default=False)
    change_request_reason = models.TextField(blank=True)
    change_requested_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='planning_jobs_change_requested',
    )
    change_requested_at = models.DateTimeField(null=True, blank=True)

    master_sync_requested = models.BooleanField(default=False)
    master_sync_reason = models.TextField(blank=True)
    master_sync_requested_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='planning_jobs_master_sync_requested',
    )
    master_sync_requested_at = models.DateTimeField(null=True, blank=True)
    master_sync_applied_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='planning_jobs_master_sync_applied',
    )
    master_sync_applied_at = models.DateTimeField(null=True, blank=True)

    rejected_qty = models.PositiveIntegerField(null=True, blank=True)
    balance_qty = models.PositiveIntegerField(null=True, blank=True)
    destination = models.CharField(max_length=120, blank=True)

    unit_cost = models.DecimalField(max_digits=12, decimal_places=2, null=True, blank=True)
    stock_bag = models.DecimalField(max_digits=12, decimal_places=1, null=True, blank=True)

    machine_name = models.CharField(max_length=120, blank=True)
    purchase_material_origin = models.CharField(max_length=20, choices=PURCHASE_MATERIAL_ORIGIN_CHOICES, blank=True)
    stock_qty = models.DecimalField(max_digits=12, decimal_places=1, null=True, blank=True)
    daily_demand = models.DecimalField(max_digits=12, decimal_places=1, null=True, blank=True)

    department = models.CharField(max_length=120, blank=True)
    plate_set_no = models.CharField(max_length=120, blank=True)
    aging_days = models.PositiveIntegerField(null=True, blank=True)

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
        if self.plan_date:
            return self.plan_date
        po_document = self.po_documents.order_by('created_at').first() if hasattr(self, 'po_documents') else None
        if po_document and po_document.created_at:
            return po_document.created_at.date()
        if self.created_at:
            return self.created_at.date()
        return self.plan_date

    @property
    def net_print_qty(self):
        """Net production qty after stock absorption; negative values are clamped to zero."""
        if self.order_qty is None:
            return None
        stock_consumption = int(self.stock_qty or 0)
        return max((self.order_qty or 0) - stock_consumption, 0)

    @property
    def ups_value(self):
        if self.ups is not None:
            return self.ups
        recipe = self.sku_recipe
        if recipe and recipe.ups is not None:
            return recipe.ups
        return None

    @property
    def calculated_sheets_required(self):
        """Auto-calculate total sheets required from net print qty, UPS and wastage."""
        net_qty = self.net_print_qty
        ups_value = self.ups_value
        if net_qty is not None and ups_value:
            return math.ceil(net_qty / ups_value) + (self.wastage_sheets or 0)
        if self.print_sheets is not None:
            return self.print_sheets + (self.wastage_sheets or 0)
        if self.actual_sheet_required is not None:
            return self.actual_sheet_required
        return None

    @property
    def calculated_purchase_sheet_required(self):
        sheets_required = self.calculated_sheets_required
        if sheets_required is None:
            return None
        purchase_sheet_ups_value = self.purchase_sheet_ups_display
        if not purchase_sheet_ups_value:
            return None
        return math.ceil(sheets_required / purchase_sheet_ups_value)

    @property
    def calculated_pkt_value(self):
        purchase_sheets = self.purchase_sheet_required_display
        if purchase_sheets is None:
            return None
        return (Decimal(purchase_sheets) / Decimal('100')).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)

    @property
    def sku_recipe(self):
        if not (self.sku or '').strip():
            return None
        return SkuRecipe.objects.filter(sku__iexact=self.sku).order_by('-updated_at').first()

    @property
    def approved_sku_recipe(self):
        if not (self.sku or '').strip():
            return None
        return SkuRecipe.objects.filter(
            sku__iexact=self.sku,
            is_active=True,
            master_data_status='approved',
        ).first()

    def master_data_sync_blocked(self):
        return self.workflow_status == 'completed'

    @property
    def awc_no_display(self):
        """AWC lives on SKU master; fall back to plate requests for the same SKU/job."""
        from planning.services import normalize_awc_no

        recipe = self.sku_recipe
        if recipe:
            value = normalize_awc_no(recipe.awc_no)
            if value:
                return value

        from django.db.models import Q
        from printing_plates.models import PlateRequest

        plate_qs = PlateRequest.objects.exclude(awc_no='').order_by('-updated_at')
        filters = Q(planning_job_id=self.pk)
        if recipe is not None:
            filters |= Q(sku_recipe_id=recipe.pk)
        sku = (self.sku or '').strip()
        if sku:
            filters |= Q(job_card__SKU__iexact=sku) | Q(sku_recipe__sku__iexact=sku)
        plate = plate_qs.filter(filters).first()
        if plate:
            return normalize_awc_no(plate.awc_no)
        return ''

    @property
    def die_cutting_display(self):
        """Die cutting is Yes/No on master; fall back to plate request."""
        from planning.services import normalize_die_cutting

        recipe = self.sku_recipe
        if recipe:
            value = normalize_die_cutting(recipe.die_cutting)
            if value:
                return value

        from django.db.models import Q
        from printing_plates.models import PlateRequest

        plate_qs = PlateRequest.objects.exclude(die_cutting='').order_by('-updated_at')
        filters = Q(planning_job_id=self.pk)
        if recipe is not None:
            filters |= Q(sku_recipe_id=recipe.pk)
        sku = (self.sku or '').strip()
        if sku:
            filters |= Q(job_card__SKU__iexact=sku) | Q(sku_recipe__sku__iexact=sku)
        plate = plate_qs.filter(filters).first()
        if plate:
            return normalize_die_cutting(plate.die_cutting)
        return ''

    @property
    def material_display(self):
        recipe = self.sku_recipe
        if (self.material or '').strip():
            return self.material
        if recipe and (recipe.material or '').strip():
            return recipe.material
        return ''

    @property
    def color_spec_display(self):
        recipe = self.approved_sku_recipe or self.sku_recipe
        if (self.color_spec or '').strip():
            return self.color_spec
        if recipe and (recipe.color_spec or '').strip():
            return recipe.color_spec
        return ''

    @property
    def application_display(self):
        recipe = self.sku_recipe
        if (self.application or '').strip():
            return self.application
        if recipe and (recipe.application or '').strip():
            return recipe.application
        return ''

    @property
    def size_w_mm_display(self):
        recipe = self.sku_recipe
        if self.size_w_mm is not None:
            return self.size_w_mm
        if recipe and recipe.size_w_mm is not None:
            return recipe.size_w_mm
        return None

    @property
    def size_h_mm_display(self):
        recipe = self.sku_recipe
        if self.size_h_mm is not None:
            return self.size_h_mm
        if recipe and recipe.size_h_mm is not None:
            return recipe.size_h_mm
        return None

    @property
    def print_sheet_size_display(self):
        recipe = self.sku_recipe
        if (self.print_sheet_size or '').strip():
            return self.print_sheet_size
        if recipe and (recipe.print_sheet_size or '').strip():
            return recipe.print_sheet_size
        return ''

    @property
    def purchase_sheet_size_display(self):
        recipe = self.sku_recipe
        if (self.purchase_sheet_size or '').strip():
            return self.purchase_sheet_size
        if recipe and (recipe.purchase_sheet_size or '').strip():
            return recipe.purchase_sheet_size
        return ''

    @property
    def ups_display(self):
        recipe = self.sku_recipe
        if self.ups is not None:
            return self.ups
        if recipe and recipe.ups is not None:
            return recipe.ups
        return None

    @property
    def purchase_sheet_ups_display(self):
        recipe = self.sku_recipe
        if self.purchase_sheet_ups is not None:
            return self.purchase_sheet_ups
        if recipe and recipe.purchase_sheet_ups is not None:
            return recipe.purchase_sheet_ups
        return None

    @property
    def purchase_sheet_required_display(self):
        calculated = self.calculated_purchase_sheet_required
        if calculated is not None:
            return calculated
        return self.purchase_sheet_required

    @property
    def actual_sheet_required_display(self):
        calculated = self.calculated_sheets_required
        if calculated is not None:
            return calculated
        return self.actual_sheet_required

    @property
    def number_of_colors(self):
        from core.print_colors import print_color_total_units

        units = print_color_total_units(self.color_spec_display)
        if units:
            return units
        if self.total_colors is not None:
            return self.total_colors
        if self.front_colors is not None or self.back_colors is not None:
            return (self.front_colors or 0) + (self.back_colors or 0)
        return None

    @property
    def remarks_display(self):
        """JC Remarks column shows SKU master notes (sheet Remarks), not SKU remarks field."""
        recipe = self.approved_sku_recipe or self.sku_recipe
        recipe_notes = (recipe.notes or '').strip() if recipe else ''
        if recipe_notes and recipe_notes != '-':
            return recipe_notes
        remarks_text = (self.remarks or '').strip()
        if remarks_text and remarks_text != '-':
            return remarks_text
        return ''

    @property
    def effective_print_passes(self):
        """Passes owned by SKU master; job field is a synced snapshot for production."""
        if self.is_cut_and_pack():
            return None
        if self.print_passes:
            return int(self.print_passes)
        recipe = self.approved_sku_recipe or self.sku_recipe
        if recipe and recipe.print_passes:
            return int(recipe.print_passes)
        return None

    @property
    def calculated_planned_total_impressions(self):
        """Auto impressions = total print sheets × no. of passes."""
        sheets = self.calculated_sheets_required
        passes = self.effective_print_passes
        if sheets is None or not passes:
            return None
        return int(sheets) * int(passes)

    def sync_planned_total_impressions(self):
        calculated = self.calculated_planned_total_impressions
        if calculated is not None:
            self.planned_total_impressions = calculated
        return self.planned_total_impressions

    @property
    def has_completed_plate_request(self):
        return self.plate_requests.filter(status='available_for_production').exists()

    @property
    def latest_cancelled_plate_request(self):
        """Most recent plate request cancelled by graphics (plates not required)."""
        for req in self.plate_requests.all():
            if getattr(req, 'is_cancelled', False):
                return req
        # Fallback if not prefetched / ordering differs
        return (
            self.plate_requests.filter(progress__istartswith='Cancelled')
            .order_by('-requested_at', '-created_at', '-id')
            .first()
        )

    @property
    def effective_machine_name(self):
        if str(self.machine_name or '').strip():
            return str(self.machine_name).strip()
        recipe = self.approved_sku_recipe or self.sku_recipe
        if recipe and str(recipe.machine_name or '').strip():
            return str(recipe.machine_name).strip()
        return ''

    @property
    def effective_plate_set_no(self):
        if str(self.plate_set_no or '').strip():
            return str(self.plate_set_no).strip()
        recipe = self.approved_sku_recipe or self.sku_recipe
        if recipe and str(recipe.plate_set_no or '').strip():
            return str(recipe.plate_set_no).strip()
        return ''

    @property
    def job_process_type_display(self):
        """Job process is owned by SKU master; planning does not override."""
        recipe = self.approved_sku_recipe or self.sku_recipe
        if recipe and (recipe.job_process_type or '').strip():
            return recipe.job_process_type.strip()
        return (self.job_process_type or 'print_and_pack').strip() or 'print_and_pack'

    @property
    def job_process_type_label(self):
        value = self.job_process_type_display
        return dict(self.JOB_PROCESS_TYPE_CHOICES).get(value, value)

    def is_cut_and_pack(self):
        return self.job_process_type_display == 'cut_and_pack'

    def sync_job_process_type_from_sku_master(self):
        """Copy job process from SKU master onto this job (no planning override)."""
        recipe = self.approved_sku_recipe or self.sku_recipe
        if not recipe and (self.sku or '').strip():
            from planning.services import get_best_sku_recipe_for_sku
            recipe = get_best_sku_recipe_for_sku(self.sku)
        process = ''
        if recipe and (recipe.job_process_type or '').strip():
            process = recipe.job_process_type.strip()
        if not process:
            return False
        if (self.job_process_type or 'print_and_pack') == process:
            return False
        self.job_process_type = process
        return True

    def _print_passes_frozen(self):
        return (self.workflow_status or '').strip().lower() in {
            'qc_approved',
            'released',
            'in_production',
            'completed',
            'closed',
        }

    def sync_print_passes_from_sku_master(self):
        """Copy print passes from SKU master onto open jobs; freeze after QC approval."""
        if self._print_passes_frozen():
            return False

        if self.is_cut_and_pack():
            if self.print_passes is not None:
                self.print_passes = None
                return True
            return False

        recipe = self.approved_sku_recipe or self.sku_recipe
        if not recipe and (self.sku or '').strip():
            from planning.services import get_best_sku_recipe_for_sku
            recipe = get_best_sku_recipe_for_sku(self.sku)
        master_passes = getattr(recipe, 'print_passes', None) if recipe else None
        if master_passes is None:
            return False
        master_passes = int(master_passes)
        if self.print_passes == master_passes:
            return False
        self.print_passes = master_passes
        return True

    def pre_submit_qc_validation_errors(self):
        errors = {}
        if self.is_cut_and_pack:
            if self.wastage_sheets is None:
                errors['wastage_sheets'] = 'Wastage is required before QC approval.'
            if not str(self.purchase_material_origin or '').strip():
                errors['purchase_material_origin'] = 'Purchase Material Origin is required before QC approval.'
            return errors
        if self.wastage_sheets is None:
            errors['wastage_sheets'] = 'Wastage is required before QC approval.'
        if not self.effective_machine_name:
            errors['machine_name'] = 'Machine Name is required before QC approval.'
        if not str(self.purchase_material_origin or '').strip():
            errors['purchase_material_origin'] = 'Purchase Material Origin is required before QC approval.'
        passes = self.effective_print_passes
        if not passes:
            errors['print_passes'] = 'No. of Passes is required on SKU master before QC approval.'
        elif self.calculated_sheets_required is None:
            errors['print_passes'] = 'Cannot calculate impressions until print sheets are available (order qty, UPS, wastage).'
        return errors

    def qc_validation_errors(self):
        if self.workflow_status not in PLANNING_QC_GATE_STATUSES:
            return {}
        return self.pre_submit_qc_validation_errors()

    def qc_missing_fields(self):
        errors = self.qc_validation_errors()
        return [field.replace('_', ' ').title() for field in errors.keys()]

    def clean(self):
        super().clean()

    def save(self, *args, **kwargs):
        from core.print_colors import apply_print_color_to_planning_job

        update_fields = kwargs.get('update_fields')
        if update_fields is not None:
            update_fields = set(update_fields)

        self.status = self.workflow_status
        if update_fields is not None:
            update_fields.add('status')

        if apply_print_color_to_planning_job(self):
            if update_fields is not None:
                update_fields.add('color_spec')
                update_fields.add('total_colors')

        if self.sync_job_process_type_from_sku_master():
            if update_fields is not None:
                update_fields.add('job_process_type')

        if self.sync_print_passes_from_sku_master():
            if update_fields is not None:
                update_fields.add('print_passes')

        calculated_total_sheet_quantity = self.calculated_sheets_required
        if calculated_total_sheet_quantity is not None:
            self.actual_sheet_required = calculated_total_sheet_quantity
            if update_fields is not None:
                update_fields.add('actual_sheet_required')

        calculated_purchase_sheet_required = self.calculated_purchase_sheet_required
        if calculated_purchase_sheet_required is not None:
            self.purchase_sheet_required = calculated_purchase_sheet_required
            if update_fields is not None:
                update_fields.add('purchase_sheet_required')

        calculated_pkt_value = self.calculated_pkt_value
        if calculated_pkt_value is not None:
            self.pkt_value = calculated_pkt_value
            if update_fields is not None:
                update_fields.add('pkt_value')

        delivered_qty_sum = self.dispatch_runs.aggregate(total= models.Sum('delivered_qty'))['total'] if self.pk else None
        if self.order_qty is not None:
            self.balance_qty = max((self.order_qty or 0) - int(delivered_qty_sum or 0), 0)
            if update_fields is not None:
                update_fields.add('balance_qty')

        calculated_number_of_colors = self.number_of_colors
        if calculated_number_of_colors is not None:
            self.total_colors = calculated_number_of_colors
            if update_fields is not None:
                update_fields.add('total_colors')

        passes = self.effective_print_passes
        if passes and self.calculated_sheets_required is not None:
            self.planned_total_impressions = int(self.calculated_sheets_required) * int(passes)
            if update_fields is not None:
                update_fields.add('planned_total_impressions')

        self.full_clean()
        if update_fields is not None:
            kwargs['update_fields'] = list(update_fields)
        result = super().save(*args, **kwargs)

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


class JobCardLayout(models.Model):
    name = models.CharField(max_length=120, default='Job Card Layout')
    layout = models.JSONField(blank=True, default=list)
    is_active = models.BooleanField(default=False)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['-is_active', '-updated_at']
        verbose_name = 'Job Card Layout'
        verbose_name_plural = 'Job Card Layouts'

    def __str__(self):
        return self.name

    @classmethod
    def get_active_layout(cls):
        return cls.objects.filter(is_active=True).order_by('-updated_at').first()


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
    legacy_produced = models.BooleanField(
        default=False,
        help_text='True when SKU master came from Google Sheet / bulk upload and was produced before ERP go-live.',
    )
    job_name = models.CharField(max_length=255, blank=True)

    material = models.CharField(max_length=120, blank=True)
    color_spec = models.CharField(max_length=60, blank=True)
    application = models.CharField(max_length=120, blank=True)
    product_type = models.CharField(max_length=100, blank=True)
    machine_name = models.CharField(max_length=120, blank=True)
    JOB_PROCESS_TYPE_CHOICES = [
        ('print_and_pack', 'Print + Pack'),
        ('cut_and_pack', 'Cut & Pack (no printing)'),
    ]
    job_process_type = models.CharField(
        max_length=20,
        choices=JOB_PROCESS_TYPE_CHOICES,
        default='print_and_pack',
        help_text='Default process for jobs using this SKU. Cut & Pack skips printing/plates.',
    )
    print_passes = models.PositiveSmallIntegerField(
        null=True,
        blank=True,
        help_text='Number of press passes (1, 2, 3, or 4) for Print + Pack SKUs.',
    )
    plate_set_no = models.CharField(max_length=120, blank=True)

    size_w_mm = models.IntegerField(null=True, blank=True)
    size_h_mm = models.IntegerField(null=True, blank=True)
    ups = models.IntegerField(null=True, blank=True)

    print_sheet_size = models.CharField(max_length=80, blank=True)
    purchase_sheet_size = models.CharField(max_length=80, blank=True)
    purchase_sheet_ups = models.IntegerField(null=True, blank=True)

    default_unit_cost = models.DecimalField(max_digits=12, decimal_places=2, null=True, blank=True)
    daily_demand = models.DecimalField(max_digits=12, decimal_places=1, null=True, blank=True)
    awc_no = models.CharField(max_length=120, blank=True)
    die_cutting = models.CharField(max_length=120, blank=True)
    notes = models.TextField(blank=True)
    remarks = models.TextField(blank=True)
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

    def get_next_notification_roles(self):
        status = (self.master_data_status or '').strip().lower()
        if status == 'draft':
            return ['qc']
        elif status == 'pending_review':
            return ['qc']
        elif status == 'reviewed':
            return ['admin', 'manager']
        return []


class JobCardChangeRequest(models.Model):
    STATUS_CHOICES = [
        ('pending', 'Pending PM Approval'),
        ('approved', 'Approved'),
        ('rejected', 'Rejected'),
    ]

    planning_job = models.ForeignKey(
        PlanningJob,
        on_delete=models.CASCADE,
        related_name='change_requests'
    )
    
    request_type = models.CharField(max_length=50, default='reopen_to_draft')
    
    current_wastage_sheets = models.PositiveIntegerField(null=True, blank=True)
    proposed_wastage_sheets = models.PositiveIntegerField(null=True, blank=True)
    
    current_machine_name = models.CharField(max_length=120, blank=True, null=True)
    proposed_machine_name = models.CharField(max_length=120, blank=True, null=True)
    
    reason = models.TextField()
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='pending')
    
    requested_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        related_name='requested_job_changes'
    )
    requested_at = models.DateTimeField(auto_now_add=True)
    
    approved_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='reviewed_job_changes'
    )
    approved_at = models.DateTimeField(null=True, blank=True)
    rejection_reason = models.TextField(blank=True, null=True)

    class Meta:
        ordering = ['-requested_at']

    def __str__(self):
        return f"Reopen Request for {self.planning_job.jc_number} ({self.status})"



