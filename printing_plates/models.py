from django.conf import settings
from django.db import models


class PlateRequest(models.Model):
    STATUS_DRAFT = 'draft'
    STATUS_SENT = 'sent_to_vendor'
    STATUS_RECEIVED = 'received_from_vendor'
    STATUS_AVAILABLE = 'available_for_production'
    STATUS_ARCHIVED = 'archived'

    STATUS_CHOICES = [
        (STATUS_DRAFT, 'Draft'),
        (STATUS_SENT, 'Sent to Vendor'),
        (STATUS_RECEIVED, 'Received from Vendor'),
        (STATUS_AVAILABLE, 'Available for Production'),
        (STATUS_ARCHIVED, 'Archived'),
    ]

    SOURCE_REPLACEMENT = 'replacement'
    SOURCE_PLANNING = 'planning'
    SOURCE_PRODUCTION_PLATE_DAMAGE = 'production_plate_damage'

    REASON_DAMAGED_DURING_RUN = 'damaged_during_run'
    REASON_DAMAGED_BEFORE_PRINTING = 'damaged_before_printing'
    REASON_WRONG_PLATES = 'wrong_plates_received'
    REASON_WORN_OUT = 'worn_out'
    REASON_VENDOR_DEFECT = 'vendor_defect'
    REASON_OTHER = 'other'

    REPLACEMENT_REASON_CHOICES = [
        (REASON_DAMAGED_DURING_RUN, 'Damaged during run'),
        (REASON_DAMAGED_BEFORE_PRINTING, 'Damaged before printing'),
        (REASON_WRONG_PLATES, 'Wrong plates received'),
        (REASON_WORN_OUT, 'Worn out'),
        (REASON_VENDOR_DEFECT, 'Vendor defect'),
        (REASON_OTHER, 'Other'),
    ]

    OPEN_STATUSES = (
        STATUS_DRAFT,
        STATUS_SENT,
        STATUS_RECEIVED,
    )

    planning_job = models.ForeignKey(
        'planning.PlanningJob',
        on_delete=models.PROTECT,
        related_name='plate_requests',
        null=True,
        blank=True,
    )
    job_card = models.ForeignKey(
        'core.JobCard',
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='plate_requests',
    )
    sku_recipe = models.ForeignKey(
        'planning.SkuRecipe',
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='plate_requests',
    )
    machine = models.ForeignKey(
        'core.Machine',
        on_delete=models.PROTECT,
        null=True,
        blank=True,
    )
    department = models.ForeignKey(
        'core.Department',
        on_delete=models.PROTECT,
        null=True,
        blank=True,
    )

    status = models.CharField(
        max_length=40,
        choices=STATUS_CHOICES,
        default=STATUS_DRAFT,
    )

    set_no = models.CharField(max_length=120, blank=True)
    new_set_no = models.CharField(max_length=120, blank=True)
    awc_no = models.CharField(max_length=120, blank=True)
    DIE_CUTTING_YES = 'YES'
    DIE_CUTTING_NO = 'NO'
    DIE_CUTTING_CHOICES = [
        ('', 'Select'),
        (DIE_CUTTING_YES, 'Yes'),
        (DIE_CUTTING_NO, 'No'),
    ]
    die_cutting = models.CharField(
        max_length=10,
        choices=DIE_CUTTING_CHOICES,
        blank=True,
        default='',
        help_text='Die cutting required: Yes or No only.',
    )
    plate_quantity = models.PositiveIntegerField(null=True, blank=True)
    sets_required = models.PositiveIntegerField(
        null=True,
        blank=True,
        help_text='Number of plate sets ordered (all issued to production together for now)',
    )
    plate_color = models.CharField(max_length=120, blank=True)
    vendor = models.CharField(max_length=120, blank=True)
    remarks = models.TextField(blank=True)
    source = models.CharField(max_length=120, blank=True)
    replacement_reason = models.CharField(
        max_length=40,
        choices=REPLACEMENT_REASON_CHOICES,
        blank=True,
        default='',
    )
    damaged_colors = models.CharField(
        max_length=120,
        blank=True,
        help_text='Colours needing remake, e.g. Cyan, Magenta',
    )
    replaces_request = models.ForeignKey(
        'self',
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='replacement_requests',
    )
    impression = models.CharField(max_length=120, blank=True)
    progress = models.CharField(max_length=120, blank=True)
    challan = models.CharField(max_length=120, blank=True)
    chalan_sign = models.BooleanField(default=False)
    box = models.CharField(max_length=120, blank=True)
    image = models.ImageField(upload_to='printing_plates/images/', null=True, blank=True)
    link = models.URLField(blank=True)

    requested_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='plate_requests_requested',
    )
    requested_at = models.DateTimeField(null=True, blank=True)
    sent_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='plate_requests_sent',
    )
    sent_at = models.DateTimeField(null=True, blank=True)
    received_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='plate_requests_received',
    )
    received_at = models.DateTimeField(null=True, blank=True)
    designer = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='plate_requests_designed',
    )

    # Smart layout merge: when a job joins a combined layout, its own plate set is
    # either scrapped (superseded by the combined plate) or parked here for a
    # future run of the same SKU instead of being thrown away.
    retained_for_reuse = models.BooleanField(default=False, db_index=True)
    retained_at = models.DateTimeField(null=True, blank=True)
    retained_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='plate_requests_retained',
    )
    retained_reason = models.TextField(blank=True)
    retained_released_at = models.DateTimeField(
        null=True,
        blank=True,
        help_text='Set when a later job for this SKU picked the retained plate set back up.',
    )

    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['-requested_at', '-created_at']
        indexes = [
            models.Index(fields=['planning_job']),
            models.Index(fields=['status']),
            models.Index(fields=['requested_at']),
            models.Index(fields=['source']),
        ]
        verbose_name = 'Plate Request'
        verbose_name_plural = 'Plate Requests'

    def __str__(self):
        identifier = self.set_no or self.new_set_no or str(self.pk)
        return f'Plate Request {identifier} ({self.get_status_display()})'

    def save(self, *args, **kwargs):
        if not self.awc_no:
            try:
                if self.sku_recipe and self.sku_recipe.awc_no:
                    self.awc_no = self.sku_recipe.awc_no
                elif self.planning_job and self.planning_job.awc_no_display:
                    self.awc_no = self.planning_job.awc_no_display
            except Exception:
                pass

        became_available = self.status == self.STATUS_AVAILABLE
        super().save(*args, **kwargs)

        # Auto-transition the PlanningJob stage to 'plate_received' when plates are
        # ready/available. Runs AFTER the row is persisted: releasing member job
        # cards re-checks this plate request's status from the DB, so it must
        # already read AVAILABLE (not the stale RECEIVED value) or release wrongly
        # blocks itself as "an open plate request is still in progress".
        if became_available:
            try:
                planning_job = self.planning_job
                if planning_job and planning_job.planning_stage in ['new_plate_making', 'repeat_plate_making']:
                    planning_job.planning_stage = 'plate_received'
                    planning_job.save(update_fields=['planning_stage', 'updated_at'])
                # Shared plates for a merge group land all member jobs at once.
                if planning_job:
                    merge_group = planning_job.active_merge_group
                    if merge_group and merge_group.lead_job_id == planning_job.id:
                        merge_group.propagate_planning_stage('plate_received')
                        from printing_plates.services import release_merge_group_to_production

                        release_merge_group_to_production(merge_group, actor=self.received_by)
                    # If a plate set for this SKU was parked for reuse, this run
                    # is the one that picks it back up.
                    if not self.retained_for_reuse:
                        from printing_plates.services import release_retained_plate_for_sku

                        release_retained_plate_for_sku(planning_job.sku, actor=self.received_by)
            except Exception:
                pass


    @property
    def job_name(self):
        if self.sku_recipe and self.sku_recipe.job_name:
            return self.sku_recipe.job_name
        if self.job_card and self.job_card.planning_job and self.job_card.planning_job.job_name:
            return self.job_card.planning_job.job_name
        if self.planning_job and self.planning_job.job_name:
            return self.planning_job.job_name
        return ''

    @property
    def material_name(self):
        if self.sku_recipe and self.sku_recipe.material:
            return self.sku_recipe.material
        if self.job_card and self.job_card.material:
            return str(self.job_card.material)
        if self.planning_job and self.planning_job.material:
            return self.planning_job.material
        return ''

    @property
    def print_color_display(self):
        from printing_plates.constants import is_plate_ink_spec

        for candidate in (
            (self.planning_job.color_spec if self.planning_job else ''),
            (self.sku_recipe.color_spec if self.sku_recipe else ''),
            (self.job_card.colour if self.job_card else ''),
        ):
            value = (candidate or '').strip()
            if value and not is_plate_ink_spec(value):
                return value
        return ''

    @property
    def no_of_colors(self):
        from core.print_colors import print_color_total_units

        if self.job_card and self.job_card.total_colors is not None:
            return self.job_card.total_colors
        if self.planning_job and self.planning_job.number_of_colors is not None:
            return self.planning_job.number_of_colors
        units = print_color_total_units(self.print_color_display)
        return units or ''

    @property
    def impressions(self):
        if self.job_card and self.job_card.total_impressions_required is not None:
            return self.job_card.total_impressions_required
        if self.planning_job and self.planning_job.planned_total_impressions is not None:
            return self.planning_job.planned_total_impressions
        return None

    @property
    def plate_quantity_display(self):
        from printing_plates.plate_set_helpers import format_plate_quantity_display

        return format_plate_quantity_display(self.plate_quantity, self.sets_required)

    @property
    def jc_number(self):
        if self.job_card and self.job_card.job_card_no:
            return self.job_card.job_card_no
        if self.planning_job and self.planning_job.jc_number:
            return self.planning_job.jc_number
        return ''

    @property
    def sku(self):
        if self.sku_recipe and self.sku_recipe.sku:
            return self.sku_recipe.sku
        if self.job_card and self.job_card.SKU:
            return self.job_card.SKU
        if self.planning_job and self.planning_job.sku:
            return self.planning_job.sku
        return ''

    @property
    def machine_display(self):
        if self.machine:
            return str(self.machine)
        if self.job_card and self.job_card.machine_name:
            return str(self.job_card.machine_name)
        if self.planning_job and self.planning_job.machine_name:
            return self.planning_job.machine_name
        return ''

    @property
    def department_display(self):
        if self.department:
            return str(self.department)
        if self.job_card and self.job_card.department:
            return str(self.job_card.department)
        if self.planning_job and self.planning_job.department:
            return self.planning_job.department
        return ''

    @property
    def plate_request_type(self):
        from .services import plate_request_type_label, resolve_plate_request_type_key

        type_key = resolve_plate_request_type_key(self)
        if type_key == 'replacement':
            return 'Replacement'
        label = plate_request_type_label(type_key)
        if label == 'New':
            return 'New Artwork'
        return label

    @property
    def is_replacement(self):
        return (self.source or '') in {
            self.SOURCE_REPLACEMENT,
            self.SOURCE_PRODUCTION_PLATE_DAMAGE,
        } or bool(self.replacement_reason)

    @property
    def display_set_no(self):
        value = (self.set_no or self.new_set_no or '').strip()
        if value:
            return value
        if self.planning_job and (self.planning_job.plate_set_no or '').strip():
            return self.planning_job.plate_set_no.strip()
        if self.sku_recipe and (self.sku_recipe.plate_set_no or '').strip():
            return self.sku_recipe.plate_set_no.strip()
        if self.job_card and (self.job_card.plate_set_no or '').strip():
            return self.job_card.plate_set_no.strip()
        return ''

    @property
    def merge_info(self):
        """Smart-merge context when this request covers a combined layout.

        None for ordinary requests, so plate templates stay unchanged for the
        normal single-SKU flow.
        """
        if hasattr(self, '_cached_merge_info'):
            return self._cached_merge_info
        self._cached_merge_info = None
        if self.planning_job_id:
            from planning.services import build_job_card_merge_context

            self._cached_merge_info = build_job_card_merge_context(self.planning_job)
        return self._cached_merge_info

    @property
    def display_awc_no(self):
        from planning.services import normalize_awc_no

        value = normalize_awc_no(self.awc_no)
        if value:
            return value
        if self.planning_job:
            return (self.planning_job.awc_no_display or '').strip()
        return ''

    @property
    def display_die_cutting(self):
        """Die cutting is Yes/No on the plate; fall back to SKU master only."""
        from planning.services import normalize_die_cutting

        value = normalize_die_cutting(self.die_cutting)
        if value:
            return value
        if self.sku_recipe:
            return normalize_die_cutting(self.sku_recipe.die_cutting)
        return ''

    @property
    def is_open(self):
        return self.status in self.OPEN_STATUSES

    @property
    def is_cancelled(self):
        progress = (self.progress or '').strip().lower()
        remarks = (self.remarks or '').lower()
        return progress.startswith('cancelled') or 'cancelled — plates not required' in remarks

    @property
    def status_label_display(self):
        if self.is_cancelled:
            return 'Cancelled'
        return self.get_status_display()

    @property
    def cancel_reason_display(self):
        if not self.is_cancelled:
            return ''
        remarks = (self.remarks or '').strip()
        for line in remarks.splitlines():
            lowered = line.lower()
            if 'plates not required' in lowered:
                return line.split(':', 1)[-1].strip() or line.strip()
        return remarks

    @property
    def replacement_reason_display(self):
        return self.get_replacement_reason_display() if self.replacement_reason else ''

    def get_next_notification_roles(self):
        status = (self.status or '').strip().lower()
        if status == self.STATUS_DRAFT:
            return ['graphics_designer']
        elif status == self.STATUS_SENT:
            return ['admin', 'manager']
        elif status == self.STATUS_RECEIVED:
            return ['graphics_designer']
        elif status == self.STATUS_AVAILABLE:
            return ['production']
        return []


