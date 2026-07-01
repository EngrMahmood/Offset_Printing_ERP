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
    plate_quantity = models.PositiveIntegerField(null=True, blank=True)
    plate_color = models.CharField(max_length=120, blank=True)
    vendor = models.CharField(max_length=120, blank=True)
    remarks = models.TextField(blank=True)
    source = models.CharField(max_length=120, blank=True)
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

    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['-requested_at', '-created_at']
        indexes = [
            models.Index(fields=['planning_job']),
            models.Index(fields=['status']),
            models.Index(fields=['requested_at']),
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

        # Auto-transition the PlanningJob stage to 'plate_received' when plates are ready/available
        if self.status == self.STATUS_AVAILABLE:
            try:
                planning_job = self.planning_job
                if planning_job and planning_job.planning_stage in ['new_plate_making', 'repeat_plate_making']:
                    planning_job.planning_stage = 'plate_received'
                    planning_job.save(update_fields=['planning_stage', 'updated_at'])
            except Exception:
                pass
        super().save(*args, **kwargs)


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
    def no_of_colors(self):
        if self.job_card and self.job_card.total_colors is not None:
            return self.job_card.total_colors
        if self.planning_job and self.planning_job.total_colors is not None:
            return self.planning_job.total_colors
        if self.planning_job and self.planning_job.color_spec:
            return self.planning_job.color_spec
        return ''

    @property
    def impressions(self):
        if self.job_card and self.job_card.total_impressions_required is not None:
            return self.job_card.total_impressions_required
        if self.planning_job and self.planning_job.planned_total_impressions is not None:
            return self.planning_job.planned_total_impressions
        return None

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
        flag = ''
        if self.planning_job:
            flag = self.planning_job.repeat_flag
        elif self.job_card and self.job_card.planning_job:
            flag = self.job_card.planning_job.repeat_flag

        if flag == 'New':
            return 'New Artwork'
        elif flag == 'Repeat':
            return 'Repeat'
        return ''


