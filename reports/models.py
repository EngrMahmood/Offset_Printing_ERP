from django.conf import settings
from django.db import models


class ScheduledReport(models.Model):
    FREQ_DAILY = 'daily'
    FREQ_WEEKLY = 'weekly'
    FREQ_MONTHLY = 'monthly'
    FREQ_QUARTERLY = 'quarterly'

    FREQUENCY_CHOICES = [
        (FREQ_DAILY, 'Daily'),
        (FREQ_WEEKLY, 'Weekly'),
        (FREQ_MONTHLY, 'Monthly'),
        (FREQ_QUARTERLY, 'Quarterly'),
    ]

    name = models.CharField(max_length=120)
    report_slug = models.CharField(max_length=120, db_index=True)
    frequency = models.CharField(max_length=20, choices=FREQUENCY_CHOICES, default=FREQ_WEEKLY)
    recipients = models.TextField(blank=True)
    filters = models.JSONField(default=dict, blank=True)
    is_active = models.BooleanField(default=True)
    last_run_at = models.DateTimeField(null=True, blank=True)
    next_run_at = models.DateTimeField(null=True, blank=True, db_index=True)
    created_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='scheduled_reports_created',
    )
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['-created_at']

    def __str__(self):
        return f'{self.name} ({self.report_slug})'


class MachinePlanningJcSelection(models.Model):
    """Shared, planner/admin-editable opt-out for a JC in a Machine Planning
    combined run (V2 plan item 3). Everyone sees the same selection state;
    an excluded JC is dropped from the report's merged totals/exports but
    stays visible (marked excluded) in the planner console."""

    jc_number = models.CharField(max_length=50, unique=True, db_index=True)
    is_excluded = models.BooleanField(default=False)
    updated_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
    )
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return f'{self.jc_number} ({"excluded" if self.is_excluded else "included"})'


class KPITarget(models.Model):
    """One row per KPI per year — mirrors the yearly KPI scorecard workbook
    (Position/Objective/Title/Description/UoM/Weightage/Min/Target/Max), so
    targets can be adjusted for a new year via admin without a code change."""

    KPI_ORDER_FULFILLMENT = 'order_fulfillment'
    KPI_WASTAGE_REDUCTION = 'wastage_reduction'
    KPI_DISPATCH_ALIGNMENT = 'dispatch_alignment'

    KPI_CHOICES = [
        (KPI_ORDER_FULFILLMENT, 'Order Fulfillment Efficiency'),
        (KPI_WASTAGE_REDUCTION, 'Wastage Reduction Efficiency'),
        (KPI_DISPATCH_ALIGNMENT, 'Dispatch vs Production Alignment'),
    ]

    kpi_slug = models.CharField(max_length=40, choices=KPI_CHOICES)
    year = models.PositiveIntegerField()
    position = models.CharField(max_length=120, blank=True)
    objective = models.CharField(max_length=120, blank=True)
    title = models.CharField(max_length=160, blank=True)
    description = models.TextField(blank=True)
    uom = models.CharField(max_length=20, default='%')
    weightage_pct = models.DecimalField(max_digits=5, decimal_places=2, default=0)
    monitoring_frequency = models.CharField(max_length=20, default='Quarterly')
    min_value = models.DecimalField(max_digits=6, decimal_places=2)
    target_value = models.DecimalField(max_digits=6, decimal_places=2)
    max_value = models.DecimalField(max_digits=6, decimal_places=2)
    higher_is_better = models.BooleanField(default=True)

    class Meta:
        unique_together = ('kpi_slug', 'year')
        ordering = ['-year', 'kpi_slug']

    def __str__(self):
        return f'{self.get_kpi_slug_display()} ({self.year})'


class KPIActionNote(models.Model):
    """Manager-editable remark/action-plan saved against a KPI for a given
    month or quarter, so the auto-suggested action text can be reviewed and
    overridden before being kept as the record for that period."""

    PERIOD_MONTH = 'month'
    PERIOD_QUARTER = 'quarter'
    PERIOD_TYPE_CHOICES = [
        (PERIOD_MONTH, 'Month'),
        (PERIOD_QUARTER, 'Quarter'),
    ]

    kpi_slug = models.CharField(max_length=40, choices=KPITarget.KPI_CHOICES)
    period_type = models.CharField(max_length=10, choices=PERIOD_TYPE_CHOICES)
    period_key = models.CharField(max_length=10)  # e.g. '2026-07' or '2026-Q3'
    note = models.TextField(blank=True)
    status = models.CharField(max_length=10, blank=True)
    updated_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
    )
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        unique_together = ('kpi_slug', 'period_type', 'period_key')

    def __str__(self):
        return f'{self.kpi_slug} {self.period_key}'
