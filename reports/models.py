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
