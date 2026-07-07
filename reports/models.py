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
