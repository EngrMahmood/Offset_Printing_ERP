from django.db import models
from django.utils import timezone


class SheetsSyncSetting(models.Model):
    enabled = models.BooleanField(
        default=False,
        help_text="Turn on to start mirroring data to the configured Google Spreadsheet.",
    )
    spreadsheet_id = models.CharField(
        max_length=100, blank=True,
        help_text="The ID from the spreadsheet's URL (…/spreadsheets/d/<this>/edit).",
    )
    service_account_json_path = models.CharField(
        max_length=255, blank=True,
        help_text=(
            "Path to the service-account credentials JSON file. Leave blank to use the "
            "GOOGLE_SERVICE_ACCOUNT_JSON / GOOGLE_APPLICATION_CREDENTIALS env vars or the "
            "default gspread credentials location."
        ),
    )
    flush_interval_seconds = models.PositiveIntegerField(
        default=4, help_text="How often (seconds) queued changes are pushed to Sheets.",
    )
    max_batch_size = models.PositiveIntegerField(
        default=500, help_text="Maximum queued changes drained per flush cycle.",
    )
    circuit_breaker_failure_threshold = models.PositiveIntegerField(
        default=5, help_text="Consecutive failed flushes before pausing sync temporarily.",
    )
    circuit_breaker_cooldown_seconds = models.PositiveIntegerField(
        default=300, help_text="How long sync stays paused after the circuit breaker opens.",
    )
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Sheets Sync Setting"
        verbose_name_plural = "Sheets Sync Settings"

    @classmethod
    def get_settings(cls):
        obj, _ = cls.objects.get_or_create(id=1)
        return obj

    def __str__(self):
        return f"Sheets Sync Settings (Enabled: {self.enabled})"


class SheetsRowIndex(models.Model):
    """Persisted map of (tab, record pk) -> spreadsheet row number.

    Sheets has no native primary-key concept, so this table is what lets the
    writer upsert a record in place instead of appending a duplicate row on
    every update. It must survive process restarts, which rules out an
    in-memory dict.
    """
    tab_name = models.CharField(max_length=100, db_index=True)
    object_pk = models.CharField(max_length=64)
    row_number = models.PositiveIntegerField(help_text="1-based data row (header excluded).")
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Sheets Row Index"
        verbose_name_plural = "Sheets Row Indexes"
        unique_together = ('tab_name', 'object_pk')
        indexes = [models.Index(fields=['tab_name', 'object_pk'])]

    def __str__(self):
        return f"{self.tab_name}#{self.object_pk} -> row {self.row_number}"


class SheetsSyncLog(models.Model):
    STATUS_CHOICES = [
        ('SUCCESS', 'Success'),
        ('FAILED', 'Failed'),
        ('SKIPPED', 'Skipped'),
    ]

    timestamp = models.DateTimeField(default=timezone.now, db_index=True)
    tab_name = models.CharField(max_length=100, blank=True)
    batch_size = models.PositiveIntegerField(default=0)
    status = models.CharField(max_length=10, choices=STATUS_CHOICES)
    error_message = models.TextField(blank=True)
    duration_ms = models.PositiveIntegerField(null=True, blank=True)

    class Meta:
        verbose_name = "Sheets Sync Log"
        verbose_name_plural = "Sheets Sync Logs"
        ordering = ['-timestamp']

    def __str__(self):
        return f"{self.tab_name or 'ALL'} - {self.status} @ {self.timestamp}"
