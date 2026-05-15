from django.conf import settings
from django.db import models


class ImportModule(models.TextChoices):
    PLANNING = 'PLANNING', 'Planning'
    PRODUCTION = 'PRODUCTION', 'Production'
    DISPATCH = 'DISPATCH', 'Dispatch'


class JobStatus(models.TextChoices):
    STAGED = 'STAGED', 'Staged'
    VALIDATED = 'VALIDATED', 'Validated'
    PARTIAL = 'PARTIAL', 'Partially Imported'
    COMPLETED = 'COMPLETED', 'Completed'
    FAILED = 'FAILED', 'Failed'


class RowImportStatus(models.TextChoices):
    PENDING = 'PENDING', 'Pending'
    VALID = 'VALID', 'Valid'
    ERROR = 'ERROR', 'Error'
    IMPORTED = 'IMPORTED', 'Imported'


class MigrationImportJob(models.Model):
    module = models.CharField(max_length=20, choices=ImportModule.choices)
    sheet_url = models.URLField(max_length=1024)
    status = models.CharField(max_length=20, choices=JobStatus.choices, default=JobStatus.STAGED)

    total_rows = models.PositiveIntegerField(default=0)
    valid_rows = models.PositiveIntegerField(default=0)
    error_rows = models.PositiveIntegerField(default=0)
    imported_rows = models.PositiveIntegerField(default=0)

    created_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='migration_jobs_created',
    )
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['-created_at']
        permissions = [
            ('view_import', 'Can view migration import module'),
            ('run_import', 'Can run migration imports'),
        ]

    def __str__(self):
        return f"{self.module} import #{self.pk}"


class PlanningImportStaging(models.Model):
    import_job = models.ForeignKey(
        MigrationImportJob,
        on_delete=models.CASCADE,
        related_name='planning_rows',
    )
    row_number = models.PositiveIntegerField()

    po_number = models.CharField(max_length=120, blank=True)
    customer = models.CharField(max_length=255, blank=True)
    sku = models.CharField(max_length=255, blank=True)
    quantity = models.PositiveIntegerField(null=True, blank=True)
    delivery_date = models.DateField(null=True, blank=True)

    raw_data = models.JSONField(default=dict, blank=True)
    import_status = models.CharField(
        max_length=20,
        choices=RowImportStatus.choices,
        default=RowImportStatus.PENDING,
    )
    error_message = models.TextField(blank=True, null=True)
    imported_reference = models.CharField(max_length=120, blank=True)

    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['row_number', 'id']
        unique_together = ('import_job', 'row_number')

    def __str__(self):
        return f"Job {self.import_job_id} row {self.row_number}"


class ComparisonModule(models.TextChoices):
    PO_INTAKE = 'PO_INTAKE', 'PO Intake'
    SKU_MASTER = 'SKU_MASTER', 'SKU Master'
    JOB_FINALIZE = 'JOB_FINALIZE', 'Job Finalize'
    PLANNING = 'PLANNING', 'Planning'
    PRODUCTION = 'PRODUCTION', 'Production'
    DISPATCH = 'DISPATCH', 'Dispatch'


class ComparisonStatus(models.TextChoices):
    PENDING = 'PENDING', 'Pending'
    COMPLETED = 'COMPLETED', 'Completed'
    REVIEW = 'REVIEW', 'Review Required'
    FAILED = 'FAILED', 'Failed'


class ComparisonJob(models.Model):
    module = models.CharField(max_length=30, choices=ComparisonModule.choices)
    sheet_url = models.URLField(max_length=1024)
    status = models.CharField(max_length=20, choices=ComparisonStatus.choices, default=ComparisonStatus.PENDING)

    total_columns = models.PositiveIntegerField(default=0)
    matched_columns = models.PositiveIntegerField(default=0)
    missing_columns = models.PositiveIntegerField(default=0)
    extra_columns = models.PositiveIntegerField(default=0)

    created_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='comparison_jobs_created',
    )
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['-created_at']

    def __str__(self):
        return f"Comparison job #{self.pk} ({self.get_module_display()})"


class ComparisonResult(models.Model):
    comparison_job = models.ForeignKey(
        ComparisonJob,
        on_delete=models.CASCADE,
        related_name='results',
    )
    sheet_column = models.CharField(max_length=255)
    erp_model = models.CharField(max_length=255, blank=True)
    erp_field = models.CharField(max_length=255, blank=True)
    match_type = models.CharField(max_length=30, blank=True)
    status = models.CharField(max_length=30, blank=True)
    confidence = models.DecimalField(max_digits=5, decimal_places=2, null=True, blank=True)
    details = models.TextField(blank=True)
    sample_values = models.JSONField(default=list, blank=True)

    class Meta:
        ordering = ['id']

    def __str__(self):
        return f"{self.sheet_column} -> {self.erp_model}.{self.erp_field or 'none'}"


class ColumnMapping(models.Model):
    comparison_job = models.ForeignKey(
        ComparisonJob,
        on_delete=models.CASCADE,
        related_name='column_mappings',
    )
    profile_name = models.CharField(max_length=120, blank=True)
    sheet_column = models.CharField(max_length=255)
    erp_model = models.CharField(max_length=255, blank=True)
    erp_field = models.CharField(max_length=255, blank=True)
    match_confidence = models.DecimalField(max_digits=5, decimal_places=2, default=0.0)
    is_confirmed = models.BooleanField(default=False)
    created_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='column_mappings_created',
    )
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['comparison_job', 'sheet_column']
        unique_together = ('comparison_job', 'sheet_column')

    def __str__(self):
        return f"{self.sheet_column} -> {self.erp_model}.{self.erp_field} ({'confirmed' if self.is_confirmed else 'pending'})"


class MigrationImportLog(models.Model):
    import_job = models.ForeignKey(
        MigrationImportJob,
        on_delete=models.CASCADE,
        related_name='logs',
    )
    imported_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='migration_import_logs',
    )
    rows_count = models.PositiveIntegerField(default=0)
    success_count = models.PositiveIntegerField(default=0)
    error_count = models.PositiveIntegerField(default=0)
    message = models.TextField(blank=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['-created_at']

    def __str__(self):
        return f"Import log #{self.pk} for job {self.import_job_id}"
