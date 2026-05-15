from django.contrib import admin

from .models import MigrationImportJob, MigrationImportLog, PlanningImportStaging, RowImportStatus
from .services.importer import import_planning_job


@admin.register(MigrationImportJob)
class MigrationImportJobAdmin(admin.ModelAdmin):
    list_display = (
        'id',
        'module',
        'status',
        'total_rows',
        'valid_rows',
        'error_rows',
        'imported_rows',
        'created_by',
        'created_at',
    )
    list_filter = ('module', 'status', 'created_at')
    search_fields = ('id', 'sheet_url', 'created_by__username')
    actions = ('delete_selected_import_jobs',)

    @admin.action(description='Delete selected import jobs and their staging rows')
    def delete_selected_import_jobs(self, request, queryset):
        count = queryset.count()
        queryset.delete()
        self.message_user(request, f'Deleted {count} import job(s) and related staging rows.')


@admin.register(PlanningImportStaging)
class PlanningImportStagingAdmin(admin.ModelAdmin):
    list_display = (
        'id',
        'import_job',
        'row_number',
        'po_number',
        'sku',
        'quantity',
        'import_status',
        'imported_reference',
        'updated_at',
    )
    list_filter = ('import_status', 'created_at')
    search_fields = ('po_number', 'sku', 'customer', 'import_job__id')
    actions = ('rerun_selected_rows_import',)

    @admin.action(description='Re-run import for selected rows (VALID/ERROR only)')
    def rerun_selected_rows_import(self, request, queryset):
        job_ids = queryset.values_list('import_job_id', flat=True).distinct()
        processed = 0
        for job_id in job_ids:
            job = MigrationImportJob.objects.filter(id=job_id).first()
            if not job:
                continue
            queryset.filter(import_job_id=job_id, import_status=RowImportStatus.ERROR).update(
                import_status=RowImportStatus.VALID,
                error_message='',
            )
            import_planning_job(job, request.user)
            processed += 1
        self.message_user(request, f'Re-ran imports for {processed} job(s).')


@admin.register(MigrationImportLog)
class MigrationImportLogAdmin(admin.ModelAdmin):
    list_display = (
        'id',
        'import_job',
        'imported_by',
        'rows_count',
        'success_count',
        'error_count',
        'created_at',
    )
    list_filter = ('created_at',)
    search_fields = ('import_job__id', 'imported_by__username', 'message')
