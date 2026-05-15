import logging

import os

from django.contrib import messages
from django.contrib.auth.mixins import LoginRequiredMixin, PermissionRequiredMixin
from django.db.models import Count
from django.core.paginator import EmptyPage, PageNotAnInteger, Paginator
from django.shortcuts import get_object_or_404, redirect
from django.urls import reverse, reverse_lazy
from django.views import View
from django.views.generic import FormView, ListView, TemplateView
from urllib.parse import urlencode

from .forms import GoogleSheetComparisonForm, GoogleSheetUploadForm
from .models import (
    ComparisonJob,
    ComparisonModule,
    ComparisonStatus,
    ImportModule,
    JobStatus,
    MigrationImportJob,
    PlanningImportStaging,
    RowImportStatus,
)
from .services.comparison_engine import compare_sheet_to_erp
from .services.field_matcher import map_row_fields
from .services.google_reader import (
    credentials_to_dict,
    get_google_credential_status,
    get_google_oauth_flow,
    has_google_sheet_access,
    is_google_oauth_configured,
    read_google_sheet,
)
from .services.importer import (
    cleanup_imported_planning_job,
    import_planning_job,
    rollback_imported_planning_jobs,
)
from .services.validators import _parse_date, _parse_int, validate_planning_rows

logger = logging.getLogger(__name__)


def _pick_value(raw_row, keys, default=''):
    for key in keys:
        if key in raw_row and raw_row[key] not in (None, ''):
            return raw_row[key]
    return default


class MigrationDashboardView(LoginRequiredMixin, PermissionRequiredMixin, TemplateView):
    template_name = 'migration/dashboard.html'
    permission_required = 'migration.view_import'

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        jobs = MigrationImportJob.objects.all()
        rows = PlanningImportStaging.objects.all()

        context.update(
            {
                'total_imports': jobs.count(),
                'pending_rows': rows.filter(import_status=RowImportStatus.PENDING).count(),
                'valid_rows': rows.filter(import_status=RowImportStatus.VALID).count(),
                'error_rows': rows.filter(import_status=RowImportStatus.ERROR).count(),
                'imported_rows': rows.filter(import_status=RowImportStatus.IMPORTED).count(),
                'recent_jobs': jobs[:10],
                'credential_status': get_google_credential_status(),
            }
        )
        return context


class MigrationDeleteImportJobView(LoginRequiredMixin, PermissionRequiredMixin, View):
    permission_required = 'migration.run_import'

    def post(self, request, *args, **kwargs):
        job = get_object_or_404(MigrationImportJob, id=kwargs['job_id'])
        job_id = job.id
        job.delete()
        messages.success(request, f'Import job #{job_id} and its staging rows were deleted.')
        return redirect(reverse('migration:dashboard'))


class MigrationUploadView(LoginRequiredMixin, PermissionRequiredMixin, FormView):
    template_name = 'migration/upload.html'
    form_class = GoogleSheetUploadForm
    permission_required = 'migration.run_import'
    success_url = reverse_lazy('migration:dashboard')

    def get(self, request, *args, **kwargs):
        oauth_token = request.session.get('google_oauth_token')
        credential_status = get_google_credential_status()
        if is_google_oauth_configured() and not oauth_token and not credential_status['available']:
            return redirect('migration:google_auth_init')
        return super().get(request, *args, **kwargs)

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        credential_status = get_google_credential_status()
        oauth_token = self.request.session.get('google_oauth_token')
        if oauth_token:
            credential_status = {
                'available': True,
                'method': 'google_oauth',
                'message': 'Google OAuth credentials are available.',
            }

        context.update(
            {
                'credential_status': credential_status,
                'oauth_authorize_url': reverse('migration:google_auth_init'),
                'oauth_enabled': is_google_oauth_configured(),
            }
        )
        return context

    def form_valid(self, form):
        module = form.cleaned_data['module']
        sheet_url = form.cleaned_data['sheet_url']
        oauth_token = self.request.session.get('google_oauth_token')
        credential_status = get_google_credential_status()

        if is_google_oauth_configured() and not oauth_token and not credential_status['available']:
            return redirect('migration:google_auth_init')

        if module != ImportModule.PLANNING:
            messages.error(self.request, 'Phase 1 supports Planning import only. Production/Dispatch will be added later.')
            return redirect('migration:upload')

        try:
            rows = read_google_sheet(sheet_url, oauth_token=oauth_token)
            if oauth_token:
                self.request.session['google_oauth_token'] = oauth_token
                self.request.session.modified = True
        except Exception as exc:
            logger.exception('Sheet read failed: %s', sheet_url)
            messages.error(self.request, f'Could not read Google Sheet: {exc}')
            return redirect('migration:upload')

        if not rows:
            messages.warning(self.request, 'No data rows found in the selected sheet.')
            return redirect('migration:upload')

        import_job = MigrationImportJob.objects.create(
            module=module,
            sheet_url=sheet_url,
            status=JobStatus.STAGED,
            created_by=self.request.user,
            total_rows=len(rows),
        )

        staging_rows = []
        for idx, raw in enumerate(rows, start=1):
            mapped_raw = map_row_fields(raw)
            po_number = _pick_value(mapped_raw, ['po_number', 'po', 'po no', 'po_no'])
            customer = _pick_value(mapped_raw, ['customer', 'customer_name', 'destination'])
            sku = _pick_value(mapped_raw, ['sku', 'item_code', 'product_code'])
            quantity_raw = _pick_value(mapped_raw, ['qty', 'quantity', 'order_qty'])
            delivery_raw = _pick_value(mapped_raw, ['delivery_date', 'delivery', 'date', 'plan_date', 'po_date', 'po_received_date', 'po_approval_date'])

            staging_rows.append(
                PlanningImportStaging(
                    import_job=import_job,
                    row_number=idx,
                    po_number=str(po_number or '').strip(),
                    customer=str(customer or '').strip(),
                    sku=str(sku or '').strip(),
                    quantity=_parse_int(quantity_raw),
                    delivery_date=_parse_date(delivery_raw),
                    raw_data=mapped_raw,
                )
            )

        PlanningImportStaging.objects.bulk_create(staging_rows, batch_size=10)

        validation_result = validate_planning_rows(import_job)
        import_job.valid_rows = validation_result['valid']
        import_job.error_rows = validation_result['errors']
        import_job.status = JobStatus.VALIDATED
        import_job.save(update_fields=['valid_rows', 'error_rows', 'status', 'updated_at'])

        messages.success(
            self.request,
            f"Sheet staged successfully. Total: {validation_result['total']}, Valid: {validation_result['valid']}, Errors: {validation_result['errors']}.",
        )
        return redirect(reverse('migration:preview', kwargs={'job_id': import_job.id}))


class ComparisonUploadView(LoginRequiredMixin, PermissionRequiredMixin, FormView):
    template_name = 'migration/comparison_upload.html'
    form_class = GoogleSheetComparisonForm
    permission_required = 'migration.run_import'
    success_url = reverse_lazy('migration:dashboard')

    def get(self, request, *args, **kwargs):
        oauth_token = request.session.get('google_oauth_token')
        credential_status = get_google_credential_status()
        if is_google_oauth_configured() and not oauth_token and not credential_status['available']:
            return redirect('migration:google_auth_init')
        return super().get(request, *args, **kwargs)

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        credential_status = get_google_credential_status()
        oauth_token = self.request.session.get('google_oauth_token')
        if oauth_token:
            credential_status = {
                'available': True,
                'method': 'google_oauth',
                'message': 'Google OAuth credentials are available.',
            }

        context.update(
            {
                'credential_status': credential_status,
                'oauth_authorize_url': reverse('migration:google_auth_init'),
                'oauth_enabled': is_google_oauth_configured(),
            }
        )
        return context

    def form_valid(self, form):
        module = form.cleaned_data['module']
        sheet_url = form.cleaned_data['sheet_url']
        oauth_token = self.request.session.get('google_oauth_token')
        credential_status = get_google_credential_status()

        if is_google_oauth_configured() and not oauth_token and not credential_status['available']:
            return redirect('migration:google_auth_init')

        try:
            comparison_job, missing_fields, match_results, erp_schema, metadata = compare_sheet_to_erp(
                sheet_url,
                module,
                oauth_token=oauth_token,
                user=self.request.user,
            )
            if oauth_token:
                self.request.session['google_oauth_token'] = oauth_token
                self.request.session.modified = True
        except Exception as exc:
            logger.exception('Comparison engine failed: %s', sheet_url)
            messages.error(self.request, f'Could not compare Google Sheet: {exc}')
            return redirect('migration:compare')

        messages.success(
            self.request,
            f"Comparison completed. {comparison_job.matched_columns}/{comparison_job.total_columns} columns matched, "
            f"{comparison_job.missing_columns} required fields missing.",
        )
        return redirect(reverse('migration:comparison_result', kwargs={'job_id': comparison_job.id}))


class ComparisonResultView(LoginRequiredMixin, PermissionRequiredMixin, TemplateView):
    template_name = 'migration/comparison_report.html'
    permission_required = 'migration.view_import'

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        comparison_job = get_object_or_404(ComparisonJob, id=self.kwargs['job_id'])
        results = comparison_job.results.all()

        context.update(
            {
                'comparison_job': comparison_job,
                'results': results,
                'missing_fields': comparison_job.missing_columns,
                'extra_columns': comparison_job.extra_columns,
            }
        )
        return context


class MigrationGoogleAuthInitView(LoginRequiredMixin, PermissionRequiredMixin, View):
    permission_required = 'migration.run_import'

    def get(self, request, *args, **kwargs):
        redirect_uri = request.build_absolute_uri(reverse('migration:google_auth_callback'))
        try:
            flow = get_google_oauth_flow(request, redirect_uri)
        except Exception as exc:
            messages.error(request, f'Google OAuth configuration error: {exc}')
            return redirect('migration:upload')

        authorization_url, state = flow.authorization_url(
            access_type='offline',
            include_granted_scopes='true',
            prompt='consent',
        )
        request.session['google_oauth_state'] = state
        request.session['google_oauth_redirect_uri'] = redirect_uri
        return redirect(authorization_url)


class MigrationGoogleAuthCallbackView(LoginRequiredMixin, PermissionRequiredMixin, View):
    permission_required = 'migration.run_import'

    def get(self, request, *args, **kwargs):
        error = request.GET.get('error')
        if error:
            messages.error(request, f'Google OAuth denied: {error}')
            return redirect('migration:upload')

        state = request.GET.get('state')
        expected_state = request.session.get('google_oauth_state')
        if not state or state != expected_state:
            messages.error(request, 'Google OAuth state mismatch. Please try again.')
            return redirect('migration:upload')

        code = request.GET.get('code')
        if not code:
            messages.error(request, 'Google OAuth callback did not return a code.')
            return redirect('migration:upload')

        redirect_uri = request.session.get('google_oauth_redirect_uri') or request.build_absolute_uri(reverse('migration:google_auth_callback'))
        try:
            flow = get_google_oauth_flow(request, redirect_uri)
            flow.fetch_token(code=code)
        except Exception as exc:
            messages.error(request, f'Google OAuth token fetch failed: {exc}')
            return redirect('migration:upload')

        request.session['google_oauth_token'] = credentials_to_dict(flow.credentials)
        request.session.modified = True
        messages.success(request, 'Google authorization successful. You may now upload the sheet.')
        return redirect('migration:upload')


class MigrationPreviewView(LoginRequiredMixin, PermissionRequiredMixin, TemplateView):
    template_name = 'migration/preview.html'
    permission_required = 'migration.view_import'

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        import_job = get_object_or_404(MigrationImportJob, id=self.kwargs['job_id'])
        status_filter = self.request.GET.get('status', 'ALL').upper()
        rows_qs = import_job.planning_rows.all().order_by('row_number', 'id')
        valid_statuses = {
            'ALL': None,
            'VALID': RowImportStatus.VALID,
            'ERROR': RowImportStatus.ERROR,
            'PENDING': RowImportStatus.PENDING,
            'IMPORTED': RowImportStatus.IMPORTED,
        }

        if status_filter in valid_statuses and valid_statuses[status_filter]:
            rows_qs = rows_qs.filter(import_status=valid_statuses[status_filter])
        elif status_filter not in valid_statuses:
            status_filter = 'ALL'

        try:
            page_size = int(self.request.GET.get('page_size', 50))
        except (TypeError, ValueError):
            page_size = 50
        if page_size not in (50, 100):
            page_size = 50

        paginator = Paginator(rows_qs, page_size)
        page_number = self.request.GET.get('page', 1)
        try:
            page_obj = paginator.page(page_number)
        except PageNotAnInteger:
            page_obj = paginator.page(1)
        except EmptyPage:
            page_obj = paginator.page(paginator.num_pages)

        context.update(
            {
                'import_job': import_job,
                'rows': page_obj.object_list,
                'page_obj': page_obj,
                'page_size': page_size,
                'page_sizes': [50, 100],
                'status_filter': status_filter,
                'status_options': ['ALL', 'VALID', 'ERROR', 'PENDING', 'IMPORTED'],
                'summary': {
                    'total': import_job.planning_rows.count(),
                    'valid': import_job.planning_rows.filter(import_status=RowImportStatus.VALID).count(),
                    'error': import_job.planning_rows.filter(import_status=RowImportStatus.ERROR).count(),
                    'imported': import_job.planning_rows.filter(import_status=RowImportStatus.IMPORTED).count(),
                    'pending': import_job.planning_rows.filter(import_status=RowImportStatus.PENDING).count(),
                },
            }
        )
        return context

    def _redirect_to_preview(self, import_job):
        params = {}
        status = self.request.POST.get('status') or self.request.GET.get('status')
        page_size = self.request.POST.get('page_size') or self.request.GET.get('page_size')
        page = self.request.POST.get('page') or self.request.GET.get('page')
        if status:
            params['status'] = status
        if page_size in ('50', '100'):
            params['page_size'] = page_size
        if page and page.isdigit():
            params['page'] = page
        url = reverse('migration:preview', kwargs={'job_id': import_job.id})
        if params:
            url = f"{url}?{urlencode(params)}"
        return redirect(url)

    def post(self, request, *args, **kwargs):
        import_job = get_object_or_404(MigrationImportJob, id=kwargs['job_id'])
        action = request.POST.get('action', '')

        if action.startswith('delete_row:'):
            try:
                row_id = int(action.split(':', 1)[1])
            except (TypeError, ValueError, IndexError):
                messages.error(request, 'Invalid row selected for deletion.')
                return self._redirect_to_preview(import_job)

            row = get_object_or_404(PlanningImportStaging, id=row_id, import_job=import_job)
            row.delete()
            messages.success(request, f'Row {row.row_number} was deleted from staging.')
            return self._redirect_to_preview(import_job)

        if action == 'cleanup_import':
            if import_job.module != ImportModule.PLANNING:
                messages.error(request, 'Cleanup is supported only for Planning import jobs.')
                return self._redirect_to_preview(import_job)

            cleanup_result = cleanup_imported_planning_job(import_job)
            if cleanup_result['deleted']:
                messages.success(
                    request,
                    f"Clean up complete. Deleted {cleanup_result['deleted']} imported PlanningJob(s) and reset {cleanup_result['reset']} staging row(s) to valid.",
                )
            else:
                messages.success(
                    request,
                    f"Clean up complete. Reset {cleanup_result['reset']} imported staging rows to valid. No PlanningJob records were deleted.",
                )
            return redirect(reverse('migration:preview', kwargs={'job_id': import_job.id}))

        if action == 'delete_selected_rows':
            selected_ids = [int(pk) for pk in request.POST.getlist('selected_row_ids') if pk.isdigit()]
            if not selected_ids:
                messages.error(request, 'Select at least one row to delete.')
                return self._redirect_to_preview(import_job)

            deleted_rows = PlanningImportStaging.objects.filter(id__in=selected_ids, import_job=import_job)
            deleted_count = deleted_rows.count()
            deleted_rows.delete()
            messages.success(request, f'Deleted {deleted_count} selected staging row(s).')
            return self._redirect_to_preview(import_job)

        messages.error(request, 'Unknown action.')
        return self._redirect_to_preview(import_job)


class MigrationPreviewRowDetailView(LoginRequiredMixin, PermissionRequiredMixin, TemplateView):
    template_name = 'migration/row_detail.html'
    permission_required = 'migration.view_import'

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        import_job = get_object_or_404(MigrationImportJob, id=self.kwargs['job_id'])
        row = get_object_or_404(PlanningImportStaging, id=self.kwargs['row_id'], import_job=import_job)

        context.update(
            {
                'import_job': import_job,
                'row': row,
                'raw_data_items': sorted((row.raw_data or {}).items()),
            }
        )
        return context


class MigrationRunImportView(LoginRequiredMixin, PermissionRequiredMixin, View):
    permission_required = 'migration.run_import'

    def post(self, request, *args, **kwargs):
        import_job = get_object_or_404(MigrationImportJob, id=kwargs['job_id'])

        if import_job.module != ImportModule.PLANNING:
            messages.error(request, 'Phase 1 import supports Planning module only.')
            return redirect(reverse('migration:preview', kwargs={'job_id': import_job.id}))

        if import_job.status == JobStatus.COMPLETED:
            messages.info(request, 'This import job has already completed. There are no valid rows left to import.')
            return redirect(reverse('migration:preview', kwargs={'job_id': import_job.id}))

        comparison_exists = ComparisonJob.objects.filter(
            module=import_job.module,
            sheet_url=import_job.sheet_url,
            status__in=[ComparisonStatus.COMPLETED, ComparisonStatus.REVIEW],
        ).exists()
        if not comparison_exists:
            messages.error(
                request,
                'A comparison must be completed for this sheet before importing. Run the compare flow first.',
            )
            return redirect(reverse('migration:preview', kwargs={'job_id': import_job.id}))

        result = import_planning_job(import_job, request.user)
        messages.success(
            request,
            f"Import completed. Success: {result['success']}, Errors: {result['errors']}, Imported rows: {result['imported']}",
        )
        return redirect(reverse('migration:preview', kwargs={'job_id': import_job.id}))


class MigrationRollbackImportView(LoginRequiredMixin, PermissionRequiredMixin, View):
    permission_required = 'migration.run_import'

    def post(self, request, *args, **kwargs):
        import_job = get_object_or_404(MigrationImportJob, id=kwargs['job_id'])
        if import_job.module != ImportModule.PLANNING:
            messages.error(request, 'Rollback is supported only for Planning import jobs.')
            return redirect(reverse('migration:preview', kwargs={'job_id': import_job.id}))

        if import_job.planning_rows.filter(import_status=RowImportStatus.IMPORTED).count() == 0:
            messages.warning(request, 'No imported rows found to rollback for this import job.')
            return redirect(reverse('migration:preview', kwargs={'job_id': import_job.id}))

        deleted_count = rollback_imported_planning_jobs(import_job)
        if deleted_count:
            messages.success(request, f'Rolled back {deleted_count} imported planning job(s).')
        else:
            messages.warning(request, 'No imported planning jobs could be rolled back.')
        return redirect(reverse('migration:preview', kwargs={'job_id': import_job.id}))


class MigrationClearImportsView(LoginRequiredMixin, PermissionRequiredMixin, TemplateView):
    template_name = 'migration/clear_imports.html'
    permission_required = 'migration.run_import'

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        import_jobs = MigrationImportJob.objects.annotate(staging_rows_count=Count('planning_rows')).order_by('-created_at')
        context.update(
            {
                'import_jobs': import_jobs,
                'import_jobs_count': import_jobs.count(),
                'staging_rows_count': PlanningImportStaging.objects.count(),
            }
        )
        return context

    def post(self, request, *args, **kwargs):
        if request.POST.get('confirm') != '1':
            messages.error(request, 'Bulk delete was not confirmed.')
            return redirect(reverse('migration:clear_imports'))

        action = request.POST.get('action')
        if action == 'delete_selected':
            selected_ids = [int(pk) for pk in request.POST.getlist('selected_job_ids') if pk.isdigit()]
            if not selected_ids:
                messages.error(request, 'Select at least one import job to delete.')
                return redirect(reverse('migration:clear_imports'))

            deleted_jobs = MigrationImportJob.objects.filter(id__in=selected_ids)
            deleted_job_count = deleted_jobs.count()
            deleted_row_count = PlanningImportStaging.objects.filter(import_job_id__in=selected_ids).count()
            deleted_jobs.delete()
            messages.success(
                request,
                f'Deleted {deleted_job_count} selected import job(s) and {deleted_row_count} staging row(s).',
            )
            return redirect(reverse('migration:clear_imports'))

        if action == 'cleanup_selected':
            selected_ids = [int(pk) for pk in request.POST.getlist('selected_job_ids') if pk.isdigit()]
            if not selected_ids:
                messages.error(request, 'Select at least one import job to clean up.')
                return redirect(reverse('migration:clear_imports'))

            total_deleted = 0
            total_reset = 0
            for cleanup_job in MigrationImportJob.objects.filter(id__in=selected_ids, module=ImportModule.PLANNING):
                result = cleanup_imported_planning_job(cleanup_job)
                total_deleted += result['deleted']
                total_reset += result['reset']

            if total_deleted or total_reset:
                messages.success(
                    request,
                    f'Cleanup complete. Deleted {total_deleted} imported PlanningJob(s) and reset {total_reset} imported staging row(s).',
                )
            else:
                messages.warning(request, 'No imported PlanningJob records were found to clean up for the selected jobs.')
            return redirect(reverse('migration:clear_imports'))

        if action == 'delete_all':
            import_jobs_count = MigrationImportJob.objects.count()
            staging_rows_count = PlanningImportStaging.objects.count()
            MigrationImportJob.objects.all().delete()
            messages.success(
                request,
                f'Bulk delete complete. Deleted {import_jobs_count} import job(s) and {staging_rows_count} staging row(s).',
            )
            return redirect(reverse('migration:dashboard'))

        messages.error(request, 'Unknown delete action.')
        return redirect(reverse('migration:clear_imports'))


class MigrationLogsView(LoginRequiredMixin, PermissionRequiredMixin, ListView):
    template_name = 'migration/logs.html'
    permission_required = 'migration.view_import'
    model = MigrationImportJob
    context_object_name = 'jobs'
    paginate_by = 30

    def get_queryset(self):
        return MigrationImportJob.objects.select_related('created_by').prefetch_related('logs').order_by('-created_at')
