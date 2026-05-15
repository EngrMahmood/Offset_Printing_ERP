from django.urls import path

from .views import (
    ComparisonResultView,
    ComparisonUploadView,
    MigrationDashboardView,
    MigrationDeleteImportJobView,
    MigrationGoogleAuthCallbackView,
    MigrationGoogleAuthInitView,
    MigrationLogsView,
    MigrationPreviewView,
    MigrationPreviewRowDetailView,
    MigrationRunImportView,
    MigrationRollbackImportView,
    MigrationClearImportsView,
    MigrationUploadView,
)

app_name = 'migration'

urlpatterns = [
    path('', MigrationDashboardView.as_view(), name='dashboard'),
    path('upload/', MigrationUploadView.as_view(), name='upload'),
    path('compare/', ComparisonUploadView.as_view(), name='compare'),
    path('compare/<int:job_id>/', ComparisonResultView.as_view(), name='comparison_result'),
    path('google-auth/', MigrationGoogleAuthInitView.as_view(), name='google_auth_init'),
    path('google-auth/callback/', MigrationGoogleAuthCallbackView.as_view(), name='google_auth_callback'),
    path('preview/<int:job_id>/row/<int:row_id>/', MigrationPreviewRowDetailView.as_view(), name='preview_row_detail'),
    path('preview/<int:job_id>/', MigrationPreviewView.as_view(), name='preview'),
    path('import/<int:job_id>/', MigrationRunImportView.as_view(), name='run_import'),
    path('rollback/<int:job_id>/', MigrationRollbackImportView.as_view(), name='rollback_import'),
    path('delete/<int:job_id>/', MigrationDeleteImportJobView.as_view(), name='delete_job'),
    path('clear/', MigrationClearImportsView.as_view(), name='clear_imports'),
    path('logs/', MigrationLogsView.as_view(), name='logs'),
]
