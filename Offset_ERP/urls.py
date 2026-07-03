from django.contrib import admin
from django.urls import path, include
from django.contrib.auth import views as auth_views
from core.views import home
from core import views
from core import notification_views

urlpatterns = [
    path('admin/', admin.site.urls),
    path('login/', auth_views.LoginView.as_view(template_name='login.html'), name='login'),
    path('logout/', auth_views.LogoutView.as_view(next_page='login'), name='logout'),
    path('', home, name='home'),
    path('planning/', include('planning.urls', namespace='planning')),
    path('printing-plates/', include('printing_plates.urls', namespace='printing_plates')),
    path('reports/', include('reports.urls', namespace='reports')),
    path('migration/', include('migration.urls', namespace='migration')),
    path('qc/', include('qc.urls', namespace='qc')),
    path('manual-working/', include('manual_working.urls', namespace='manual_working')),
    path('supply-chain/', include('supply_chain.urls', namespace='supply_chain')),

    path('bulk-upload-jobcards/', views.bulk_upload_jobcards, name='bulk_upload_jobcards'),

    path(
        'download-template/',
        views.download_template,
        name='jobcard_template_download'   # âœ… ADD THIS
    ),

    path('', include('production.urls')),
    path('job-card-entry/', views.job_card_entry, name='job_card_entry'),
    path('job-card-records/', views.job_card_records, name='job_card_records'),
    path('dispatch-entry/', views.dispatch_entry, name='dispatch_entry'),
    path('dispatch-job-card-search/', views.dispatch_job_card_search, name='dispatch_job_card_search'),
    path('dispatch-dc-duplicate-check/', views.dispatch_dc_duplicate_check, name='dispatch_dc_duplicate_check'),
    path('dispatch-records/', views.dispatch_records, name='dispatch_records'),
    path('change-history/<str:entity_type>/<int:record_id>/', views.change_history, name='change_history'),
    path('delete-record/<str:entity_type>/<int:record_id>/', views.delete_record, name='delete_record'),
    path('archived-records/', views.archived_records, name='archived_records'),
    path('restore-record/<str:entity_type>/<int:record_id>/', views.restore_record, name='restore_record'),
    path('quick-add-master/', views.quick_add_master, name='quick_add_master'),
    path('manage-user-roles/', views.manage_user_roles, name='manage_user_roles'),
    path('request-edit-override/<str:entity_type>/<int:record_id>/', views.request_edit_override, name='request_edit_override'),
    path('override-requests/', views.override_requests, name='override_requests'),
    path('review-override/<int:override_id>/', views.review_override_request, name='review_override_request'),
    path('shift-config/', views.shift_config, name='shift_config'),
    path('master-data/', views.master_data, name='master_data'),
    path('machine-master-tools/', views.machine_master_tools, name='machine_master_tools'),
    path('erp-readme/', views.erp_readme, name='erp_readme'),
    path('erp-readme/download/', views.download_erp_readme, name='download_erp_readme'),
    path('version/', views.erp_version, name='erp_version'),
    path('notifications/', notification_views.notification_list, name='notification_list'),
    path('notifications/mark-all-read/', notification_views.notification_mark_all_read, name='notification_mark_all_read'),
    path('notifications/<int:pk>/read/', notification_views.notification_mark_read, name='notification_mark_read'),
]
