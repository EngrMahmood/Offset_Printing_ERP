from django.contrib import admin
from django.urls import path
from django.contrib.auth import views as auth_views
from core.views import home
from core import views

urlpatterns = [
    path('admin/', admin.site.urls),
    path('login/', auth_views.LoginView.as_view(template_name='login.html'), name='login'),
    path('logout/', auth_views.LogoutView.as_view(next_page='login'), name='logout'),
    path('', home, name='home'),

    path('bulk-upload-jobcards/', views.bulk_upload_jobcards, name='bulk_upload_jobcards'),

    path(
        'download-template/',
        views.download_template,
        name='jobcard_template_download'   # ✅ ADD THIS
    ),

    path('production-entry/', views.production_entry, name='production_entry'),
    path('production-dashboard/', views.production_dashboard, name='production_dashboard'),
    path('production-records/', views.production_records, name='production_records'),
    path('job-card-entry/', views.job_card_entry, name='job_card_entry'),
    path('job-card-records/', views.job_card_records, name='job_card_records'),
    path('dispatch-entry/', views.dispatch_entry, name='dispatch_entry'),
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
    path('machine-master-tools/', views.machine_master_tools, name='machine_master_tools'),
    path('erp-readme/', views.erp_readme, name='erp_readme'),
    path('erp-readme/download/', views.download_erp_readme, name='download_erp_readme'),
]