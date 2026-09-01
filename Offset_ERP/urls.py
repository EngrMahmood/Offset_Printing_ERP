from django.conf import settings
from django.conf.urls.static import static
from django.contrib import admin
from django.urls import path, include
from django.contrib.auth import views as auth_views
from django.views.static import serve as static_serve
from core.views import home
from core import views
from core import notification_views
from core import job_card_finalization

urlpatterns = [
    path('admin/', admin.site.urls),
    # Served at the root (not under /static/) so its default scope covers the
    # whole app — required for "Add to Home Screen" / Play Store TWA install.
    path(
        'service-worker.js',
        static_serve,
        {'document_root': settings.BASE_DIR / 'core/static/core/pwa', 'path': 'service-worker.js'},
        name='service_worker',
    ),
    # Digital Asset Links — proves this domain and the Play Store TWA app
    # (org.duckdns.offseterp.twa) are controlled by the same party, required
    # for the TWA to open without a browser address bar.
    path(
        '.well-known/assetlinks.json',
        static_serve,
        {'document_root': settings.BASE_DIR / 'core/static/core/pwa', 'path': 'assetlinks.json'},
        name='asset_links',
    ),
    path('login/', auth_views.LoginView.as_view(template_name='login.html'), name='login'),
    path('logout/', auth_views.LogoutView.as_view(next_page='login'), name='logout'),
    path('', home, name='home'),
    path('planning/', include('planning.urls', namespace='planning')),
    path('job-summary/', include('job_summary.urls', namespace='job_summary')),
    path('printing-plates/', include('printing_plates.urls', namespace='printing_plates')),
    path('reports/', include('reports.urls', namespace='reports')),
    path('migration/', include('migration.urls', namespace='migration')),
    path('qc/', include('qc.urls', namespace='qc')),
    path('manual-working/', include('manual_working.urls', namespace='manual_working')),
    path('supply-chain/', include('supply_chain.urls', namespace='supply_chain')),
    path('audit/', include('audit.urls', namespace='audit')),
    path('tasks/', include('tasks.urls', namespace='tasks')),
    path('backup/', include('backup.urls', namespace='backup')),
    path('sheets-sync/', include('sheets_sync.urls', namespace='sheets_sync')),
    path('bot/', include('bot.urls', namespace='bot')),
    path('maintenance/', include('maintenance.urls', namespace='maintenance')),
    path('floor-dashboard/', include('floor_dashboard.urls', namespace='floor_dashboard')),
    path('chat/', include('chat.urls', namespace='chat')),
    path('api/chat/', include('chat.api_urls')),

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
    path('dispatch/<int:dispatch_id>/request-change/', views.request_dispatch_change, name='request_dispatch_change'),
    path('dispatch/change-requests/', views.dispatch_change_queue, name='dispatch_change_queue'),
    path('dispatch/change-requests/<int:request_id>/approve/', views.approve_dispatch_change_request, name='approve_dispatch_change_request'),
    path('dispatch/change-requests/<int:request_id>/reject/', views.reject_dispatch_change_request, name='reject_dispatch_change_request'),
    path('change-history/<str:entity_type>/<int:record_id>/', views.change_history, name='change_history'),
    path('delete-record/<str:entity_type>/<int:record_id>/', views.delete_record, name='delete_record'),
    path('archived-records/', views.archived_records, name='archived_records'),
    path('restore-record/<str:entity_type>/<int:record_id>/', views.restore_record, name='restore_record'),
    path('quick-add-master/', views.quick_add_master, name='quick_add_master'),
    path('manage-user-roles/', views.manage_user_roles, name='manage_user_roles'),
    path('request-edit-override/<str:entity_type>/<int:record_id>/', views.request_edit_override, name='request_edit_override'),
    path('override-requests/', views.override_requests, name='override_requests'),
    path('my-override-requests/', views.my_override_requests, name='my_override_requests'),
    path('review-override/<int:override_id>/', views.review_override_request, name='review_override_request'),
    path('shift-config/', views.shift_config, name='shift_config'),
    path('master-data/', views.master_data, name='master_data'),
    path('job-cards/finalize/', job_card_finalization.job_card_finalization_queue, name='job_card_finalization_queue'),
    path('job-cards/finalize/close/', job_card_finalization.job_card_finalization_close, name='job_card_finalization_close'),
    path('job-cards/finalize/reopen/', job_card_finalization.job_card_finalization_reopen, name='job_card_finalization_reopen'),
    path('job-cards/finalize/export/', job_card_finalization.job_card_finalization_export, name='job_card_finalization_export'),
    path('machine-master-tools/', views.machine_master_tools, name='machine_master_tools'),
    path('erp-readme/', views.erp_readme, name='erp_readme'),
    path('erp-readme/download/', views.download_erp_readme, name='download_erp_readme'),
    path('version/', views.erp_version, name='erp_version'),
    path('nav-layout/', views.save_nav_layout, name='save_nav_layout'),
    path('flexo/', views.coming_soon, {'feature': 'Flexo'}, name='flexo_placeholder'),
    path('sublimation/', views.coming_soon, {'feature': 'Sublimation'}, name='sublimation_placeholder'),
    path('notifications/', notification_views.notification_list, name='notification_list'),
    path('notifications/mark-all-read/', notification_views.notification_mark_all_read, name='notification_mark_all_read'),
    path('notifications/<int:pk>/read/', notification_views.notification_mark_read, name='notification_mark_read'),
    
    # Settings & Forgot Password Paths
    path('settings/', views.notification_settings_home, name='notification_settings_home'),
    path('settings/rules/add/', views.notification_rule_add, name='notification_rule_add'),
    path('settings/rules/<int:rule_id>/delete/', views.notification_rule_delete, name='notification_rule_delete'),
    path('settings/transitions/add/', views.workflow_transition_add, name='workflow_transition_add'),
    path('settings/transitions/<int:transition_id>/delete/', views.workflow_transition_delete, name='workflow_transition_delete'),
    path('settings/access-control/roles/add/', views.access_role_create, name='access_role_create'),
    path('settings/access-control/roles/<int:role_id>/delete/', views.access_role_delete, name='access_role_delete'),
    path('settings/access-control/roles/permissions/', views.access_role_permissions_edit, name='access_role_permissions_edit'),
    path('settings/access-control/user-overrides/', views.access_user_overrides_edit, name='access_user_overrides_edit'),
    path('settings/access-control/users/role/', views.access_user_role_update, name='access_user_role_update'),
    path('settings/access-control/users/official-email/', views.access_user_official_email_update, name='access_user_official_email_update'),
    path('settings/access-control/users/password-reset/', views.access_user_password_reset, name='access_user_password_reset'),
    path('settings/access-control/users/toggle-active/', views.access_user_toggle_active, name='access_user_toggle_active'),
    path('settings/users/create/', views.user_create, name='user_create'),
    path('settings/email/', views.email_settings_edit, name='email_settings_edit'),
    path('settings/permission-audit/', views.permission_audit_view, name='permission_audit_view'),
    path('settings/ai/', views.ai_settings_edit, name='ai_settings_edit'),
    path('forgot-password/', views.forgot_password, name='forgot_password'),

    # Self-service password reset (Gmail SMTP — see settings.py)
    path(
        'password-reset/',
        auth_views.PasswordResetView.as_view(
            template_name='password_reset_form.html',
            email_template_name='password_reset_email.html',
            subject_template_name='password_reset_subject.txt',
            success_url='/password-reset/done/',
        ),
        name='password_reset',
    ),
    path(
        'password-reset/done/',
        auth_views.PasswordResetDoneView.as_view(template_name='password_reset_done.html'),
        name='password_reset_done',
    ),
    path(
        'reset/<uidb64>/<token>/',
        auth_views.PasswordResetConfirmView.as_view(
            template_name='password_reset_confirm.html',
            success_url='/reset/done/',
        ),
        name='password_reset_confirm',
    ),
    path(
        'reset/done/',
        auth_views.PasswordResetCompleteView.as_view(template_name='password_reset_complete.html'),
        name='password_reset_complete',
    ),
]

if settings.DEBUG:
    urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)
    # runserver auto-serves STATIC_URL, but the standalone `daphne` process used
    # for the HTTPS listener (see DEPLOYMENT.md step 6) doesn't, so serve it here.
    from django.contrib.staticfiles.urls import staticfiles_urlpatterns
    urlpatterns += staticfiles_urlpatterns()
