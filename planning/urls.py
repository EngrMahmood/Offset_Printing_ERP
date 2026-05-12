from django.urls import path

from . import views

app_name = 'planning'

urlpatterns = [
    path('', views.planning_welcome, name='home'),
    path('po/', views.planning_po_root, name='po_root'),
    path('po/upload/', views.upload_po, name='po_upload'),
    path('po/manual/', views.manual_po_entry, name='po_manual'),
    path('actions/pending/', views.planning_pending_actions, name='pending_actions'),
    path('jobs/', views.planning_home, name='jobs'),
    path('jobs/drafts/', views.planning_jobs_drafts, name='jobs_drafts'),
    path('jobs/locked/', views.planning_jobs_locked, name='jobs_locked'),
    path('jobs/archived/', views.planning_jobs_archived, name='jobs_archived'),
    path('sku/', views.planning_sku_queue, name='sku_queue'),
    path('sku/recipes/', views.planning_sku_recipes_list, name='sku_recipes_list'),
    path('reports/', views.planning_report, name='reports'),
    path('job-cards/', views.planning_job_card_layout_builder, name='job_cards'),
    path('scan/', views.planning_scan, name='scan'),
    path('scan/open/<str:jc_number>/', views.planning_scan_open, name='scan_open'),
    path('report/', views.planning_report, name='report'),
    path('import-sheet/', views.import_planning_sheet, name='import_sheet'),
    path('job/<int:job_id>/', views.planning_job_detail, name='job_detail'),
    path('job/<int:job_id>/edit/', views.planning_job_edit, name='job_edit'),
    path('job/<int:job_id>/status/', views.planning_job_status_update, name='job_status_update'),
    path('job/<int:job_id>/print/', views.planning_job_card_print, name='job_card_print'),
    path('job-card-layout/', views.planning_job_card_layout_builder, name='job_card_layout_builder'),
    path('readme/', views.planning_readme, name='planning_readme'),
    path('readme/download/', views.download_planning_readme, name='download_planning_readme'),
    path('po/upload/', views.upload_po, name='upload_po'),
    path('po/manual-entry/', views.manual_po_entry, name='manual_po_entry'),
    path('po/inbox/', views.po_inbox, name='po_inbox'),
    path('po/debug/', views.po_debug_extract, name='po_debug'),
    # Backward-compatible alias for users typing /planning/po_debug
    path('po_debug/', views.po_debug_extract, name='po_debug_alias'),
]
