from django.urls import path

from . import views

app_name = 'planning'

urlpatterns = [
    path('', views.planning_welcome, name='home'),
    path('jobs/', views.planning_home, name='jobs'),
    path('scan/', views.planning_scan, name='scan'),
    path('scan/open/<str:jc_number>/', views.planning_scan_open, name='scan_open'),
    path('report/', views.planning_report, name='report'),
    path('import-sheet/', views.import_planning_sheet, name='import_sheet'),
    path('job/<int:job_id>/', views.planning_job_detail, name='job_detail'),
    path('job/<int:job_id>/edit/', views.planning_job_edit, name='job_edit'),
    path('job/<int:job_id>/status/', views.planning_job_status_update, name='job_status_update'),
    path('job/<int:job_id>/print/', views.planning_job_card_print, name='job_card_print'),
    path('approval-queue/', views.approval_queue, name='approval_queue'),
    path('readme/', views.planning_readme, name='planning_readme'),
    path('readme/download/', views.download_planning_readme, name='download_planning_readme'),
    path('po/upload/', views.upload_po, name='upload_po'),
    path('po/inbox/', views.po_inbox, name='po_inbox'),
    path('po/<int:doc_id>/review/', views.po_review, name='po_review'),
    path('po/<int:doc_id>/new-skus/', views.po_new_skus, name='po_new_skus'),
    path('po/debug/', views.po_debug_extract, name='po_debug'),
    # Backward-compatible alias for users typing /planning/po_debug
    path('po_debug/', views.po_debug_extract, name='po_debug_alias'),
    path('pending-skus/', views.pending_skus, name='pending_skus'),
    path('pending-skus/master-entry/', views.pending_sku_master_entry, name='pending_sku_master_entry'),
    path('sku-recipes/', views.sku_recipes_list, name='sku_recipes'),
    path('sku-recipes/bulk-upload/', views.sku_recipe_bulk_upload, name='sku_recipe_bulk_upload'),
    path('sku-recipes/add/', views.sku_recipe_edit, name='sku_recipe_add'),
    path('sku-recipes/<int:recipe_id>/edit/', views.sku_recipe_edit, name='sku_recipe_edit'),
]
