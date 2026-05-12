from django.urls import path
from . import views

app_name = 'qc'

urlpatterns = [
    path('', views.approval_queue, name='qc_home'),
    path('approval-queue/', views.approval_queue, name='approval_queue'),
    path('job/<int:job_id>/status/', views.planning_job_status_update, name='planning_job_status_update'),
    path('po/<int:doc_id>/review/', views.po_review, name='po_review'),
    path('po/<int:doc_id>/new-skus/', views.po_new_skus, name='po_new_skus'),
    path('pending-skus/', views.pending_skus, name='pending_skus'),
    path('pending-skus/ignored/', views.pending_skus_ignored, name='pending_skus_ignored'),
    path('pending-skus/master-entry/', views.pending_sku_master_entry, name='pending_sku_master_entry'),
    path('sku-recipes/', views.sku_recipes_list, name='sku_recipes'),
    path('sku-recipes/draft/', views.sku_recipes_status, {'status': 'draft'}, name='sku_recipes_draft'),
    path('sku-recipes/pending-review/', views.sku_recipes_status, {'status': 'pending_review'}, name='sku_recipes_pending_review'),
    path('sku-recipes/reviewed/', views.sku_recipes_status, {'status': 'reviewed'}, name='sku_recipes_reviewed'),
    path('sku-recipes/approved/', views.sku_recipes_status, {'status': 'approved'}, name='sku_recipes_approved'),
    path('sku-recipes/archived/', views.sku_recipes_archived, name='sku_recipes_archived'),
    path('sku-recipes/bulk-upload/', views.sku_recipe_bulk_upload, name='sku_recipe_bulk_upload'),
    path('sku-recipes/template/', views.sku_recipe_template_download, name='sku_recipe_template_download'),
    path('sku-recipes/add/', views.sku_recipe_edit, name='sku_recipe_add'),
    path('sku-recipes/<int:recipe_id>/edit/', views.sku_recipe_edit, name='sku_recipe_edit'),
]
