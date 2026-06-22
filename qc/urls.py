from django.views.generic import RedirectView
from django.urls import path
from . import views

app_name = 'qc'

urlpatterns = [
    path('', views.approval_queue, name='qc_home'),
    path('approvals/', views.approval_queue, name='approvals'),
    path('approvals/history/', views.approval_history, name='approvals_history'),
    path('approval-queue/', views.approval_queue, name='approval_queue'),
    path('job-card-approval/', views.approval_queue, name='job_card_approval'),
    path('release-approval/', views.approval_queue, name='release_approval'),
    path('release/', views.approval_queue, name='release'),
    path('job/<int:job_id>/status/', views.planning_job_status_update, name='planning_job_status_update'),
    path('po/<int:doc_id>/review/', views.po_review, name='po_review'),
    path('po/<int:doc_id>/new-skus/', views.po_new_skus, name='po_new_skus'),
    path('pending-skus/', views.planner_pending_skus_redirect, name='pending_skus'),
    path('master-review/', views.master_sku_review_queue, name='master_review'),
    path('pending-skus/ignored/', views.planner_pending_skus_ignored_redirect, name='pending_skus_ignored'),
    path('pending-skus/master-entry/', views.planner_pending_sku_master_entry_redirect, name='pending_sku_master_entry'),
    path('history/', views.approval_history, name='history'),
    path('sku-recipes/', RedirectView.as_view(pattern_name='planning:sku_recipes', permanent=False), name='sku_recipes'),
    path('sku-recipes/draft/', RedirectView.as_view(pattern_name='planning:sku_recipes_draft', permanent=False), name='sku_recipes_draft'),
    path('sku-recipes/pending-review/', RedirectView.as_view(pattern_name='planning:sku_recipes_pending_review', permanent=False), name='sku_recipes_pending_review'),
    path('sku-recipes/reviewed/', RedirectView.as_view(pattern_name='planning:sku_recipes_reviewed', permanent=False), name='sku_recipes_reviewed'),
    path('sku-recipes/approved/', RedirectView.as_view(pattern_name='planning:sku_recipes_approved', permanent=False), name='sku_recipes_approved'),
    path('sku-recipes/archived/', RedirectView.as_view(pattern_name='planning:sku_recipes_archived', permanent=False), name='sku_recipes_archived'),
    path('sku-recipes/bulk-upload/', RedirectView.as_view(pattern_name='planning:sku_recipe_bulk_upload', permanent=False), name='sku_recipe_bulk_upload'),
    path('sku-recipes/template/', RedirectView.as_view(pattern_name='planning:sku_recipe_template_download', permanent=False), name='sku_recipe_template_download'),
    path('sku-recipes/add/', RedirectView.as_view(pattern_name='planning:sku_recipe_add', permanent=False), name='sku_recipe_add'),
    path('sku-recipes/<int:recipe_id>/edit/', RedirectView.as_view(pattern_name='planning:sku_recipe_edit', permanent=False), name='sku_recipe_edit'),
]
