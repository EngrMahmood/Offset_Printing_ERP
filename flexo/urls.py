from django.urls import path

from . import views

app_name = 'flexo'

urlpatterns = [
    path('', views.flexo_home, name='home'),

    path('jobs/', views.jobs, name='jobs'),
    path('jobs/archived/', views.jobs_archived, name='jobs_archived'),
    path('job/<int:job_id>/', views.job_detail, name='job_detail'),
    path('job/<int:job_id>/edit/', views.job_edit, name='job_edit'),
    path('job/<int:job_id>/status/', views.job_status_update, name='job_status_update'),
    path('job/<int:job_id>/priority/', views.job_priority_update, name='job_priority_update'),
    path('job/<int:job_id>/action/', views.job_action, name='job_action'),

    path('approval-queue/', views.approval_queue, name='approval_queue'),
    path('change-requests/<int:request_id>/approve/', views.change_request_approve, name='change_request_approve'),
    path('change-requests/<int:request_id>/reject/', views.change_request_reject, name='change_request_reject'),

    path('sku-recipes/', views.sku_recipes, name='sku_recipes'),
    path('sku-recipes/draft/', views.sku_recipes, {'status': 'draft'}, name='sku_recipes_draft'),
    path('sku-recipes/pending-review/', views.sku_recipes, {'status': 'pending_review'}, name='sku_recipes_pending_review'),
    path('sku-recipes/reviewed/', views.sku_recipes, {'status': 'reviewed'}, name='sku_recipes_reviewed'),
    path('sku-recipes/approved/', views.sku_recipes, {'status': 'approved'}, name='sku_recipes_approved'),
    path('sku-recipes/archived/', views.sku_recipes_archived, name='sku_recipes_archived'),
    path('sku-recipes/add/', views.sku_recipe_edit, name='sku_recipe_add'),
    path('sku-recipes/<int:recipe_id>/edit/', views.sku_recipe_edit, name='sku_recipe_edit'),

    path('manual-entry/', views.manual_po_entry, name='manual_po_entry'),
    path('po/upload/', views.po_upload, name='po_upload'),
    path('po/inbox/', views.po_inbox, name='po_inbox'),
    path('po/<int:doc_id>/review/', views.po_review, name='po_review'),

    path('job-card/<int:pk>/', views.job_card_detail, name='job_card_detail'),
    path('job-card/<int:pk>/print/', views.job_card_print, name='job_card_print'),
    path('job-card/<int:pk>/printing-entry/add/', views.job_card_add_printing_entry, name='job_card_add_printing_entry'),
    path('job-card/<int:pk>/slitting-entry/add/', views.job_card_add_slitting_entry, name='job_card_add_slitting_entry'),
]
