from django.urls import path

from . import views

app_name = 'planning'

urlpatterns = [
    path('', views.planning_home, name='home'),
    path('scan/', views.planning_scan, name='scan'),
    path('scan/open/<str:jc_number>/', views.planning_scan_open, name='scan_open'),
    path('report/', views.planning_report, name='report'),
    path('import-sheet/', views.import_planning_sheet, name='import_sheet'),
    path('job/<int:job_id>/', views.planning_job_detail, name='job_detail'),
    path('job/<int:job_id>/edit/', views.planning_job_edit, name='job_edit'),
    path('job/<int:job_id>/status/', views.planning_job_status_update, name='job_status_update'),
    path('job/<int:job_id>/print/', views.planning_job_card_print, name='job_card_print'),
    path('po/upload/', views.upload_po, name='upload_po'),
    path('po/<int:doc_id>/review/', views.po_review, name='po_review'),
    path('po/<int:doc_id>/new-skus/', views.po_new_skus, name='po_new_skus'),
]
