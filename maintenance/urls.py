from django.urls import path

from . import views

app_name = 'maintenance'

urlpatterns = [
    path('', views.dashboard, name='dashboard'),
    path('records/', views.record_list, name='record_list'),
    path('records/report/', views.complaint_create, name='complaint_create'),
    path('records/bulk-delete/', views.bulk_delete, name='bulk_delete'),
    path('records/<int:pk>/', views.record_detail, name='record_detail'),
    path('records/<int:pk>/triage/', views.record_triage, name='record_triage'),
    path('records/<int:pk>/comment/', views.record_comment, name='record_comment'),
    path('records/<int:pk>/reassign/', views.record_reassign, name='record_reassign'),
    path('records/<int:pk>/spare-parts/add/', views.spare_part_add, name='spare_part_add'),
    path('records/<int:pk>/service-jobs/add/', views.service_job_add, name='service_job_add'),
    path('records/<int:pk>/attachments/add/', views.attachment_add, name='attachment_add'),
    path('records/<int:pk>/raise-demand/', views.raise_demand, name='raise_demand'),
    path('records/<int:pk>/repair/', views.record_repair_update, name='record_repair_update'),
    path('records/<int:pk>/checklist/', views.record_checklist_update, name='record_checklist_update'),
    path('records/<int:pk>/transition/', views.record_transition, name='record_transition'),
    path('machines/', views.machine_board, name='machine_board'),
    path('downtime/', views.downtime_list, name='downtime_list'),
    path('pm-plans/', views.pm_plan_list, name='pm_plan_list'),
    path('pm-plans/new/', views.pm_plan_create, name='pm_plan_create'),
    path('pm-plans/<int:pk>/edit/', views.pm_plan_edit, name='pm_plan_edit'),
    path('pm-plans/<int:pk>/toggle/', views.pm_plan_toggle, name='pm_plan_toggle'),
    path('reports/', views.reports, name='reports'),
]
