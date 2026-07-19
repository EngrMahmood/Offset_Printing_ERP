from django.urls import path

from . import views

app_name = 'maintenance'

urlpatterns = [
    path('', views.dashboard, name='dashboard'),
    path('records/', views.record_list, name='record_list'),
    path('records/new/', views.record_create, name='record_create'),
    path('records/bulk-delete/', views.bulk_delete, name='bulk_delete'),
    path('records/<int:pk>/', views.record_detail, name='record_detail'),
    path('records/<int:pk>/spare-parts/add/', views.spare_part_add, name='spare_part_add'),
    path('records/<int:pk>/service-jobs/add/', views.service_job_add, name='service_job_add'),
    path('records/<int:pk>/attachments/add/', views.attachment_add, name='attachment_add'),
    path('records/<int:pk>/raise-demand/', views.raise_demand, name='raise_demand'),
    path('records/<int:pk>/transition/', views.record_transition, name='record_transition'),
    path('records/<int:pk>/review/', views.record_review, name='record_review'),
    path('approval-queue/', views.approval_queue, name='approval_queue'),
    path('downtime/', views.downtime_list, name='downtime_list'),
    path('pm-plans/', views.pm_plan_list, name='pm_plan_list'),
    path('pm-plans/new/', views.pm_plan_create, name='pm_plan_create'),
    path('reports/', views.reports, name='reports'),
]
