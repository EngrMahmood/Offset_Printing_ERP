from django.urls import path

from reports.api import views

app_name = 'reports_api'

urlpatterns = [
    path('reports/', views.list_reports, name='list_reports'),
    path('reports/<slug:slug>/run/', views.run_report_api, name='run_report'),
    path('reports/<slug:slug>/export/', views.export_report_api, name='export_report'),
    path('schedules/', views.list_schedules_api, name='list_schedules'),
    path('schedules/create/', views.create_schedule_api, name='create_schedule'),
]
