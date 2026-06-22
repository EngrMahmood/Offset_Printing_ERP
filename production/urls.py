from django.urls import path
from . import views

urlpatterns = [
    path('production-entry/', views.production_entry, name='production_entry'),
    path('production-wip/', views.production_wip, name='production_wip'),
    path('production-dashboard/', views.production_dashboard, name='production_dashboard'),
    path('production-records/', views.production_records, name='production_records'),
    path('create-operator/', views.create_operator_ajax, name='create_operator_ajax'),
    path('create-machine/', views.create_machine_ajax, name='create_machine_ajax'),
    path('create-supervisor/', views.create_supervisor_ajax, name='create_supervisor_ajax'),
]
