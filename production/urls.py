from django.urls import path
from . import views
from .packing_entry import packing_production_entry, packing_job_card_search, packing_records
from .released_jobs import released_jobs, request_plate_remake_view

urlpatterns = [
    path('production-entry/', views.production_entry, name='production_entry'),
    path('production-entry/printing/', views.printing_production_entry, name='printing_production_entry'),
    path('production-entry/packing/', packing_production_entry, name='packing_production_entry'),
    path('released-jobs/', released_jobs, name='released_jobs'),
    path('released-jobs/request-plates/', request_plate_remake_view, name='request_plate_remake'),
    path('printing-job-card-search/', views.printing_job_card_search, name='printing_job_card_search'),
    path('packing-job-card-search/', packing_job_card_search, name='packing_job_card_search'),
    path('production-wip/', views.production_wip, name='production_wip'),
    path('production-dashboard/', views.production_dashboard, name='production_dashboard'),
    path('production-records/', views.production_records, name='production_records'),
    path('production-records/packing/', packing_records, name='packing_records'),
    path('create-operator/', views.create_operator_ajax, name='create_operator_ajax'),
    path('create-machine/', views.create_machine_ajax, name='create_machine_ajax'),
    path('create-supervisor/', views.create_supervisor_ajax, name='create_supervisor_ajax'),
    path('create-sorter/', views.create_sorter_ajax, name='create_sorter_ajax'),
]
