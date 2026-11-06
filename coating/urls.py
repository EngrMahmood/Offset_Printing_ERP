from django.urls import path

from . import views

urlpatterns = [
    path('coating/', views.coating_entry, name='coating_entry'),
    path('coating/job-search/', views.coating_job_search, name='coating_job_search'),
    path('coating/records/', views.coating_records, name='coating_records'),
]
