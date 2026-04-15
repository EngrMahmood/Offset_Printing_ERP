from django.contrib import admin
from django.urls import path
from core.views import home
from core import views

urlpatterns = [
    path('admin/', admin.site.urls),
    path('', home),

    path('bulk-upload-jobcards/', views.bulk_upload_jobcards, name='bulk_upload_jobcards'),

    path(
        'download-template/',
        views.download_template,
        name='jobcard_template_download'   # ✅ ADD THIS
    ),

    path('production-entry/', views.production_entry, name='production_entry'),
    path('production-dashboard/', views.production_dashboard, name='production_dashboard'),
    path('production-records/', views.production_records, name='production_records'),
    path('job-card-entry/', views.job_card_entry, name='job_card_entry'),
    path('job-card-records/', views.job_card_records, name='job_card_records'),
    path('dispatch-entry/', views.dispatch_entry, name='dispatch_entry'),
    path('dispatch-records/', views.dispatch_records, name='dispatch_records'),
    path('quick-add-master/', views.quick_add_master, name='quick_add_master'),
]