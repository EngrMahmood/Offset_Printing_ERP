from django.contrib import admin
from django.urls import path
from core.views import home
from core import views

urlpatterns = [
    path('admin/', admin.site.urls),
    path('', home),
    path('bulk-upload-jobcards/', views.bulk_upload_jobcards),
    path('download-template/', views.download_template),
]