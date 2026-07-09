from django.urls import path

from . import views

app_name = 'job_summary'

urlpatterns = [
    path('', views.job_summary_home, name='home'),
]
