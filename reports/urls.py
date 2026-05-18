from django.urls import path

from . import views

app_name = 'reports'

urlpatterns = [
    path('', views.reports_home, name='home'),
    path('<slug:report_type>/', views.report_detail, name='detail'),
]
