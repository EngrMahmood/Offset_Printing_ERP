from django.urls import path
from django.urls import include

from . import views

app_name = 'reports'

urlpatterns = [
    path('', views.reports_home, name='home'),
    path('api/', include(('reports.api.urls', 'reports_api'), namespace='reports_api')),
    path('kpi-scorecard/save-note/', views.kpi_scorecard_save_note, name='kpi_save_note'),
    path('<slug:report_type>/', views.report_detail, name='detail'),
]
