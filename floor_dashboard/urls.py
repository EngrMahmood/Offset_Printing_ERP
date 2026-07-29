from django.urls import path

from . import views

app_name = 'floor_dashboard'

urlpatterns = [
    path('', views.dashboard_view, name='dashboard'),
    path('api/data/', views.dashboard_data_api, name='dashboard_data'),
]
