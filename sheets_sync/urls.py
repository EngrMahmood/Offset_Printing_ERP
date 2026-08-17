from django.urls import path
from sheets_sync import views

app_name = 'sheets_sync'

urlpatterns = [
    path('', views.dashboard, name='dashboard'),
]
