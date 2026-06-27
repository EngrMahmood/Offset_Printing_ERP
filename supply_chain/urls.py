from django.urls import path
from . import views

app_name = 'supply_chain'

urlpatterns = [
    path('dashboard/', views.dashboard, name='dashboard'),
]
