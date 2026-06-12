from django.urls import path

from . import views

app_name = 'manual_working'

urlpatterns = [
    path('', views.manual_working_list, name='manual_working_list'),
]
