from django.urls import path

from . import views

urlpatterns = [
    path('coating/', views.coating_entry, name='coating_entry'),
    path('coating/records/', views.coating_records, name='coating_records'),
]
