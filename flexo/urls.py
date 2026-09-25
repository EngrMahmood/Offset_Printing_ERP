from django.urls import path

from . import views

app_name = 'flexo'

urlpatterns = [
    path('', views.flexo_home, name='home'),

    path('<str:family>/planning/', views.planning_list, name='planning_list'),
    path('<str:family>/planning/new/', views.planning_create, name='planning_create'),
    path('planning/<int:pk>/', views.planning_detail, name='planning_detail'),
    path('planning/<int:pk>/edit/', views.planning_edit, name='planning_edit'),
    path('planning/<int:planning_pk>/job-card/create/', views.job_card_create, name='job_card_create'),

    path('job-card/<int:pk>/', views.job_card_detail, name='job_card_detail'),
    path('job-card/<int:pk>/print/', views.job_card_print, name='job_card_print'),
    path('job-card/<int:pk>/printing-entry/add/', views.job_card_add_printing_entry, name='job_card_add_printing_entry'),
    path('job-card/<int:pk>/slitting-entry/add/', views.job_card_add_slitting_entry, name='job_card_add_slitting_entry'),
]
