from django.urls import path
from . import views

app_name = 'tasks'

urlpatterns = [
    path('', views.dashboard, name='dashboard'),
    path('create/', views.create_task, name='create'),
    path('<int:pk>/', views.task_detail, name='detail'),
    path('<int:pk>/edit/', views.edit_task, name='edit'),
    path('<int:pk>/delete/', views.delete_task, name='delete'),
    path('<int:pk>/update-status/', views.update_status, name='update_status'),
    path('<int:pk>/score/', views.grade_task, name='score'),
    path('teams/', views.teams_list, name='teams'),
    path('automation/', views.automation, name='automation'),
]
