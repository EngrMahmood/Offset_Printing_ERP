from django.urls import path

from bot import views

app_name = 'bot'

urlpatterns = [
    path('', views.bot_list, name='bot_list'),
    path('create/', views.bot_create, name='bot_create'),
    path('<int:pk>/edit/', views.bot_edit, name='bot_edit'),
    path('<int:pk>/preview/', views.bot_preview, name='bot_preview'),
    path('<int:pk>/test-send/', views.bot_test_send, name='bot_test_send'),
    path('<int:pk>/run-now/', views.bot_run_now, name='bot_run_now'),
    path('<int:pk>/toggle/', views.bot_toggle, name='bot_toggle'),
    path('global-toggle/', views.bot_global_toggle, name='bot_global_toggle'),
    path('executions/', views.execution_list, name='execution_list'),
    path('executions/<int:pk>/', views.execution_detail, name='execution_detail'),
]
