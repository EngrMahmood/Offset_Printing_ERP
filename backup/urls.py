from django.urls import path
from backup import views

app_name = 'backup'

urlpatterns = [
    path('', views.backup_dashboard, name='dashboard'),
    path('manual/', views.run_manual_backup, name='run_manual_backup'),
    path('settings/', views.update_settings, name='settings'),
    path('restore/<int:backup_id>/', views.restore_backup, name='restore_backup'),
    path('download/<int:backup_id>/', views.download_backup, name='download_backup'),
    path('delete/<int:backup_id>/', views.delete_backup, name='delete_backup'),
]
