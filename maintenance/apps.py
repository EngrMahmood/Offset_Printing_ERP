from django.apps import AppConfig


class MaintenanceConfig(AppConfig):
    default_auto_field = 'django.db.models.BigAutoField'
    name = 'maintenance'

    def ready(self):
        from maintenance.scheduler import start_scheduler
        start_scheduler()
