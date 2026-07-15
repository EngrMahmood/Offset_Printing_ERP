from django.apps import AppConfig


class BackupConfig(AppConfig):
    default_auto_field = 'django.db.models.BigAutoField'
    name = 'backup'

    def ready(self):
        from backup.scheduler import start_scheduler
        start_scheduler()
