from django.apps import AppConfig


class SheetsSyncConfig(AppConfig):
    default_auto_field = 'django.db.models.BigAutoField'
    name = 'sheets_sync'
    verbose_name = 'Google Sheets Sync'

    def ready(self):
        from sheets_sync.signals import connect_all
        from sheets_sync.queue_worker import start_worker
        connect_all()
        start_worker()
