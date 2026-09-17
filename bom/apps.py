from django.apps import AppConfig


class BomConfig(AppConfig):
    default_auto_field = 'django.db.models.BigAutoField'
    name = 'bom'
    verbose_name = 'Bill of Materials'

    def ready(self):
        from . import signals  # noqa: F401
