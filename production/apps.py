from django.apps import AppConfig


class ProductionConfig(AppConfig):
    name = 'production'

    def ready(self):
        # Production app ready hook
        pass
