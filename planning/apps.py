from django.apps import AppConfig


class PlanningConfig(AppConfig):
    name = 'planning'

    def ready(self):
        # Import signal handlers
        try:
            from . import signals  # noqa: F401
        except Exception:
            pass
