from django.apps import AppConfig


class ReportsConfig(AppConfig):
    default_auto_field = 'django.db.models.BigAutoField'
    name = 'reports'

    def ready(self):
        # Register built-in BI report plugins on app load.
        from reports.report_registry import builtin_reports  # noqa: F401
