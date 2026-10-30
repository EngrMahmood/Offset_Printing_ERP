from django.apps import AppConfig


class ProductionPlanConfig(AppConfig):
    default_auto_field = 'django.db.models.BigAutoField'
    name = 'production_plan'
    verbose_name = 'Production Plan'

    def ready(self):
        from production_plan.scheduler import start_scheduler
        start_scheduler()
