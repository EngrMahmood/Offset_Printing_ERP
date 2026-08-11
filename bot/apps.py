from django.apps import AppConfig


class BotConfig(AppConfig):
    default_auto_field = 'django.db.models.BigAutoField'
    name = 'bot'
    verbose_name = 'Bot / Automation'

    def ready(self):
        from bot.scheduler import start_scheduler
        start_scheduler()
