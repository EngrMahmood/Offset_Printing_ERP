from django.conf import settings
from django.contrib.auth.models import User
from django.core.management.base import BaseCommand


class Command(BaseCommand):
    help = 'Creates the ai-assistant bot User if it does not already exist.'

    def handle(self, *args, **options):
        username = settings.CHAT_AI_ASSISTANT_USERNAME
        user, created = User.objects.get_or_create(
            username=username,
            defaults={
                'first_name': 'AI Assistant',
                'is_active': True,
                'is_staff': False,
            },
        )
        if created:
            user.set_unusable_password()
            user.save()
        self.stdout.write(self.style.SUCCESS(
            f'{username} ready.' if created else f'{username} already exists.'
        ))
