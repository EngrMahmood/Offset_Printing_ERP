from django.core.management.base import BaseCommand
from django.utils import timezone

from chat.models import CallSession

STALE_RINGING_MINUTES = 2
STALE_ACTIVE_HOURS = 6


class Command(BaseCommand):
    help = "Safety net: closes out CallSession rows stuck in 'ringing' or 'active' (e.g. from a crashed browser tab)."

    def handle(self, *args, **options):
        now = timezone.now()

        stale_ringing = CallSession.objects.filter(
            status='ringing', started_at__lt=now - timezone.timedelta(minutes=STALE_RINGING_MINUTES),
        )
        ringing_count = stale_ringing.update(status='missed', ended_at=now, end_reason='timeout')

        stale_active = CallSession.objects.filter(
            status='active', started_at__lt=now - timezone.timedelta(hours=STALE_ACTIVE_HOURS),
        )
        active_count = stale_active.update(status='ended', ended_at=now, end_reason='timeout')

        self.stdout.write(self.style.SUCCESS(
            f'Expired {ringing_count} stale ringing call(s) and {active_count} stale active call(s).'
        ))
