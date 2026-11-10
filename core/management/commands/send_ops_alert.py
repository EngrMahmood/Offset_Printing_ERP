"""Send an operations alert email (server down, certificate expiring, disk full ...).

Used by scripts/cloud/health_alert.sh, which runs from the host's cron so the check does not depend on the
web app being healthy. Mail goes out through the same Gmail account the ERP already uses for its own emails
(Settings -> Email, or the GMAIL_* environment variables).

    python manage.py send_ops_alert --subject "..." --body "..." [--to a@b.com]
"""
import os

from django.core.mail import EmailMessage
from django.core.management.base import BaseCommand, CommandError

DEFAULT_RECIPIENT = 'mahmood7692@gmail.com'


class Command(BaseCommand):
    help = 'Email an operations alert (certificate expiry, site down, disk full) to the administrator.'

    def add_arguments(self, parser):
        parser.add_argument('--subject', required=True)
        parser.add_argument('--body', required=True)
        parser.add_argument('--to', default=os.environ.get('OPS_ALERT_EMAIL', DEFAULT_RECIPIENT))

    def handle(self, *args, **options):
        recipients = [addr.strip() for addr in options['to'].split(',') if addr.strip()]
        if not recipients:
            raise CommandError('No recipient given.')
        message = EmailMessage(subject=options['subject'], body=options['body'], to=recipients)
        sent = message.send(fail_silently=False)
        if not sent:
            raise CommandError('The mail backend did not send the message.')
        self.stdout.write(f'Alert sent to {", ".join(recipients)}.')
