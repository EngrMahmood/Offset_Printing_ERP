from django.core.management.base import BaseCommand

from core.data_retention import prune_old_data


class Command(BaseCommand):
    help = (
        'Delete notifications older than NOTIFICATION_RETENTION_DAYS (default 30) and clear the stored HTML of bot '
        'emails older than BOT_EMAIL_BODY_RETENTION_DAYS (default 14). Runs automatically before every backup.'
    )

    def add_arguments(self, parser):
        parser.add_argument('--dry-run', action='store_true', help='Only report what would be removed.')

    def handle(self, *args, **options):
        result = prune_old_data(dry_run=options['dry_run'])
        verb = 'Would remove' if options['dry_run'] else 'Removed'
        self.stdout.write(
            f"{verb} {result['notifications']} old notification(s); "
            f"{'would clear' if options['dry_run'] else 'cleared'} {result['bot_bodies']} old bot email body(ies)."
        )
