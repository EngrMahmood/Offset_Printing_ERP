from django.conf import settings
from django.core.management.base import BaseCommand
from django.utils import timezone

from chat.models import Attachment


class Command(BaseCommand):
    help = "Deletes chat attachments older than CHAT_ATTACHMENT_RETENTION_DAYS (files + DB rows)."

    def add_arguments(self, parser):
        parser.add_argument('--dry-run', action='store_true', help='List what would be deleted without deleting.')

    def handle(self, *args, **options):
        retention_days = getattr(settings, 'CHAT_ATTACHMENT_RETENTION_DAYS', 180)
        cutoff = timezone.now() - timezone.timedelta(days=retention_days)
        queryset = Attachment.objects.filter(created_at__lt=cutoff)
        count = queryset.count()

        if options['dry_run']:
            for attachment in queryset[:200]:
                self.stdout.write(f'Would delete: {attachment.id} {attachment.original_filename} ({attachment.created_at})')
            self.stdout.write(self.style.WARNING(f'{count} attachment(s) older than {retention_days} days (dry run, nothing deleted).'))
            return

        # .delete() triggers chat.signals.delete_attachment_files (post_delete) to remove files from disk.
        deleted, _details = queryset.delete()
        self.stdout.write(self.style.SUCCESS(f'Deleted {deleted} chat attachment row(s) older than {retention_days} days.'))
