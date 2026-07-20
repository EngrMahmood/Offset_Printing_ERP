from django.core.management.base import BaseCommand
from backup.services import create_backup
from django.contrib.auth import get_user_model

class Command(BaseCommand):
    help = 'Executes a database backup from the command line'

    def add_arguments(self, parser):
        parser.add_argument(
            '--type',
            choices=['AUTO', 'MANUAL'],
            default='AUTO',
            help="Backup type to record in history (default: AUTO, for scheduled runs).",
        )

    def handle(self, *args, **options):
        backup_type = options['type']
        self.stdout.write(f"Starting {backup_type} database backup...")
        try:
            history = create_backup(backup_type=backup_type)
            if history.status == 'SUCCESS':
                self.stdout.write(self.style.SUCCESS(f"Successfully created backup: {history.file_name}"))
            else:
                self.stdout.write(self.style.ERROR(f"Backup failed: {history.error_message}"))
        except Exception as e:
            self.stdout.write(self.style.ERROR(f"Failed to run backup: {str(e)}"))
