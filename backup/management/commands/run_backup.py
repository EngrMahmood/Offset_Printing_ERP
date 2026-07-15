from django.core.management.base import BaseCommand
from backup.services import create_backup
from django.contrib.auth import get_user_model

class Command(BaseCommand):
    help = 'Executes a database backup manually'

    def handle(self, *args, **options):
        self.stdout.write("Starting database backup...")
        try:
            history = create_backup(backup_type='MANUAL')
            if history.status == 'SUCCESS':
                self.stdout.write(self.style.SUCCESS(f"Successfully created backup: {history.file_name}"))
            else:
                self.stdout.write(self.style.ERROR(f"Backup failed: {history.error_message}"))
        except Exception as e:
            self.stdout.write(self.style.ERROR(f"Failed to run backup: {str(e)}"))
