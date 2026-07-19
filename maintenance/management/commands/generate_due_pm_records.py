from django.contrib.auth import get_user_model
from django.core.management.base import BaseCommand

from maintenance.services import generate_due_pm_records


class Command(BaseCommand):
    help = 'Generate MaintenanceRecords for any preventive-maintenance plan that is now due.'

    def handle(self, *args, **options):
        User = get_user_model()
        system_user = User.objects.filter(is_superuser=True).order_by('id').first()
        if system_user is None:
            self.stderr.write('No superuser found to attribute auto-generated records to. Aborting.')
            return
        created = generate_due_pm_records(actor=system_user)
        self.stdout.write(self.style.SUCCESS(f'Created {len(created)} due preventive-maintenance record(s).'))
