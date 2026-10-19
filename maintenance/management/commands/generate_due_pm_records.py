from django.contrib.auth import get_user_model
from django.core.management.base import BaseCommand

from maintenance.services import generate_due_pm_records, remind_overdue_pm


class Command(BaseCommand):
    help = (
        'Raise a maintenance ticket for every preventive-maintenance plan that is due, and remind engineers '
        'about overdue ones. Safe to run as often as you like: a plan never raises two tickets for one due occurrence.'
    )

    def handle(self, *args, **options):
        system_user = get_user_model().objects.filter(is_superuser=True).order_by('id').first()
        if system_user is None:
            self.stderr.write('No superuser found to attribute auto-generated records to. Aborting.')
            return
        result = generate_due_pm_records(actor=system_user)
        reminded = remind_overdue_pm(actor=system_user)
        self.stdout.write(self.style.SUCCESS(
            f"PM run: {len(result['created'])} ticket(s) created, {result['skipped']} plan(s) not due, "
            f"{len(result['failed'])} failed, {len(reminded)} overdue reminder(s) sent."
        ))
        for record in result['created']:
            self.stdout.write(f'  created {record.record_no} - {record.machine} - {record.pm_plan.title}')
        for plan, message in result['failed']:
            self.stderr.write(f'  FAILED plan {plan.pk} ({plan.title}): {message}')
