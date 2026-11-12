from django.core.management.base import BaseCommand

from planning.models import PlanningJob
from planning.sku_classification import refresh_stale_repeat_flag


class Command(BaseCommand):
    help = (
        'Draft jobs flagged Repeat only because of an earlier WO that was terminated (nothing produced, no approved '
        'master) go back to New. Shows what it would change unless --apply is given.'
    )

    def add_arguments(self, parser):
        parser.add_argument('--apply', action='store_true', help='Save the change (default is a dry run).')

    def handle(self, *args, **options):
        apply = options['apply']
        changed = 0
        for job in PlanningJob.objects.filter(is_active=True, repeat_flag__iexact='Repeat').order_by('id'):
            if refresh_stale_repeat_flag(job, save=apply):
                changed += 1
                self.stdout.write(f'{"Fixed" if apply else "Would fix"}: {job.jc_number} ({job.po_number}) {job.sku} -> New')
        self.stdout.write(f'{changed} job(s) {"changed" if apply else "would change"}.')
