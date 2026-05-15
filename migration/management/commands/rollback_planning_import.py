from django.core.management.base import BaseCommand, CommandError

from migration.models import MigrationImportJob, RowImportStatus
from migration.services.importer import get_imported_planning_jobs, rollback_imported_planning_jobs


class Command(BaseCommand):
    help = 'Rollback PlanningJob records created by a migration import job.'

    def add_arguments(self, parser):
        parser.add_argument('job_id', type=int, help='MigrationImportJob ID to rollback')
        parser.add_argument(
            '--dry-run',
            action='store_true',
            help='Show which PlanningJob records would be deleted without deleting them.',
        )

    def handle(self, *args, **options):
        job_id = options['job_id']
        dry_run = options['dry_run']

        import_job = MigrationImportJob.objects.filter(id=job_id).first()
        if not import_job:
            raise CommandError(f'MigrationImportJob #{job_id} does not exist.')

        jobs = get_imported_planning_jobs(import_job)
        if not jobs:
            self.stdout.write(self.style.WARNING('No imported PlanningJob records found for this import job.'))
            return

        self.stdout.write(f'Found {len(jobs)} PlanningJob record(s) imported by job #{job_id}.')
        for job in jobs:
            self.stdout.write(f'  - {job.id} {job.jc_number} / {job.po_number} / {job.sku}')

        if dry_run:
            self.stdout.write(self.style.SUCCESS('Dry run complete. No records were deleted.'))
            return

        deleted_count = rollback_imported_planning_jobs(import_job)
        self.stdout.write(self.style.SUCCESS(f'Deleted {deleted_count} PlanningJob record(s).'))
