from django.contrib.auth import get_user_model
from django.core.management.base import BaseCommand

from printing_plates.services import bulk_cancel_stale_open_plate_requests


class Command(BaseCommand):
    help = 'Cancel and archive stale open plate requests on released/in-production jobs.'

    def add_arguments(self, parser):
        parser.add_argument(
            '--dry-run',
            action='store_true',
            help='Show how many requests would be cancelled without changing data.',
        )
        parser.add_argument(
            '--username',
            default='admin',
            help='User recorded as actor on cancelled requests (default: admin).',
        )

    def handle(self, *args, **options):
        User = get_user_model()
        actor = User.objects.filter(username=options['username']).first()
        if actor is None:
            actor = User.objects.filter(is_superuser=True).order_by('id').first()
        if actor is None:
            self.stderr.write(self.style.ERROR('No actor user found. Create an admin user or pass --username.'))
            return

        result = bulk_cancel_stale_open_plate_requests(
            actor=actor,
            dry_run=options['dry_run'],
        )
        if options['dry_run']:
            self.stdout.write(
                self.style.WARNING(
                    f'Dry run: {result["total"]} stale open plate request(s) would be cancelled.'
                )
            )
            for jc in result.get('sample_jc_numbers') or []:
                self.stdout.write(f'  - {jc}')
            return

        self.stdout.write(
            self.style.SUCCESS(
                f'Cancelled {result["cancelled"]} of {result["total"]} stale open plate request(s).'
            )
        )
        if result.get('errors'):
            self.stdout.write(self.style.WARNING(f'{len(result["errors"])} error(s):'))
            for item in result['errors'][:20]:
                self.stdout.write(f'  PR#{item["id"]}: {item["message"]}')
