import os

from django.conf import settings
from django.core.management.base import BaseCommand


class Command(BaseCommand):
    help = (
        'Delete every file stored under media/planning/po_docs. The ERP reads a PO/WO PDF once to extract its '
        'data and never opens it again, so the stored copies only take space (and backup time). PO records and '
        'their extracted data are untouched. Dry run by default; pass --apply to delete.'
    )

    def add_arguments(self, parser):
        parser.add_argument('--apply', action='store_true')

    def handle(self, *args, **options):
        root = os.path.join(settings.MEDIA_ROOT, 'planning', 'po_docs')
        if not os.path.isdir(root):
            self.stdout.write('No planning/po_docs folder - nothing to do.')
            return
        files = [os.path.join(root, n) for n in os.listdir(root) if os.path.isfile(os.path.join(root, n))]
        size = sum(os.path.getsize(f) for f in files)
        self.stdout.write(f"{'Deleting' if options['apply'] else 'Would delete'} {len(files)} file(s), {size / 1e6:.1f} MB.")
        if options['apply']:
            for path in files:
                os.remove(path)
        else:
            self.stdout.write('Re-run with --apply to delete them.')
