import hashlib
import os
import time

from django.conf import settings
from django.core.management.base import BaseCommand

from planning.models import PoDocument

MIN_AGE_SECONDS = 24 * 3600


def find_orphan_duplicates(root, now=None):
    """Split the PO/WO folder into (files to delete, unique unreferenced files left alone).

    A file is deletable only if (a) no PoDocument points to it, (b) another file with
    IDENTICAL content is being kept (a referenced one, else the oldest copy), and (c) it is
    older than a day, so an upload in progress is never touched. An unreferenced file whose
    content exists nowhere else is reported but never deleted.
    """
    now = now or time.time()
    referenced = {os.path.basename(n) for n in PoDocument.objects.exclude(po_file='').values_list('po_file', flat=True)}
    groups = {}
    for name in os.listdir(root):
        path = os.path.join(root, name)
        if not os.path.isfile(path):
            continue
        with open(path, 'rb') as handle:
            digest = hashlib.sha256(handle.read()).hexdigest()
        groups.setdefault(digest, []).append(name)

    doomed, unique_orphans = [], []
    for names in groups.values():
        names.sort(key=lambda n: os.path.getmtime(os.path.join(root, n)))
        keep = next((n for n in names if n in referenced), names[0])
        for name in names:
            if name == keep or name in referenced:
                continue
            if now - os.path.getmtime(os.path.join(root, name)) < MIN_AGE_SECONDS:
                continue
            doomed.append(name)
        if len(names) == 1 and names[0] not in referenced:
            unique_orphans.append(names[0])
    return sorted(doomed), sorted(unique_orphans)


class Command(BaseCommand):
    help = (
        'Remove duplicate PO/WO PDFs that no PoDocument points to (left behind when a PO is re-uploaded). '
        'Dry run by default; pass --apply to delete. Unique, unreferenced files are never deleted.'
    )

    def add_arguments(self, parser):
        parser.add_argument('--apply', action='store_true', help='Actually delete the files.')

    def handle(self, *args, **options):
        root = os.path.join(settings.MEDIA_ROOT, 'planning', 'po_docs')
        if not os.path.isdir(root):
            self.stdout.write('No planning/po_docs folder - nothing to do.')
            return
        doomed, unique_orphans = find_orphan_duplicates(root)
        size = sum(os.path.getsize(os.path.join(root, n)) for n in doomed)
        verb = 'Deleting' if options['apply'] else 'Would delete'
        self.stdout.write(f'{verb} {len(doomed)} duplicate orphan file(s), {size / 1e6:.1f} MB.')
        if options['apply']:
            for name in doomed:
                os.remove(os.path.join(root, name))
        if unique_orphans:
            self.stdout.write(f'{len(unique_orphans)} unreferenced file(s) with unique content were left alone:')
            for name in unique_orphans[:20]:
                self.stdout.write(f'  {name}')
        if not options['apply'] and doomed:
            self.stdout.write('Re-run with --apply to delete them.')
