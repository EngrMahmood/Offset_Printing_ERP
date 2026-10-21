import os
import shutil
import tempfile
import time
from io import StringIO

from django.core.management import call_command
from django.test import TestCase, override_settings

from planning.models import PoDocument


class CleanupPoOrphansTests(TestCase):
    """Duplicate PO/WO PDFs that no record points to are removed; anything else is left alone."""

    def setUp(self):
        self.media = tempfile.mkdtemp(prefix='po_media_')
        self.addCleanup(shutil.rmtree, self.media, True)
        override = override_settings(MEDIA_ROOT=self.media)
        override.enable()
        self.addCleanup(override.disable)
        self.folder = os.path.join(self.media, 'planning', 'po_docs')
        os.makedirs(self.folder)

    def make_file(self, name, content, age_days=5):
        path = os.path.join(self.folder, name)
        with open(path, 'wb') as handle:
            handle.write(content)
        stamp = time.time() - age_days * 86400
        os.utime(path, (stamp, stamp))
        return name

    def run_command(self, *args):
        out = StringIO()
        call_command('cleanup_po_orphans', *args, stdout=out)
        return out.getvalue()

    def files(self):
        return sorted(os.listdir(self.folder))

    def reference(self, name):
        return PoDocument.objects.create(po_file=f'planning/po_docs/{name}', extraction_status='processed')

    def test_dry_run_is_the_default_and_deletes_nothing(self):
        self.make_file('WO-1.pdf', b'same')
        self.reference('WO-1.pdf')
        self.make_file('WO-1_abc.pdf', b'same')
        output = self.run_command()
        self.assertIn('Would delete 1', output)
        self.assertEqual(len(self.files()), 2)

    def test_apply_removes_unreferenced_copies_but_keeps_the_referenced_file(self):
        self.make_file('WO-1.pdf', b'same')
        self.reference('WO-1.pdf')
        for suffix in ('abc', 'def', 'ghi'):
            self.make_file(f'WO-1_{suffix}.pdf', b'same')
        self.run_command('--apply')
        self.assertEqual(self.files(), ['WO-1.pdf'])

    def test_when_the_referenced_copy_is_not_the_oldest_it_is_still_the_one_kept(self):
        self.make_file('old_copy.pdf', b'same', age_days=30)
        self.make_file('WO-2.pdf', b'same', age_days=2)
        self.reference('WO-2.pdf')
        self.run_command('--apply')
        self.assertEqual(self.files(), ['WO-2.pdf'])

    def test_an_unreferenced_file_with_unique_content_is_never_deleted(self):
        self.make_file('only_copy.pdf', b'the one and only')
        output = self.run_command('--apply')
        self.assertEqual(self.files(), ['only_copy.pdf'])
        self.assertIn('unique content were left alone', output)

    def test_unreferenced_duplicates_with_no_referenced_twin_keep_exactly_one_copy(self):
        self.make_file('a.pdf', b'dup', age_days=9)
        self.make_file('b.pdf', b'dup', age_days=8)
        self.make_file('c.pdf', b'dup', age_days=7)
        self.run_command('--apply')
        self.assertEqual(self.files(), ['a.pdf'])                  # the oldest survives

    def test_files_newer_than_a_day_are_never_touched(self):
        self.make_file('WO-3.pdf', b'same')
        self.reference('WO-3.pdf')
        self.make_file('upload_in_progress.pdf', b'same', age_days=0)
        self.run_command('--apply')
        self.assertIn('upload_in_progress.pdf', self.files())

    def test_a_referenced_file_is_never_deleted_even_if_it_duplicates_another_referenced_one(self):
        for name in ('WO-4.pdf', 'WO-4_x.pdf'):
            self.make_file(name, b'same')
            self.reference(name)
        self.run_command('--apply')
        self.assertEqual(self.files(), ['WO-4.pdf', 'WO-4_x.pdf'])

    def test_a_missing_folder_is_not_an_error(self):
        shutil.rmtree(self.folder)
        self.assertIn('nothing to do', self.run_command('--apply'))
