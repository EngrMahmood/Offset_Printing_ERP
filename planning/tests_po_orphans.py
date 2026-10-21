import os
import shutil
import tempfile
import time
from io import StringIO

from django.core.management import call_command
from django.urls import reverse, reverse_lazy
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


class PoPdfsAreReadOnceTests(TestCase):
    """The ERP extracts a PO/WO PDF once and does not keep it: nothing in the app opens it again."""

    def setUp(self):
        self.media = tempfile.mkdtemp(prefix='po_media2_')
        self.addCleanup(shutil.rmtree, self.media, True)
        override = override_settings(MEDIA_ROOT=self.media)
        override.enable()
        self.addCleanup(override.disable)
        from django.contrib.auth import get_user_model
        from core.models import Permission

        Permission.objects.get_or_create(code='action.edit_jobcard', defaults={'name': 'Edit job card'})
        self.user = get_user_model().objects.create_superuser('po_up', password='pw12345678')
        self.client.login(username='po_up', password='pw12345678')

    def upload(self, name='WO-10-2026-99999.pdf'):
        from unittest import mock

        from django.core.files.uploadedfile import SimpleUploadedFile

        payload = {
            'po_number': 'WO-10-2026-99999', 'document_type': 'work_order',
            'items': [{'sku': 'SKU-ONCE-1', 'qty': 100, 'description': 'x'}],
        }
        with mock.patch('planning.views.extract_po_from_pdf', return_value=payload), \
                mock.patch('planning.views._sync_repeat_jobs_from_po', return_value={'created': 0, 'updated': 0, 'missing_recipe': [], 'locked': 0}):
            return self.client.post(
                reverse_lazy('planning:upload_po'),
                {'po_pdf': SimpleUploadedFile(name, b'%PDF-1.4 fake pdf bytes', content_type='application/pdf')},
            )

    def stored_files(self):
        folder = os.path.join(self.media, 'planning', 'po_docs')
        return os.listdir(folder) if os.path.isdir(folder) else []

    def test_an_uploaded_pdf_is_read_but_not_kept(self):
        self.upload()
        doc = PoDocument.objects.get()
        self.assertEqual(doc.extracted_payload['po_number'], 'WO-10-2026-99999')    # the data is kept
        self.assertEqual(doc.extraction_status, 'processed')
        self.assertFalse(doc.po_file)                                                # the file is not
        self.assertEqual(self.stored_files(), [])

    def test_re_uploading_the_same_po_leaves_nothing_behind(self):
        self.upload()
        self.upload()
        self.assertEqual(PoDocument.objects.count(), 1)
        self.assertEqual(self.stored_files(), [])

    @override_settings(PO_KEEP_PDF_FILES=True)
    def test_storage_can_be_switched_back_on(self):
        self.upload()
        self.assertEqual(len(self.stored_files()), 1)
        self.assertTrue(PoDocument.objects.get().po_file)

    @override_settings(PO_KEEP_PDF_FILES=True)
    def test_replacing_a_stored_pdf_deletes_the_old_one(self):
        self.upload('WO-10-2026-99999.pdf')
        self.upload('WO-10-2026-99999-v2.pdf')
        self.assertEqual(len(self.stored_files()), 1)

    def test_records_without_a_file_do_not_break_the_po_screens(self):
        self.upload()
        for name in ('planning:po_inbox',):
            try:
                url = reverse(name)
            except Exception:
                continue
            self.assertIn(self.client.get(url).status_code, (200, 302))


class PurgePoFilesTests(TestCase):
    def setUp(self):
        self.media = tempfile.mkdtemp(prefix='po_media3_')
        self.addCleanup(shutil.rmtree, self.media, True)
        override = override_settings(MEDIA_ROOT=self.media)
        override.enable()
        self.addCleanup(override.disable)
        self.folder = os.path.join(self.media, 'planning', 'po_docs')
        os.makedirs(self.folder)
        for name in ('WO-1.pdf', 'WO-2.pdf', 'manual_po_1.txt'):
            with open(os.path.join(self.folder, name), 'wb') as handle:
                handle.write(b'x' * 100)
        PoDocument.objects.create(po_file='planning/po_docs/WO-1.pdf', extracted_payload={'po_number': 'WO-1'},
                                  extraction_status='processed')

    def run_command(self, *args):
        out = StringIO()
        call_command('purge_po_files', *args, stdout=out)
        return out.getvalue()

    def test_dry_run_is_the_default(self):
        self.assertIn('Would delete 3 file(s)', self.run_command())
        self.assertEqual(len(os.listdir(self.folder)), 3)

    def test_apply_deletes_the_files_but_keeps_every_po_record_and_its_data(self):
        self.run_command('--apply')
        self.assertEqual(os.listdir(self.folder), [])
        doc = PoDocument.objects.get()
        self.assertEqual(doc.extracted_payload, {'po_number': 'WO-1'})
        self.assertEqual(doc.extraction_status, 'processed')

    def test_only_the_po_folder_is_touched(self):
        other = os.path.join(self.media, 'item_requests')
        os.makedirs(other)
        with open(os.path.join(other, 'quote.pdf'), 'wb') as handle:
            handle.write(b'keep')
        self.run_command('--apply')
        self.assertTrue(os.path.isfile(os.path.join(other, 'quote.pdf')))

    def test_a_missing_folder_is_not_an_error(self):
        shutil.rmtree(self.folder)
        self.assertIn('nothing to do', self.run_command('--apply'))
