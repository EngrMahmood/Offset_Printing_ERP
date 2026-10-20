from unittest import mock

from django.test import TestCase

from core.jobcard_service import transition_job_card_status
from core.models import JobCard


class MultiFormCompletionCascadeTests(TestCase):
    """A multi-form book is one deliverable: its forms follow the base (Cover) job card."""

    def setUp(self):
        import datetime

        from core.models import Machine

        machine = Machine.objects.create(name='BK Machine', is_active=True)

        def make(no, parent=None, label=''):
            return JobCard.objects.create(
                job_card_no=no, order_qty=100, ups=1, total_sheet_quantity=100, status='in_production', is_active=True,
                SKU='SKU-BOOK', po_date=datetime.date.today(), total_colors=4, total_impressions_required=100, machine_name=machine,
                parent_job_card=parent, form_label=label, unit_type='book', pcs_per_unit=1,
            )

        self.root = make('JC-BK-1', label='Cover')
        self.inner_a = make('JC-BK-1.1', parent=self.root, label='Inner')
        self.inner_b = make('JC-BK-1.2', parent=self.root, label='Inner 2')
        self.other = make('JC-OTHER-1')
        patcher = mock.patch('core.jobcard_service.job_card_completion_blockers', return_value=[])
        patcher.start()
        self.addCleanup(patcher.stop)

    def statuses(self):
        return [JobCard.objects.get(pk=c.pk).workflow_status for c in (self.root, self.inner_a, self.inner_b)]

    def test_completing_the_base_card_completes_every_form(self):
        transition_job_card_status(self.root, 'completed')
        self.assertEqual(self.statuses(), ['completed', 'completed', 'completed'])

    def test_closing_the_base_card_closes_every_form(self):
        transition_job_card_status(self.root, 'completed')
        transition_job_card_status(JobCard.objects.get(pk=self.root.pk), 'closed')
        self.assertEqual(self.statuses(), ['closed', 'closed', 'closed'])

    def test_forms_need_no_packing_or_dispatch_of_their_own(self):
        # Without the cascade exemption the inner forms (printing only) could never complete.
        with mock.patch('core.jobcard_service.job_card_completion_blockers', return_value=['No packing entries have been logged for this job.']) as blockers:
            blockers.side_effect = lambda card: [] if card.pk == self.root.pk else ['No packing entries have been logged for this job.']
            transition_job_card_status(self.root, 'completed')
        self.assertEqual(self.statuses(), ['completed', 'completed', 'completed'])

    def test_reopening_the_base_card_reopens_the_forms(self):
        transition_job_card_status(self.root, 'completed')
        transition_job_card_status(JobCard.objects.get(pk=self.root.pk), 'in_production')
        self.assertEqual(self.statuses(), ['in_production', 'in_production', 'in_production'])

    def test_a_form_that_was_still_released_follows_too(self):
        JobCard.objects.filter(pk=self.inner_b.pk).update(status='released')
        transition_job_card_status(self.root, 'completed')
        self.assertEqual(self.statuses(), ['completed', 'completed', 'completed'])

    def test_unrelated_job_cards_are_untouched(self):
        transition_job_card_status(self.root, 'completed')
        self.assertEqual(JobCard.objects.get(pk=self.other.pk).workflow_status, 'in_production')

    def test_completing_a_standalone_card_still_works(self):
        transition_job_card_status(self.other, 'completed')
        self.assertEqual(JobCard.objects.get(pk=self.other.pk).workflow_status, 'completed')
        self.assertEqual(self.statuses(), ['in_production'] * 3)

    def test_a_form_cannot_be_driven_through_the_cascade_path_from_outside(self):
        # The exemption exists only for the system cascade; a normal transition on a form keeps its rules.
        from django.core.exceptions import ValidationError

        with mock.patch('core.jobcard_service.job_card_completion_blockers', return_value=['blocked']):
            with self.assertRaises(ValidationError):
                transition_job_card_status(self.inner_a, 'completed')

    def test_each_cascaded_form_is_logged(self):
        from core.models import ChangeLog

        transition_job_card_status(self.root, 'completed')
        entry = ChangeLog.objects.filter(entity_type='job_card', record_id=self.inner_a.pk).order_by('-id').first()
        self.assertIsNotNone(entry)
        self.assertIn('JC-BK-1', entry.change_reason)
