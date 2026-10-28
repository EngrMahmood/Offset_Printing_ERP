import datetime

from django.contrib.auth import get_user_model
from django.test import RequestFactory, TestCase
from django.urls import reverse

from core.models import JobCardWipStatus, PrintingCompletionRequest, Production, UserProfile
from planning.tests_merge import make_card, make_job
from production.printing_completion import (
    guard_completed_entry, is_printing_complete, pending_printing_pcs, printing_completion_status,
)
from production.wip_service import evaluate_and_update_job_wip_status, get_system_calculated_status_name


class PrintingCompletionRuleTests(TestCase):
    """Order 1,000 pcs at 4 ups (250 sheets needed), 5% tolerance."""

    def setUp(self):
        self.user = get_user_model().objects.create_user('sup_user', password='pw12345678')
        job = make_job('PC1', 1000, sku='SKU-COMPLETION')
        self.card = make_card(job)
        self.card.total_sheet_quantity = 260
        self.card.production_tolerance_percent = 5
        self.card.save()

    def entry(self, good, waste=0, status='in_progress', **extra):
        return Production.objects.create(
            job_card=self.card, entry_type='printing', date='2026-10-27', shift='A', print_pass_number=1,
            output_sheets=good, waste_sheets=waste, impressions=good + waste, status=status, **extra,
        )

    def test_rule_1_order_quantity_reached(self):
        self.entry(250)
        status = printing_completion_status(self.card)
        self.assertEqual((status['complete'], status['rule']), (True, 'order_qty'))
        self.assertEqual(get_system_calculated_status_name(self.card), 'Printing Completed')
        self.assertEqual(pending_printing_pcs(self.card), 0)

    def test_rule_2_good_plus_waste_use_every_planned_sheet(self):
        self.entry(248, waste=12)                       # 992 pcs, 8 short, but all 260 planned sheets are gone
        status = printing_completion_status(self.card)
        self.assertEqual((status['complete'], status['rule']), (True, 'sheets_used'))
        self.assertEqual(status['shortfall_pcs'], 8)
        self.assertEqual(get_system_calculated_status_name(self.card), 'Printing Completed')
        self.assertEqual(pending_printing_pcs(self.card), 0)                 # leaves the Pending Work printing list

    def test_rule_3_marked_completed_within_tolerance(self):
        self.entry(245, status='completed')             # 980 pcs: 2% short, sheets left over
        status = printing_completion_status(self.card)
        self.assertEqual((status['complete'], status['rule']), (True, 'completed_within_tolerance'))
        self.assertEqual(get_system_calculated_status_name(self.card), 'Printing Completed')

    def test_within_tolerance_but_not_marked_completed_stays_partial(self):
        self.entry(245)                                 # 2% short, nobody said it is done
        self.assertFalse(is_printing_complete(self.card))
        self.assertEqual(get_system_calculated_status_name(self.card), 'Partial Printing')
        self.assertEqual(pending_printing_pcs(self.card), 20)

    def test_genuinely_short_stays_partial_even_when_marked_completed(self):
        self.entry(100, waste=2, status='completed')    # 400 of 1,000 pcs, 102 of 260 sheets
        status = printing_completion_status(self.card)
        self.assertFalse(status['complete'])
        self.assertEqual(get_system_calculated_status_name(self.card), 'Partial Printing')
        self.assertEqual(pending_printing_pcs(self.card), 600)

    def test_rule_4_production_manager_approval_completes_it(self):
        self.entry(100, waste=2, status='in_progress')
        PrintingCompletionRequest.objects.create(job_card=self.card, requested_by=self.user, reason='customer accepted',
                                                 status='approved')
        status = printing_completion_status(self.card)
        self.assertEqual((status['complete'], status['rule']), (True, 'pm_approved'))
        self.assertEqual(pending_printing_pcs(self.card), 0)

    def test_a_job_that_never_reached_the_final_pass_is_not_complete(self):
        self.card.planning_job.print_passes = 2                                    # a two-pass job
        self.card.planning_job.save()
        Production.objects.create(job_card=self.card, entry_type='printing', date='2026-10-27', shift='A',
                                  print_pass_number=1, output_sheets=0, waste_sheets=200, impressions=200, intermediate_pass=True)
        self.assertFalse(printing_completion_status(self.card)['final_pass_started'])


class CompletedEntryGuardTests(TestCase):
    def setUp(self):
        self.supervisor = get_user_model().objects.create_user('sup2', password='pw12345678')
        self.manager = get_user_model().objects.create_user('pm_user', password='pw12345678', first_name='Pat')
        profile, _ = UserProfile.objects.get_or_create(user=self.manager)
        profile.role = 'production_manager'
        profile.save()
        self.card = make_card(make_job('PC2', 1000, sku='SKU-GUARD'))
        self.card.total_sheet_quantity = 260
        self.card.production_tolerance_percent = 5
        self.card.save()

    def entry(self, good, waste=0, status='completed'):
        return Production.objects.create(
            job_card=self.card, entry_type='printing', date='2026-10-27', shift='A', print_pass_number=1,
            output_sheets=good, waste_sheets=waste, impressions=good + waste, status=status, created_by=self.supervisor,
        )

    def test_a_finished_job_marked_completed_passes_untouched(self):
        entry = self.entry(248, waste=12)
        self.assertIsNone(guard_completed_entry(self.card, entry, self.supervisor, ''))
        entry.refresh_from_db()
        self.assertEqual(entry.status, 'completed')

    def test_short_job_marked_completed_is_refused_without_a_reason(self):
        entry = self.entry(100, waste=2)
        with self.assertRaises(ValueError) as caught:
            guard_completed_entry(self.card, entry, self.supervisor, '   ')
        self.assertIn('600 pcs', str(caught.exception))
        self.assertIn('Production Manager', str(caught.exception))

    def test_short_job_with_a_reason_is_kept_in_progress_and_sent_for_approval(self):
        entry = self.entry(100, waste=2)
        request = guard_completed_entry(self.card, entry, self.supervisor, 'Customer cut the order')
        entry.refresh_from_db()
        self.assertEqual(entry.status, 'in_progress')
        self.assertEqual((request.status, request.shortfall_pcs, request.reason), ('pending', 600, 'Customer cut the order'))
        self.assertEqual(request.requested_by, self.supervisor)
        # asking again for the same job updates the open request instead of piling up duplicates
        again = guard_completed_entry(self.card, self.entry(110, waste=2), self.supervisor, 'Still short')
        self.assertEqual(again.pk, request.pk)
        self.assertEqual(PrintingCompletionRequest.objects.count(), 1)

    def test_an_intermediate_pass_marked_completed_is_not_guarded(self):
        self.card.planning_job.print_passes = 2                                    # pass 1 of 2 is not the final pass
        self.card.planning_job.save()
        entry = Production.objects.create(
            job_card=self.card, entry_type='printing', date='2026-10-27', shift='A', print_pass_number=1,
            output_sheets=0, waste_sheets=5, impressions=5, intermediate_pass=True, status='completed',
        )
        self.assertIsNone(guard_completed_entry(self.card, entry, self.supervisor, ''))

    def test_manager_approval_completes_the_entry_and_the_wip_status(self):
        entry = self.entry(100, waste=2)
        request = guard_completed_entry(self.card, entry, self.supervisor, 'Customer cut the order')
        evaluate_and_update_job_wip_status(self.card)
        self.assertEqual(JobCardWipStatus.objects.get(job_card=self.card).status.name, 'Partial Printing')

        self.client.force_login(self.manager)
        page = self.client.get(reverse('production_approvals'), {'tab': 'printing'})
        self.assertEqual(page.status_code, 200)
        self.assertIn('Customer cut the order', page.content.decode())
        response = self.client.post(reverse('review_printing_completion', args=[request.id]),
                                    {'action': 'approve', 'review_note': 'ok'})
        self.assertEqual(response.status_code, 302)
        request.refresh_from_db()
        entry.refresh_from_db()
        self.assertEqual((request.status, request.reviewed_by), ('approved', self.manager))
        self.assertEqual(entry.status, 'completed')
        self.assertEqual(JobCardWipStatus.objects.get(job_card=self.card).status.name, 'Printing Completed')
        self.assertEqual(pending_printing_pcs(self.card), 0)

    def test_rejection_keeps_printing_open_and_needs_a_note(self):
        entry = self.entry(100, waste=2)
        request = guard_completed_entry(self.card, entry, self.supervisor, 'please')
        self.client.force_login(self.manager)
        self.client.post(reverse('review_printing_completion', args=[request.id]), {'action': 'reject'})
        request.refresh_from_db()
        self.assertEqual(request.status, 'pending')                                    # a rejection must explain itself
        self.client.post(reverse('review_printing_completion', args=[request.id]), {'action': 'reject', 'review_note': 'print the rest'})
        request.refresh_from_db()
        entry.refresh_from_db()
        self.assertEqual((request.status, entry.status), ('rejected', 'in_progress'))
        self.assertFalse(is_printing_complete(self.card))

    def test_only_managers_can_review(self):
        entry = self.entry(100, waste=2)
        request = guard_completed_entry(self.card, entry, self.supervisor, 'please')
        self.client.force_login(self.supervisor)
        page = self.client.get(reverse('production_approvals'), {'tab': 'review'})        # not a reviewer: sees only their own requests
        self.assertNotIn('Approve: printing is complete', page.content.decode())
        self.client.post(reverse('review_printing_completion', args=[request.id]), {'action': 'approve'})
        request.refresh_from_db()
        self.assertEqual(request.status, 'pending')


class PendingWorkUsesTheSharedRuleTests(TestCase):
    def test_a_job_whose_sheets_are_used_up_leaves_the_pending_printing_list(self):
        from reports.services import build_pending_work_context

        user = get_user_model().objects.create_superuser('pw_admin', password='pw12345678')
        done = make_card(make_job('PC3', 1000, sku='SKU-PW-DONE'))
        short = make_card(make_job('PC4', 1000, sku='SKU-PW-SHORT'))
        for card, good, waste in ((done, 248, 12), (short, 100, 2)):
            card.total_sheet_quantity = 260
            card.save()
            Production.objects.create(job_card=card, entry_type='printing', date=datetime.date.today(), shift='A',
                                      print_pass_number=1, output_sheets=good, waste_sheets=waste, impressions=good + waste)
        request = RequestFactory().get('/reports/pending-work/')
        request.user = user
        rows = {r['job_card_no']: r for r in build_pending_work_context(request)['printing_rows']}
        self.assertNotIn(done.job_card_no, rows)
        self.assertIn(short.job_card_no, rows)
        self.assertEqual(rows[short.job_card_no]['pending_qty'], 600)


class EntryFormFlowTests(TestCase):
    """The real printing entry form: a supervisor marking a short job completed."""

    def setUp(self):
        from core.models import Operator

        from core.models import Permission, UserPermissionOverride

        self.supervisor = get_user_model().objects.create_superuser('form_sup', password='pw12345678')
        UserProfile.objects.get_or_create(user=self.supervisor)
        permission, _ = Permission.objects.get_or_create(code='action.edit_production', defaults={'name': 'Edit production'})
        UserPermissionOverride.objects.get_or_create(user=self.supervisor, permission=permission, defaults={'granted': True})
        self.client.force_login(self.supervisor)
        self.operator = Operator.objects.create(name='Op One')
        self.card = make_card(make_job('PC5', 1000, sku='SKU-FORM'))
        self.card.total_sheet_quantity = 260
        self.card.production_tolerance_percent = 5
        self.card.save()

    def post(self, good, waste=0, remarks='', status='completed'):
        data = {
            'job_card': self.card.id, 'operator': self.operator.id, 'shift': 'A', 'date': '2026-10-27',
            'impressions': good + waste, 'output_sheets': good, 'waste_sheets': waste, 'print_pass_number': 1,
            'status': status, 'remarks': remarks, 'planned_time': 0, 'run_time': 0,
        }
        if waste:
            data['waste_reason'] = 'paper_jam'
        return self.client.post(reverse('printing_production_entry'), data, follow=True)

    def messages(self, response):
        return [str(m) for m in response.context['messages']]

    def test_finished_job_marked_completed_saves_normally(self):
        response = self.post(248, waste=12)
        self.assertEqual(Production.objects.filter(job_card=self.card, status='completed').count(), 1)
        self.assertFalse(PrintingCompletionRequest.objects.exists())
        self.assertTrue(any('saved successfully' in m for m in self.messages(response)))
        self.assertEqual(JobCardWipStatus.objects.get(job_card=self.card).status.name, 'Printing Completed')

    def test_short_job_marked_completed_without_a_reason_shows_the_error_and_saves_nothing(self):
        response = self.post(100, waste=2)
        self.assertEqual(Production.objects.filter(job_card=self.card).count(), 0)
        self.assertTrue(any('not finished' in m and 'Production Manager' in m for m in self.messages(response)))

    def test_short_job_marked_completed_with_a_reason_goes_to_the_production_manager(self):
        response = self.post(100, waste=2, remarks='Customer reduced the order')
        entry = Production.objects.get(job_card=self.card)
        self.assertEqual(entry.status, 'in_progress')
        request = PrintingCompletionRequest.objects.get()
        self.assertEqual((request.status, request.reason, request.shortfall_pcs), ('pending', 'Customer reduced the order', 600))
        self.assertTrue(any('sent to the Production Manager for approval' in m for m in self.messages(response)))
        self.assertEqual(JobCardWipStatus.objects.get(job_card=self.card).status.name, 'Partial Printing')

    def test_saving_a_short_job_as_in_progress_needs_no_approval(self):
        self.post(100, waste=2, status='in_progress')
        self.assertEqual(Production.objects.get(job_card=self.card).status, 'in_progress')
        self.assertFalse(PrintingCompletionRequest.objects.exists())


class RecalculateCommandTests(TestCase):
    def test_moves_finished_jobs_off_partial_printing_and_leaves_manual_overrides(self):
        from io import StringIO

        from django.core.management import call_command

        from production.wip_service import update_wip_status_for_job

        def make(number, good, waste):
            card = make_card(make_job(number, 1000, sku='SKU-' + number))
            card.total_sheet_quantity = 260
            card.save()
            Production.objects.create(job_card=card, entry_type='printing', date='2026-10-27', shift='A', print_pass_number=1,
                                      output_sheets=good, waste_sheets=waste, impressions=good + waste)
            return card

        stuck = make('PC6', 248, 12)                    # all sheets used: finished, but stored as Partial Printing
        short = make('PC7', 100, 2)                     # genuinely short
        manual = make('PC8', 248, 12)                   # finished, but a supervisor set the status by hand
        for card in (stuck, short):
            update_wip_status_for_job(card, 'Partial Printing', is_manual=False, force=True)
        update_wip_status_for_job(manual, 'Partial Printing', is_manual=True, force=True)

        out = StringIO()
        call_command('recalculate_printing_wip', stdout=out)                       # dry run
        self.assertIn('Would move 1 job', out.getvalue())
        self.assertIn(stuck.job_card_no, out.getvalue())
        self.assertIn(manual.job_card_no, out.getvalue())                           # listed as left alone
        self.assertEqual(JobCardWipStatus.objects.get(job_card=stuck).status.name, 'Partial Printing')

        call_command('recalculate_printing_wip', '--apply', stdout=out)
        name = lambda card: JobCardWipStatus.objects.get(job_card=card).status.name
        self.assertEqual(name(stuck), 'Printing Completed')
        self.assertEqual(name(short), 'Partial Printing')
        self.assertEqual(name(manual), 'Partial Printing')                         # the supervisor's choice stands


class ProductionApprovalsInboxTests(TestCase):
    """Edit overrides, printing completions and My requests live on one page; the old pages lead there."""

    def setUp(self):
        from core.models import EditOverrideRequest

        self.supervisor = get_user_model().objects.create_user('inbox_sup', password='pw12345678', first_name='Sue')
        self.manager = get_user_model().objects.create_user('inbox_pm', password='pw12345678')
        profile, _ = UserProfile.objects.get_or_create(user=self.manager)
        profile.role = 'production_manager'
        profile.save()
        self.card = make_card(make_job('PC9', 1000, sku='SKU-INBOX'))
        self.card.total_sheet_quantity = 260
        self.card.save()
        entry = Production.objects.create(job_card=self.card, entry_type='printing', date='2026-10-27', shift='A',
                                          print_pass_number=1, output_sheets=100, waste_sheets=2, impressions=102, status='completed')
        self.printing_request = guard_completed_entry(self.card, entry, self.supervisor, 'Customer cut the order')
        self.override = EditOverrideRequest.objects.create(
            entity_type='production', record_id=entry.id, record_label='Printing record #1', requested_by=self.supervisor,
            reason='Wrong impressions typed',
        )

    def test_reviewer_sees_both_kinds_in_one_inbox_with_decision_buttons(self):
        self.client.force_login(self.manager)
        html = self.client.get(reverse('production_approvals')).content.decode()                  # default tab: To review
        self.assertIn('Customer cut the order', html)
        self.assertIn('Wrong impressions typed', html)
        self.assertIn('Edit override', html)
        self.assertIn('Printing completion', html)
        self.assertIn('Approve: allow the edit', html)
        self.assertIn('Approve: printing is complete', html)
        only_printing = self.client.get(reverse('production_approvals'), {'tab': 'printing'}).content.decode()
        self.assertIn('Customer cut the order', only_printing)
        self.assertNotIn('Wrong impressions typed', only_printing)

    def test_both_kinds_are_decided_from_the_inbox_and_land_back_on_it(self):
        self.client.force_login(self.manager)
        response = self.client.post(reverse('review_override_request', args=[self.override.id]),
                                    {'action': 'approve', 'review_note': 'ok'})
        self.assertEqual(response.status_code, 302)
        self.assertIn(reverse('production_approvals'), response['Location'])
        self.override.refresh_from_db()
        self.assertEqual((self.override.status, self.override.reviewed_by), ('approved', self.manager))
        response = self.client.post(reverse('review_printing_completion', args=[self.printing_request.id]),
                                    {'action': 'approve', 'review_note': 'ok'})
        self.assertIn(reverse('production_approvals'), response['Location'])
        self.printing_request.refresh_from_db()
        self.assertEqual(self.printing_request.status, 'approved')

    def test_my_requests_shows_the_requesters_own_items_of_both_kinds(self):
        self.client.force_login(self.supervisor)
        page = self.client.get(reverse('production_approvals'))                                    # not a reviewer: lands on My requests
        html = page.content.decode()
        self.assertEqual(page.context['tab'], 'mine')
        self.assertIn('Customer cut the order', html)
        self.assertIn('Wrong impressions typed', html)
        self.assertNotIn('Approve: allow the edit', html)                                         # no decision buttons for the requester
        self.assertEqual(len(page.context['waiting']), 2)
        # the review tabs are for reviewers only
        self.assertEqual(self.client.get(reverse('production_approvals'), {'tab': 'review'}).context['tab'], 'mine')

    def test_an_approved_override_shows_as_ready_to_edit_for_the_requester(self):
        self.client.force_login(self.manager)
        self.client.post(reverse('review_override_request', args=[self.override.id]), {'action': 'approve'})
        self.client.force_login(self.supervisor)
        page = self.client.get(reverse('production_approvals'), {'tab': 'mine'})
        self.assertEqual(len(page.context['ready']), 1)
        self.assertIn('Edit now', page.content.decode())

    def test_the_old_pages_lead_to_the_new_inbox(self):
        self.client.force_login(self.manager)
        for name, expected in (('override_requests', 'tab=overrides'), ('my_override_requests', 'tab=mine'),
                               ('printing_completion_requests', reverse('production_approvals'))):
            response = self.client.get(reverse(name))
            self.assertEqual(response.status_code, 302, name)
            self.assertIn(expected, response['Location'], name)

    def test_the_menu_has_one_approvals_tab_with_one_badge(self):
        self.client.force_login(self.manager)
        html = self.client.get(reverse('production_approvals')).content.decode()
        self.assertEqual(html.count('>Approvals'), 1)
        self.assertNotIn('Printing Approvals', html)
        self.assertNotIn('Override Requests', html)
        self.assertIn('2 waiting', html)                                                          # one edit override + one printing completion
