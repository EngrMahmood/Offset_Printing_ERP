import datetime
from decimal import Decimal

from django.contrib.auth import get_user_model
from django.test import TestCase
from django.urls import reverse
from unittest import mock

from core.models import Production, UserProfile
from planning.merge_refresh import refresh_group, stale_changes
from planning.models import MergeGroup, MergeGroupItem
from planning.tests_merge import make_card, make_job


class MergeRefreshTests(TestCase):
    """MG-2026-0061 in miniature: accepted while 'carried forward' stock reduced the net quantities, stock later cleared."""

    def setUp(self):
        self.big = make_job('RF-BIG', 499220, stock_qty=Decimal('179696'), ups=8)      # net 319,524 when accepted
        self.small = make_job('RF-SMALL', 70000, stock_qty=Decimal('23900'), ups=8)    # net 46,100 when accepted
        self.group = MergeGroup.objects.create(
            code='MG-TEST-1', status='layout_done', total_sheet_ups=8, run_sheets=46100, lead_job=self.big, combined_wastage_sheets=5)
        MergeGroupItem.objects.create(merge_group=self.group, planning_job=self.big, allocated_ups=7, is_lead=True,
                                      planned_produced_qty=322700, net_qty=319524, overage_pct=Decimal('0.99'))
        MergeGroupItem.objects.create(merge_group=self.small, planning_job=self.small, allocated_ups=1,
                                      planned_produced_qty=46100, net_qty=46100, overage_pct=Decimal('0')) if False else \
            MergeGroupItem.objects.create(merge_group=self.group, planning_job=self.small, allocated_ups=1,
                                          planned_produced_qty=46100, net_qty=46100, overage_pct=Decimal('0'))

    def clear_stock(self):
        for job in (self.big, self.small):
            job.stock_qty = 0
            job.save(update_fields=['stock_qty'])

    def test_group_matches_while_nothing_changed(self):
        self.assertEqual(stale_changes(self.group), [])
        self.assertEqual(refresh_group(self.group)[0], 'unchanged')

    def test_clearing_carried_forward_stock_makes_the_group_stale_and_refresh_fixes_it(self):
        self.clear_stock()
        self.assertTrue(any('499,220' in c for c in stale_changes(self.group)))
        status, message = refresh_group(self.group)
        self.assertEqual(status, 'updated', message)
        self.group.refresh_from_db()
        self.assertEqual(self.group.run_sheets, 71318)                         # ceil(499,220 / 7), the longest member run
        big, small = (self.group.items.get(planning_job=j) for j in (self.big, self.small))
        self.assertEqual((big.allocated_ups, small.allocated_ups), (7, 1))     # the approved layout is untouched
        self.assertEqual((big.net_qty, small.net_qty), (499220, 70000))
        self.assertEqual(big.planned_produced_qty, 71318 * 7)
        self.assertEqual(small.planned_produced_qty, 71318)
        self.assertEqual(stale_changes(self.group), [])

    def test_a_group_that_started_printing_is_left_alone(self):
        self.clear_stock()
        Production.objects.bulk_create([Production(job_card=make_card(self.big), entry_type='printing', date=datetime.date(2026, 10, 20),
                                                   shift='A', output_sheets=100, impressions=100, is_active=True)])
        self.assertEqual(refresh_group(self.group)[0], 'locked')
        self.group.refresh_from_db()
        self.assertEqual(self.group.run_sheets, 46100)

    def test_tabs_and_refresh_button(self):
        user = get_user_model().objects.create_user('planner', password='pw12345678')
        UserProfile.objects.get_or_create(user=user)
        self.client.force_login(user)
        patcher = mock.patch.object(UserProfile, '_has_permission', return_value=True)    # role matrix is not seeded in tests
        patcher.start()
        self.addCleanup(patcher.stop)
        self.clear_stock()
        groups = self.client.get(reverse('planning:merge_board'), {'tab': 'groups'})
        self.assertEqual(groups.status_code, 200)
        self.assertContains(groups, 'out of date')
        self.assertContains(self.client.get(reverse('planning:merge_board'), {'tab': 'groups', 'q': 'RF-SMALL'}), 'MG-TEST-1')
        self.assertNotContains(self.client.get(reverse('planning:merge_board'), {'tab': 'groups', 'q': 'nothing-like-this'}), 'MG-TEST-1')
        for tab in ('suggestions', 'blocked', 'closed'):
            self.assertEqual(self.client.get(reverse('planning:merge_board'), {'tab': tab}).status_code, 200, tab)
        self.assertContains(self.client.get(reverse('planning:merge_detail', args=[self.group.id])), 'no longer matches its jobs')
        self.client.post(reverse('planning:merge_refresh'), {'group_id': self.group.id})
        self.group.refresh_from_db()
        self.assertEqual(self.group.run_sheets, 71318)
