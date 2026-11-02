import datetime

from django.contrib.auth import get_user_model
from django.test import TestCase
from django.urls import reverse

from core.models import Machine, Operator, Production
from planning.tests_merge import make_card, make_job
from reports import operator_efficiency as oe

DAY = datetime.date(2026, 10, 20)


class OperatorEfficiencyTests(TestCase):
    def setUp(self):
        self.user = get_user_model().objects.create_superuser('boss', 'b@example.com', 'pw12345678')
        self.machine, _ = Machine.objects.get_or_create(name='GTO 1A')
        self.card = make_card(make_job('OE1', 9000, sku='SKU-OE'))
        self.fast = Operator.objects.create(name='Fast Fahad')
        self.slow = Operator.objects.create(name='Slow Sami')
        self.new = Operator.objects.create(name='New Nadeem')

    def entry(self, operator, day, imp, minutes, waste=0, shift='A', **extra):
        return Production.objects.bulk_create([Production(      # bulk_create skips the sheet-budget check; only the figures matter here
            job_card=self.card, entry_type='printing', date=day, shift=shift, machine=self.machine, operator=operator,
            output_sheets=imp, impressions=imp, run_time=minutes, waste_sheets=waste, **extra)])[0]

    def fixture(self):
        for n in range(6):
            day = DAY + datetime.timedelta(days=n)
            self.entry(self.fast, day, 30000, 360, waste=30)
            self.entry(self.slow, day, 15000, 360, waste=150, shift='B')
        self.entry(self.new, DAY, 40000, 300)                                   # one shift: shown, not ranked

    def test_ranking_uses_peers_on_the_same_machine(self):
        self.fixture()
        result = oe.compute(DAY, DAY + datetime.timedelta(days=10))
        ranked = {r['operator']: r for r in result['rows']}
        self.assertEqual([r['operator'] for r in result['ranked']], ['Fast Fahad', 'Slow Sami'])
        self.assertEqual(ranked['Fast Fahad']['rank'], 1)
        self.assertEqual(ranked['Fast Fahad']['impressions_per_hour'], 5000)
        self.assertGreater(ranked['Fast Fahad']['score'], ranked['Slow Sami']['score'])
        self.assertFalse(ranked['New Nadeem']['ranked'])
        self.assertEqual(ranked['New Nadeem']['rank'], '')

    def test_combined_layout_copies_are_not_counted_twice(self):
        self.fixture()
        lead = Production.objects.filter(operator=self.fast).first()
        Production.objects.filter(pk=self.entry(self.fast, DAY, 9999, 10).pk).update(merge_parent=lead)
        row = next(r for r in oe.compute(DAY, DAY)['rows'] if r['operator'] == 'Fast Fahad')
        self.assertEqual(row['impressions'], 30000)

    def test_page_board_detail_and_exports(self):
        self.fixture()
        self.client.force_login(self.user)
        q = {'period': 'custom', 'date_from': DAY.isoformat(), 'date_to': (DAY + datetime.timedelta(days=10)).isoformat()}
        page = self.client.get(reverse('reports:detail', args=['operator-efficiency']), q)
        self.assertEqual(page.status_code, 200)
        self.assertContains(page, 'Fast Fahad')
        detail = self.client.get(reverse('reports:detail', args=['operator-efficiency']), {**q, 'operator': self.fast.id})
        self.assertContains(detail, 'Day by day')
        board = self.client.get(reverse('reports:operator_board'), q)
        self.assertContains(board, 'OPERATOR RANKING')
        for kind in ('xlsx', 'pdf'):
            r = self.client.get(reverse('reports:reports_api:export_report', args=['operator-efficiency']), {**q, 'type': kind})
            self.assertEqual(r.status_code, 200, kind)


class OperatorFairnessTests(OperatorEfficiencyTests):
    def test_an_operator_who_was_mostly_absent_is_not_ranked_however_fast(self):
        fresh = [Operator.objects.create(name=f'Regular {n}') for n in range(3)]
        for n in range(20):
            for op in fresh:
                self.entry(op, DAY + datetime.timedelta(days=n), 20000, 400)
        for n in range(6):                                               # 6 shifts: over the minimum, but under half of 20
            self.entry(self.fast, DAY + datetime.timedelta(days=n), 40000, 300)
        result = oe.compute(DAY, DAY + datetime.timedelta(days=30))
        by = {r['operator']: r for r in result['rows']}
        self.assertFalse(by['Fast Fahad']['ranked'])
        self.assertTrue(all(by[f'Regular {n}']['ranked'] for n in range(3)))
        self.assertEqual(by['Fast Fahad']['attendance_pct'], 30)         # 6 of the 20 shifts the regulars worked

    def test_changeovers_and_make_ready_are_counted_per_job_started(self):
        other = make_card(make_job('OE2', 9000, sku='SKU-OE2'), prefix='OE2')
        self.entry(self.fast, DAY, 10000, 200, make_ready_time=30)
        Production.objects.bulk_create([Production(job_card=other, entry_type='printing', date=DAY, shift='A', machine=self.machine,
                                                   operator=self.fast, output_sheets=5000, impressions=5000, run_time=100, make_ready_time=60)])
        row = next(r for r in oe.compute(DAY, DAY)['rows'] if r['operator'] == 'Fast Fahad')
        self.assertEqual((row['changeovers'], row['jobs'], row['shifts']), (2, 2, 1))
        self.assertEqual(row['make_ready_min_per_changeover'], 45)
