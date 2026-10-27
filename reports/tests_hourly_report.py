import datetime

from django.contrib.auth import get_user_model
from django.test import TestCase
from django.urls import reverse
from django.utils import timezone

from core.models import ChangeLog, Dispatch, Machine, Production
from planning.tests_merge import make_card, make_job
from reports.hourly_report import build_hourly_report, build_text, render_image, shift_for
from reports.models import HourlyReportSnapshot


def local(year, month, day, hour, minute=0):
    return timezone.make_aware(datetime.datetime(year, month, day, hour, minute))


class ShiftTests(TestCase):
    def test_day_shift_is_8am_to_8pm_and_the_night_shift_belongs_to_the_day_it_started(self):
        d = datetime.date
        self.assertEqual(shift_for(local(2026, 10, 27, 8, 0)), ('A', d(2026, 10, 27)))
        self.assertEqual(shift_for(local(2026, 10, 27, 19, 59)), ('A', d(2026, 10, 27)))
        self.assertEqual(shift_for(local(2026, 10, 27, 20, 0)), ('B', d(2026, 10, 27)))
        self.assertEqual(shift_for(local(2026, 10, 27, 23, 30)), ('B', d(2026, 10, 27)))
        self.assertEqual(shift_for(local(2026, 10, 28, 3, 15)), ('B', d(2026, 10, 27)))     # after midnight: still the 27th's night
        self.assertEqual(shift_for(local(2026, 10, 28, 7, 59)), ('B', d(2026, 10, 27)))


class HourlyReportTests(TestCase):
    def setUp(self):
        self.user = get_user_model().objects.create_user('supervisor', password='pw12345678', first_name='Sam', last_name='Super')
        self.machine_a, _ = Machine.objects.get_or_create(name='GTO 1A')
        self.machine_b, _ = Machine.objects.get_or_create(name='GTO 1B')
        Machine.objects.get_or_create(name='GTO 2A', defaults={'machine_type': 'offset_printing'})   # never used in the fixture: reported as idle
        Machine.objects.get_or_create(name='Lamination', defaults={'machine_type': 'other'})         # not a press: never listed as idle
        self.now = local(2026, 10, 27, 14, 0)                                        # Shift A of the 27th
        self.date = datetime.date(2026, 10, 27)
        self.card = make_card(make_job('HR1', 5000, sku='SKU-HOURLY-ONE'))

    def printing(self, machine, impressions, shift='A', date=None, **extra):
        return Production.objects.create(
            job_card=self.card, entry_type='printing', date=date or self.date, shift=shift, machine=machine,
            output_sheets=impressions, impressions=impressions, **extra,
        )

    def logged_at(self, obj, when):
        type(obj).objects.filter(pk=obj.pk).update(created_at=when)

    def test_printing_impressions_by_machine_for_the_shift_and_the_day(self):
        self.printing(self.machine_a, 1000)
        self.printing(self.machine_a, 500)
        self.printing(self.machine_b, 300)
        self.printing(self.machine_b, 7000, shift='B')                               # night shift of the same day
        self.printing(self.machine_b, 9999, date=self.date - datetime.timedelta(days=1))   # another day
        report = build_hourly_report(self.now)
        printing = report['printing']
        self.assertEqual(printing['shift_impressions'], 1800)
        self.assertEqual(printing['day_impressions'], 8800)
        self.assertEqual([(m['name'], m['impressions']) for m in printing['machines']], [('GTO 1A', 1500), ('GTO 1B', 300)])
        self.assertIn('GTO 2A', printing['idle_machines'])
        self.assertNotIn('GTO 1A', printing['idle_machines'])
        self.assertNotIn('Lamination', printing['idle_machines'])

    def test_combined_layout_split_entries_do_not_count_impressions_twice(self):
        lead = self.printing(self.machine_a, 1000)
        Production.objects.filter(pk=self.printing(self.machine_a, 250).pk).update(merge_parent=lead)    # a member's copy
        self.assertEqual(build_hourly_report(self.now)['printing']['shift_impressions'], 1000)

    def test_packing_dispatch_and_releases_count_what_was_logged_since_8am(self):
        before, after = local(2026, 10, 27, 7, 0), local(2026, 10, 27, 9, 30)
        old_pack, new_pack = Production.objects.bulk_create([
            Production(job_card=self.card, entry_type='packing', date=self.date, shift='A', packing_qty=111, sorting_waste_qty=1),
            Production(job_card=self.card, entry_type='packing', date=self.date - datetime.timedelta(days=1), shift='A',
                       packing_qty=800, sorting_waste_qty=12),
        ])
        self.logged_at(old_pack, before)                                              # logged before 8 am: yesterday's work
        self.logged_at(new_pack, after)                                               # logged this morning for yesterday: counts today
        old_dc, new_dc = Dispatch.objects.bulk_create([
            Dispatch(job_card=self.card, dc_no='DC-OLD', dispatch_date=self.date, dispatch_qty=50),
            Dispatch(job_card=self.card, dc_no='DC-NEW', dispatch_date=self.date, dispatch_qty=400),
        ])
        Dispatch.objects.filter(dc_no='DC-OLD').update(created_at=before)
        Dispatch.objects.filter(dc_no='DC-NEW').update(created_at=after)
        old_log = ChangeLog.objects.create(entity_type='job_card', record_id=self.card.id, record_label='x', action='release')
        ChangeLog.objects.filter(pk=old_log.pk).update(created_at=before)
        other = make_card(make_job('HR2', 2500, sku='SKU-HOURLY-TWO'))
        new_log = ChangeLog.objects.create(entity_type='job_card', record_id=other.id, record_label='y', action='release')
        ChangeLog.objects.filter(pk=new_log.pk).update(created_at=after)

        report = build_hourly_report(self.now)
        self.assertEqual((report['packing']['packed_pcs'], report['packing']['sorting_waste_pcs']), (800, 12))
        self.assertEqual((report['dispatch']['pcs'], report['dispatch']['dcs']), (400, 1))
        self.assertEqual(report['planning']['released_jobs'], 1)
        self.assertEqual(report['planning']['released'][0]['jc_no'], other.job_card_no)
        self.assertEqual(report['planning']['released'][0]['qty_pcs'], 2500)

    def test_since_last_report_compares_with_the_latest_shared_snapshot_of_the_day(self):
        self.printing(self.machine_a, 1000)
        first = build_hourly_report(local(2026, 10, 27, 13, 0))
        self.assertIsNone(first['delta'])
        snapshot = HourlyReportSnapshot.objects.create(business_date=self.date, shift='A', shared_by=self.user,
                                                       payload={'totals': first['totals']})
        HourlyReportSnapshot.objects.filter(pk=snapshot.pk).update(taken_at=local(2026, 10, 27, 13, 0))
        self.printing(self.machine_b, 400)
        second = build_hourly_report(self.now)
        self.assertEqual(second['delta']['since'], '13:00')
        self.assertEqual(second['delta']['impressions'], 400)
        self.assertIn('Since last report (13:00)', second['text'])

    def test_text_is_a_compact_whatsapp_message(self):
        self.printing(self.machine_a, 1500)
        report = build_hourly_report(self.now)
        text = build_text(report, shared_by='Sam Super')
        self.assertTrue(text.startswith('*HOURLY PRODUCTION REPORT*'))
        for expected in ('Shift A (8 am - 8 pm)', '*PRINTING - impressions*', '- GTO 1A: 1,500', '*PACKING - pcs',
                         '*DISPATCH - pcs', '*PLANNING - released today*', 'No job released yet', '_Shared by Sam Super_'):
            self.assertIn(expected, text)
        self.assertLess(len(text), 1800)

    def test_image_is_a_png(self):
        self.printing(self.machine_a, 1500)
        png = render_image(build_hourly_report(self.now), shared_by='Sam Super')
        self.assertEqual(png[:8], b'\x89PNG\r\n\x1a\n')
        self.assertGreater(len(png), 5000)


class HourlyReportViewTests(TestCase):
    def setUp(self):
        self.admin = get_user_model().objects.create_superuser('hr_admin', password='pw12345678', first_name='Ann')
        self.client.force_login(self.admin)

    def test_page_image_and_share(self):
        page = self.client.get(reverse('reports:hourly_report'))
        self.assertEqual(page.status_code, 200)
        html = page.content.decode()
        for expected in ('Share on WhatsApp', 'Copy text', 'Download image', 'PRINTING', 'PLANNING'):
            self.assertIn(expected, html)

        image = self.client.get(reverse('reports:hourly_report_image'), {'by': 1})
        self.assertEqual((image.status_code, image['Content-Type']), (200, 'image/png'))

        response = self.client.post(reverse('reports:hourly_report_share'), {'channel': 'whatsapp'})
        data = response.json()
        self.assertTrue(data['ok'])
        self.assertIn('_Shared by Ann_', data['text'])
        snapshot = HourlyReportSnapshot.objects.get(pk=data['snapshot_id'])
        self.assertEqual((snapshot.shared_by, snapshot.channel), (self.admin, 'whatsapp'))
        self.assertIn('totals', snapshot.payload)
        self.assertEqual(self.client.get(reverse('reports:hourly_report_share')).status_code, 405)   # POST only

        history = self.client.get(reverse('reports:hourly_report')).content.decode()
        self.assertIn('Ann', history)                                                # appears in "Shared today"

    def test_needs_sign_in_and_report_access(self):
        self.client.logout()
        self.assertEqual(self.client.get(reverse('reports:hourly_report')).status_code, 302)
        self.assertEqual(self.client.post(reverse('reports:hourly_report_share')).status_code, 401)
        plain = get_user_model().objects.create_user('plain_user', password='pw12345678')
        self.client.force_login(plain)
        self.assertEqual(self.client.get(reverse('reports:hourly_report')).status_code, 403)
        self.assertEqual(self.client.post(reverse('reports:hourly_report_share')).status_code, 403)
