import json
import re

from django.contrib.auth import get_user_model
from django.test import TestCase
from django.urls import reverse

from core.models import Operator, Permission, Production, Sorter, Supervisor, UserPermissionOverride, UserProfile
from planning.tests_merge import make_card, make_job


def repost_payload(response):
    """The JSON the page hands to form_repost.js (what was typed + the field errors)."""
    match = re.search(r'<script id="repost-data" type="application/json">(.*?)</script>', response.content.decode(), re.S)
    return json.loads(match.group(1)) if match else None


class EntryFormErrorTests(TestCase):
    """A failed save must keep what was typed and name exactly the fields to fix (they are shown in red)."""

    def setUp(self):
        user = get_user_model().objects.create_superuser('err_user', password='pw12345678')
        UserProfile.objects.get_or_create(user=user)
        for code in ('action.edit_production',):
            perm, _ = Permission.objects.get_or_create(code=code, defaults={'name': code})
            UserPermissionOverride.objects.get_or_create(user=user, permission=perm, defaults={'granted': True})
        self.client.force_login(user)
        self.operator = Operator.objects.create(name='Op')
        self.sup = Supervisor.objects.create(name='Sup')
        self.sorter = Sorter.objects.create(name='Sorter')
        self.card = make_card(make_job('FE1', 1000, sku='SKU-FE'))
        self.card.total_sheet_quantity = 260
        self.card.save()

    def post(self, **over):
        data = {'job_card': self.card.id, 'operator': self.operator.id, 'supervisor': self.sup.id, 'shift': 'A', 'date': '2026-10-27',
                'impressions': 250, 'output_sheets': 250, 'waste_sheets': 0, 'print_pass_number': 1, 'status': 'in_progress',
                'remarks': 'typed remark', 'planned_time': 0, 'run_time': 60}
        data.update(over)
        return self.client.post(reverse('printing_production_entry'), data)

    def test_the_page_data_is_one_json_object_not_a_double_encoded_string(self):
        payload = repost_payload(self.post(operator=''))
        self.assertIsInstance(payload, dict)
        self.assertEqual(payload['data']['remarks'], 'typed remark')
        self.assertEqual(payload['data']['print_pass_number'], '1')       # a field with "pass" in its name is kept too

    def test_missing_operator_and_supervisor_are_both_named_and_nothing_is_saved(self):
        response = self.post(operator='', supervisor='')
        self.assertEqual(response.status_code, 200)
        self.assertEqual(Production.objects.count(), 0)
        payload = repost_payload(response)
        self.assertEqual(set(payload['errors']), {'operator', 'supervisor'})
        self.assertEqual(payload['data']['impressions'], '250')

    def test_waste_without_a_reason_marks_the_reason_field(self):
        payload = repost_payload(self.post(waste_sheets=10, output_sheets=240))
        self.assertEqual(set(payload['errors']), {'waste_reason'})

    def test_downtime_row_without_a_reason_marks_that_row(self):
        response = self.post(**{'downtime_category[]': ['', 'operator'], 'downtime_minutes_detail[]': ['15', '5'], 'downtime_note[]': ['', '']})
        payload = repost_payload(response)
        self.assertEqual(set(payload['errors']), {'downtime_category[]#0'})
        self.assertEqual(payload['data']['downtime_minutes_detail[]'], ['15', '5'])

    def test_all_problems_are_listed_together(self):
        payload = repost_payload(self.post(operator='', supervisor='', shift='', waste_sheets=5))
        self.assertEqual(set(payload['errors']), {'operator', 'supervisor', 'shift', 'waste_reason'})

    def test_a_complete_entry_still_saves(self):
        response = self.post()
        self.assertEqual(response.status_code, 302)
        self.assertEqual(Production.objects.filter(job_card=self.card).count(), 1)

    def test_packing_without_a_sorter_shows_the_form_instead_of_a_404(self):
        self.card.is_print_job = False
        self.card.status = 'in_production'
        self.card.save(update_fields=['is_print_job', 'status'])
        response = self.client.post(reverse('packing_production_entry'), {'job_card': self.card.id, 'shift': 'A', 'date': '2026-10-27',
                                                                          'packing_qty': 10, 'remarks': 'keep me'})
        self.assertEqual(response.status_code, 200)
        payload = repost_payload(response)
        self.assertEqual(set(payload['errors']), {'sorter'})
        self.assertEqual(payload['data']['remarks'], 'keep me')
