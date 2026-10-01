"""AskAssistantView — synchronous 'ask a question' endpoint used by the
dashboard-wide Ask AI widget. Covers the AISettings gate added alongside
that widget (previously this view ignored chat_assistant_enabled, unlike
the in-room assistant reply path)."""
from __future__ import annotations

from datetime import date
from unittest.mock import patch

from django.contrib.auth import get_user_model
from django.test import TestCase
from django.urls import reverse

from core.models import AISettings, JobCard, Machine

User = get_user_model()


class AskAssistantViewTests(TestCase):
    def setUp(self):
        self.user = User.objects.create_superuser(
            username='ai_widget_tester', email='t@example.com', password='testpass123',
        )
        self.client.force_login(self.user)
        self.url = reverse('assistant-ask')
        AISettings.objects.create(pk=1, ai_enabled=True, chat_assistant_enabled=True)

    def test_requires_question(self):
        response = self.client.post(self.url, data={}, content_type='application/json')
        self.assertEqual(response.status_code, 400)

    @patch('chat.ai_assistant.resolve_and_reply')
    def test_returns_answer_when_enabled(self, mock_resolve):
        mock_resolve.return_value = 'JC-07-26-PP-0701 is in production.'
        response = self.client.post(
            self.url, data={'question': 'status of JC-07-26-PP-0701'}, content_type='application/json',
        )
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json()['answer'], 'JC-07-26-PP-0701 is in production.')
        mock_resolve.assert_called_once()

    @patch('chat.ai_assistant.resolve_and_reply')
    def test_blocked_when_chat_assistant_disabled(self, mock_resolve):
        AISettings.objects.filter(pk=1).update(chat_assistant_enabled=False)
        response = self.client.post(
            self.url, data={'question': 'status of JC-07-26-PP-0701'}, content_type='application/json',
        )
        self.assertEqual(response.status_code, 503)
        mock_resolve.assert_not_called()

    @patch('chat.ai_assistant.resolve_and_reply')
    def test_blocked_when_ai_disabled_globally(self, mock_resolve):
        AISettings.objects.filter(pk=1).update(ai_enabled=False)
        response = self.client.post(
            self.url, data={'question': 'status of JC-07-26-PP-0701'}, content_type='application/json',
        )
        self.assertEqual(response.status_code, 503)
        mock_resolve.assert_not_called()


class ResolveAndReplyPoLookupTests(TestCase):
    """A PO/WO/PR number routinely covers several job cards (one WO funding
    several SKUs), and a planner often has only the bare trailing digits
    ("wo 9828"), not the full "WO-09-2026-09828" — both were previously
    unrecognised: _find_by_po only tried an exact match and only ever
    returned one job card even when several shared the number."""

    def setUp(self):
        AISettings.objects.create(pk=1, ai_enabled=False)  # deterministic: raw facts, no LLM call
        self.machine = Machine.objects.create(name='Ask AI Test Machine')

    def _make_job_card(self, **overrides):
        defaults = dict(
            SKU='SKU-TEST', order_qty=100, status='in_production',
            po_date=date(2026, 1, 1), total_sheet_quantity=10, total_colors=1,
            machine_name=self.machine,
        )
        defaults.update(overrides)
        return JobCard.objects.create(**defaults)

    def test_bare_trailing_digits_resolve_to_full_wo_number(self):
        from chat.ai_assistant import resolve_and_reply

        self._make_job_card(job_card_no='JC-09-26-PP-2822', PO_No='WO-09-2026-09828')
        reply = resolve_and_reply('wo 9828')
        self.assertIn('WO-09-2026-09828', reply)
        self.assertIn('JC-09-26-PP-2822', reply)

    def test_multiple_job_cards_under_one_wo_are_all_listed(self):
        from chat.ai_assistant import resolve_and_reply

        self._make_job_card(job_card_no='JC-09-26-PP-2822', PO_No='WO-09-2026-09828')
        self._make_job_card(job_card_no='JC-09-26-PP-2823', PO_No='WO-09-2026-09828')
        self._make_job_card(job_card_no='JC-09-26-PP-2824', PO_No='WO-09-2026-09828')
        reply = resolve_and_reply('wo 9828')
        self.assertIn('Job cards under this number: 3', reply)
        self.assertIn('JC-09-26-PP-2822', reply)
        self.assertIn('JC-09-26-PP-2823', reply)
        self.assertIn('JC-09-26-PP-2824', reply)

    def test_single_job_card_under_a_wo_gets_full_detail_reply(self):
        from chat.ai_assistant import resolve_and_reply

        self._make_job_card(job_card_no='JC-09-26-PP-2822', PO_No='WO-09-2026-09828', plate_set_no='SET-77')
        reply = resolve_and_reply('wo 9828')
        self.assertIn('Plate Set No: SET-77', reply)

    def test_exact_po_number_still_matches(self):
        from chat.ai_assistant import resolve_and_reply

        self._make_job_card(job_card_no='JC-09-26-PP-2822', PO_No='WO-09-2026-09828')
        reply = resolve_and_reply('po no WO-09-2026-09828')
        self.assertIn('JC-09-26-PP-2822', reply)

    def test_unmatched_numeric_wo_reports_not_found(self):
        from chat.ai_assistant import resolve_and_reply

        reply = resolve_and_reply('wo 404404')
        self.assertIn("couldn't find", reply)

    def test_non_numeric_value_does_not_fall_back_to_partial_match(self):
        from chat.ai_assistant import resolve_and_reply

        self._make_job_card(job_card_no='JC-09-26-PP-2822', PO_No='CUSTOMER-PO-ABC123')
        reply = resolve_and_reply('po no ABC123')
        self.assertIn("couldn't find", reply)


class ResolveAndReplyFlexibilityTests(TestCase):
    """Further shorthand/fuzziness gaps found alongside the WO fix: "job
    card 105" / "job 105" wording wasn't recognised (only "jc 105"), and SKU
    / raw-material-SKU lookups required an exact match with no partial-match
    fallback despite both being long, easy-to-mistype codes in practice."""

    def setUp(self):
        AISettings.objects.create(pk=1, ai_enabled=False)  # deterministic: raw facts, no LLM call
        self.machine = Machine.objects.create(name='Ask AI Flex Test Machine')

    def _make_job_card(self, **overrides):
        defaults = dict(
            SKU='SKU-TEST', order_qty=100, status='in_production',
            po_date=date(2026, 1, 1), total_sheet_quantity=10, total_colors=1,
            machine_name=self.machine,
        )
        defaults.update(overrides)
        return JobCard.objects.create(**defaults)

    def test_job_card_wording_resolves_by_serial(self):
        from chat.ai_assistant import resolve_and_reply

        self._make_job_card(job_card_no='JC-09-26-PP-2822')
        self.assertIn('JC-09-26-PP-2822', resolve_and_reply('job card 2822'))
        self.assertIn('JC-09-26-PP-2822', resolve_and_reply('job card no 2822'))

    def test_bare_job_wording_resolves_by_serial(self):
        from chat.ai_assistant import resolve_and_reply

        self._make_job_card(job_card_no='JC-09-26-PP-2822')
        self.assertIn('JC-09-26-PP-2822', resolve_and_reply('status of job 2822'))

    def test_sku_partial_match_resolves_when_unique(self):
        from chat.ai_assistant import resolve_and_reply

        self._make_job_card(job_card_no='JC-09-26-PP-2823', SKU='INSERTCARD-UTBATHTOWEL100150-MIG-UK-EU-SEP2026')
        reply = resolve_and_reply('sku INSERTCARD-UTBATHTOWEL100150')
        self.assertIn('JC-09-26-PP-2823', reply)

    def test_sku_partial_match_lists_candidates_when_ambiguous(self):
        from chat.ai_assistant import resolve_and_reply

        self._make_job_card(job_card_no='JC-09-26-PP-2822', SKU='INSERTCARD-UTSLTOWELSET-MIG-UK-EU-SEP2026')
        self._make_job_card(job_card_no='JC-09-26-PP-2823', SKU='INSERTCARD-UTBATHTOWEL100150-MIG-UK-EU-SEP2026')
        reply = resolve_and_reply('sku INSERTCARD-UT')
        self.assertIn('which one did you mean', reply)
        self.assertIn('INSERTCARD-UTSLTOWELSET-MIG-UK-EU-SEP2026', reply)
        self.assertIn('INSERTCARD-UTBATHTOWEL100150-MIG-UK-EU-SEP2026', reply)
        # Must NOT have silently picked one and returned a full detail reply.
        self.assertNotIn('Job Card:', reply)

    def test_sku_exact_match_still_wins_over_partial(self):
        from chat.ai_assistant import resolve_and_reply

        self._make_job_card(job_card_no='JC-09-26-PP-2822', SKU='INSERTCARD-UT')
        self._make_job_card(job_card_no='JC-09-26-PP-2823', SKU='INSERTCARD-UTBATHTOWEL100150-MIG-UK-EU-SEP2026')
        reply = resolve_and_reply('sku INSERTCARD-UT')
        self.assertIn('JC-09-26-PP-2822', reply)
        self.assertNotIn('which one did you mean', reply)

    def test_raw_material_sku_partial_match_resolves_when_unique(self):
        from chat.ai_assistant import resolve_and_reply
        from core.models import Material
        from supply_chain.models import RawMaterialSku

        material = Material.objects.create(name='Rubber Roller Covering')
        RawMaterialSku.objects.create(
            sku='RUBBER COVERING OF SM-74 ROLLER SIZE DIA 75MM',
            material=material, purchase_sheet_size='N/A',
        )
        reply = resolve_and_reply('material RUBBER COVERING OF SM-74')
        self.assertIn('Rubber Roller Covering', reply)

    def test_raw_material_sku_partial_match_lists_candidates_when_ambiguous(self):
        from chat.ai_assistant import resolve_and_reply
        from core.models import Material
        from supply_chain.models import RawMaterialSku

        material = Material.objects.create(name='Rubber Roller Covering')
        RawMaterialSku.objects.create(
            sku='RUBBER COVERING SM-74 75MM', material=material, purchase_sheet_size='75MM',
        )
        RawMaterialSku.objects.create(
            sku='RUBBER COVERING SM-52 60MM', material=material, purchase_sheet_size='60MM',
        )
        reply = resolve_and_reply('material RUBBER COVERING')
        self.assertIn('which one did you mean', reply)
