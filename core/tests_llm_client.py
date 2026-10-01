"""core.llm.client health tracking — see AISettings.last_success_at /
last_error_at, read by the Settings > AI page so a dead LLM endpoint doesn't
go unnoticed the way the standby server's misconfigured .env did."""
from __future__ import annotations

from unittest.mock import Mock, patch

from django.test import TestCase

from core.llm.client import call_chat
from core.models import AISettings


class LLMCallOutcomeTrackingTests(TestCase):
    def setUp(self):
        AISettings.objects.create(pk=1, ai_enabled=True)

    def _mock_response(self, content='hello'):
        response = Mock()
        response.raise_for_status = Mock()
        response.json.return_value = {
            'choices': [{'message': {'content': content}}],
            'usage': {'prompt_tokens': 1, 'completion_tokens': 1},
        }
        return response

    @patch('core.llm.client.requests.post')
    def test_successful_call_records_last_success_at(self, mock_post):
        mock_post.return_value = self._mock_response('hello')

        result = call_chat([{'role': 'user', 'content': 'hi'}])

        self.assertEqual(result, 'hello')
        row = AISettings.objects.get(pk=1)
        self.assertIsNotNone(row.last_success_at)
        self.assertIsNone(row.last_error_at)

    @patch('core.llm.client.requests.post')
    def test_failed_call_records_last_error_at_and_message(self, mock_post):
        mock_post.side_effect = ConnectionError('Connection refused')

        result = call_chat([{'role': 'user', 'content': 'hi'}])

        self.assertIsNone(result)
        row = AISettings.objects.get(pk=1)
        self.assertIsNotNone(row.last_error_at)
        self.assertIn('Connection refused', row.last_error_message)
        self.assertIsNone(row.last_success_at)

    @patch('core.llm.client.requests.post')
    def test_success_after_earlier_failure_updates_both_timestamps(self, mock_post):
        mock_post.side_effect = ConnectionError('Connection refused')
        call_chat([{'role': 'user', 'content': 'hi'}])
        first_error_at = AISettings.objects.get(pk=1).last_error_at
        self.assertIsNotNone(first_error_at)

        mock_post.side_effect = None
        mock_post.return_value = self._mock_response('hello')
        call_chat([{'role': 'user', 'content': 'hi'}])

        row = AISettings.objects.get(pk=1)
        self.assertIsNotNone(row.last_success_at)
        self.assertEqual(row.last_error_at, first_error_at)

    @patch('core.llm.client.requests.post')
    def test_disabled_ai_skips_call_and_does_not_touch_health_fields(self, mock_post):
        AISettings.objects.filter(pk=1).update(ai_enabled=False)

        result = call_chat([{'role': 'user', 'content': 'hi'}])

        self.assertIsNone(result)
        mock_post.assert_not_called()
        row = AISettings.objects.get(pk=1)
        self.assertIsNone(row.last_success_at)
        self.assertIsNone(row.last_error_at)
