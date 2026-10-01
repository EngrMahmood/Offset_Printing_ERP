"""AskAssistantView — synchronous 'ask a question' endpoint used by the
dashboard-wide Ask AI widget. Covers the AISettings gate added alongside
that widget (previously this view ignored chat_assistant_enabled, unlike
the in-room assistant reply path)."""
from __future__ import annotations

from unittest.mock import patch

from django.contrib.auth import get_user_model
from django.test import TestCase
from django.urls import reverse

from core.models import AISettings

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
