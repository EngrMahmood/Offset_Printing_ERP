"""Bot execution: fetch report -> render email -> send -> log.

Every path through run_bot() finishes a BotExecution row. Nothing raises out of
this module into the scheduler thread — a failure is data, not a crash.
"""
from __future__ import annotations

import logging
import traceback

from django.core.mail import EmailMultiAlternatives, get_connection
from django.utils import timezone

from bot import report_adapter, template_engine
from bot.models import (
    STATUS_FAILED,
    STATUS_PENDING,
    STATUS_SENT,
    STATUS_SKIPPED,
    TRIGGER_AUTO,
    TRIGGER_MANUAL,
    TRIGGER_TEST,
    BotExecution,
)
from bot.schedule import calculate_next_run

logger = logging.getLogger(__name__)


def _emails_for_roles(role_keys):
    """Live email addresses of active users holding any of these roles.

    Resolved at send time so staff changes never require editing the bot.
    """
    if not role_keys:
        return []
    from django.contrib.auth import get_user_model

    User = get_user_model()
    return list(
        User.objects.filter(
            is_active=True,
            profile__role__in=role_keys,
        )
        .exclude(email='')
        .values_list('email', flat=True)
    )


def _dedupe(addresses):
    """Case-insensitive dedupe that preserves the configured order."""
    seen = set()
    result = []
    for address in addresses:
        key = (address or '').strip().lower()
        if not key or key in seen:
            continue
        seen.add(key)
        result.append(address.strip())
    return result


def resolve_recipients(bot):
    """(to, cc, bcc) — explicit addresses plus everyone in the selected roles.

    An address already present in To is never repeated in CC/BCC.
    """
    to = _dedupe(list(bot.to_addresses) + _emails_for_roles(bot.role_keys))
    cc = [a for a in _dedupe(bot.cc_addresses) if a.lower() not in {x.lower() for x in to}]
    used = {x.lower() for x in to + cc}
    bcc = [a for a in _dedupe(bot.bcc_addresses) if a.lower() not in used]
    return to, cc, bcc


def _build_ai_summary(bot, headers, labels, rows, payload):
    """Best-effort narration paragraph. Never raises — a summary failure must
    never block or blank out the report email itself."""
    if not bot.use_ai_summary:
        return ''
    from core.models import AISettings
    ai_settings = AISettings.get_solo()
    if not ai_settings.ai_enabled or not ai_settings.bot_summaries_enabled:
        return ''
    try:
        from core.llm.client import call_chat
        from core.llm.prompts import NARRATION_SYSTEM_PROMPT

        facts = report_adapter.build_narration_facts(payload, headers, labels, rows, fallback_title=bot.name)

        return call_chat([
            {'role': 'system', 'content': NARRATION_SYSTEM_PROMPT},
            {'role': 'user', 'content': f'Summarize this report:\n{facts}'},
        ]) or ''
    except Exception:
        logger.warning('AI summary failed for bot %s', bot.code, exc_info=True)
        return ''


def build_email_parts(bot, payload=None, now=None):
    """Render everything the email needs without sending it.

    Shared by run_bot() and the preview screen, so what an admin previews is
    exactly what goes out.
    """
    now = timezone.localtime(now or timezone.now())
    if payload is None:
        payload = report_adapter.fetch_report(bot)

    headers, labels, rows = report_adapter.extract_rows(payload)
    ai_summary = _build_ai_summary(bot, headers, labels, rows, payload)
    context = template_engine.build_context(bot, payload, headers, labels, rows, now, ai_summary=ai_summary)

    return {
        'payload': payload,
        'headers': headers,
        'labels': labels,
        'rows': rows,
        'context': context,
        'subject': template_engine.render_subject(bot, context),
        'html_body': template_engine.render_body(bot, context),
        'text_body': template_engine.render_text_body(bot, context),
        'record_count': len(rows),
    }


def _finish(execution, status, error='', **fields):
    execution.status = status
    execution.error_message = error
    execution.finished_at = timezone.now()
    if execution.started_at:
        execution.duration_seconds = round(
            (execution.finished_at - execution.started_at).total_seconds(), 3
        )
    for name, value in fields.items():
        setattr(execution, name, value)
    execution.save()
    return execution


def run_bot(bot, trigger=TRIGGER_AUTO, actor=None, override_recipients=None):
    """Execute one bot and return its BotExecution row.

    `override_recipients` is used by the test-send path to redirect the mail to
    a single address without touching the bot's configuration.
    """
    execution = BotExecution.objects.create(
        bot=bot,
        trigger=trigger,
        status=STATUS_PENDING,
        triggered_by=actor,
    )

    try:
        parts = build_email_parts(bot)
        record_count = parts['record_count']

        if override_recipients is not None:
            to, cc, bcc = list(override_recipients), [], []
        else:
            to, cc, bcc = resolve_recipients(bot)

        execution.record_count = record_count
        execution.recipients_to = ', '.join(to)
        execution.recipients_cc = ', '.join(cc)
        execution.recipients_bcc = ', '.join(bcc)
        execution.rendered_subject = parts['subject'][:255]
        execution.rendered_body = parts['html_body']

        if record_count == 0 and not bot.send_when_empty and trigger != TRIGGER_TEST:
            return _finish(execution, STATUS_SKIPPED)

        if not (to or cc or bcc):
            return _finish(
                execution,
                STATUS_FAILED,
                error='No recipients resolved. Set Email To, or pick at least one recipient role '
                      'whose users have an email address on file.',
            )

        attachment_name = ''
        attachment = None
        if bot.attach_report and record_count:
            filename, content, mimetype = report_adapter.build_attachment(
                parts['payload'], bot.report_slug, bot.attachment_format
            )
            attachment = (filename, content, mimetype)
            attachment_name = filename

        message = EmailMultiAlternatives(
            subject=parts['subject'],
            body=parts['text_body'],
            # None lets DynamicGmailEmailBackend stamp the authenticated Gmail
            # address, matching how the rest of the ERP sends mail.
            from_email=None,
            to=to,
            cc=cc,
            bcc=bcc,
            connection=get_connection(),
        )
        message.attach_alternative(parts['html_body'], 'text/html')
        if attachment:
            message.attach(*attachment)

        sent = message.send(fail_silently=False)
        if not sent:
            return _finish(
                execution,
                STATUS_FAILED,
                error='The email backend reported that no messages were sent.',
                attachment_name=attachment_name,
            )

        return _finish(execution, STATUS_SENT, attachment_name=attachment_name)

    except Exception as exc:  # noqa: BLE001 — a failed bot must never crash the caller
        logger.exception('Bot %s failed', bot.code)
        return _finish(execution, STATUS_FAILED, error=f'{exc}\n\n{traceback.format_exc()}')

    finally:
        # Bookkeeping reflects every attempt, so the list screen shows the real
        # last-run time even when the run failed.
        bot.last_run_at = timezone.now()
        bot.last_status = execution.status
        if trigger == TRIGGER_AUTO:
            bot.next_run_at = calculate_next_run(bot, bot.last_run_at)
        bot.save(update_fields=['last_run_at', 'last_status', 'next_run_at', 'updated_at'])


def run_bot_manually(bot, actor=None):
    return run_bot(bot, trigger=TRIGGER_MANUAL, actor=actor)


def send_test_email(bot, to_address, actor=None):
    return run_bot(bot, trigger=TRIGGER_TEST, actor=actor, override_recipients=[to_address])


def refresh_next_run(bot):
    """Recompute next_run_at after a config change so the UI never shows a
    stale schedule."""
    bot.next_run_at = calculate_next_run(bot) if bot.is_active else None
    bot.save(update_fields=['next_run_at', 'updated_at'])
    return bot.next_run_at
