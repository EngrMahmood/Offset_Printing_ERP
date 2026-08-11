"""Tests for the bot engine.

The report itself is stubbed out — these tests are about the bot's own logic
(scheduling, recipients, rendering, execution logging), not about whether the
reports app produces correct rows.
"""
import datetime
import json
from unittest import mock

from django.contrib.auth import get_user_model
from django.core import mail
from django.test import TestCase, override_settings
from django.utils import timezone

from bot import schedule, services, template_engine
from bot.forms import BotAutomationForm
from bot.models import (
    DEFAULT_BODY_TEMPLATE,
    DEFAULT_SUBJECT_TEMPLATE,
    STATUS_FAILED,
    STATUS_PENDING,
    STATUS_SENT,
    STATUS_SKIPPED,
    TRIGGER_AUTO,
    BotAutomation,
    BotExecution,
)

User = get_user_model()


def make_payload(rows):
    """A minimal engine payload of the shape reports.report_engine returns."""
    return {
        'report': {'slug': 'pending-work', 'title': 'Pending Work'},
        'filters': {'stage': 'not_released'},
        'generated_at': '2026-08-11T08:30:00',
        'data': {
            'export_rows': rows,
            'headers': ['job_card_no', 'sku', 'status'],
            'header_labels': {'job_card_no': 'Job Card', 'sku': 'SKU', 'status': 'Status'},
        },
    }


SAMPLE_ROWS = [
    {'job_card_no': 'JC-001', 'sku': 'SKU-A', 'status': 'qc_approved'},
    {'job_card_no': 'JC-002', 'sku': 'SKU-B', 'status': 'draft'},
]


def make_bot(**overrides):
    defaults = {
        'code': 'TEST_BOT',
        'name': 'Test Bot',
        'report_slug': 'pending-work',
        'report_filters': {'stage': 'not_released'},
        'frequency': 'DAILY',
        'send_time': datetime.time(8, 30),
        'email_to': 'planning@example.com',
        'subject_template': DEFAULT_SUBJECT_TEMPLATE,
        'body_template': DEFAULT_BODY_TEMPLATE,
        'is_active': True,
        'attach_report': False,
    }
    defaults.update(overrides)
    return BotAutomation.objects.create(**defaults)


def local(year, month, day, hour=0, minute=0):
    return timezone.make_aware(
        datetime.datetime(year, month, day, hour, minute), timezone.get_current_timezone()
    )


def make_execution(bot, status, when):
    """BotExecution with a controlled started_at (the field is auto_now_add, so
    it has to be written back) — the due-ness rules are all time-relative."""
    execution = BotExecution.objects.create(bot=bot, trigger=TRIGGER_AUTO, status=status)
    BotExecution.objects.filter(pk=execution.pk).update(started_at=when)
    execution.refresh_from_db()
    return execution


class ScheduleCalendarTests(TestCase):
    """matches_calendar_day / calculate_next_run — no DB writes needed."""

    def test_daily_with_blank_weekdays_matches_every_day(self):
        bot = BotAutomation(frequency='DAILY', weekdays='', send_time=datetime.time(8, 30))
        for day in range(11, 18):  # a full week in Aug 2026
            self.assertTrue(schedule.matches_calendar_day(bot, datetime.date(2026, 8, day)))

    def test_daily_with_weekdays_only_matches_listed_days(self):
        bot = BotAutomation(frequency='DAILY', weekdays='0,2', send_time=datetime.time(8, 30))
        self.assertTrue(schedule.matches_calendar_day(bot, datetime.date(2026, 8, 10)))   # Monday
        self.assertFalse(schedule.matches_calendar_day(bot, datetime.date(2026, 8, 11)))  # Tuesday
        self.assertTrue(schedule.matches_calendar_day(bot, datetime.date(2026, 8, 12)))   # Wednesday

    def test_weekly_defaults_to_monday(self):
        bot = BotAutomation(frequency='WEEKLY', weekdays='', send_time=datetime.time(8, 30))
        self.assertTrue(schedule.matches_calendar_day(bot, datetime.date(2026, 8, 10)))
        self.assertFalse(schedule.matches_calendar_day(bot, datetime.date(2026, 8, 11)))

    def test_monthly_clamps_to_last_day_of_short_month(self):
        bot = BotAutomation(frequency='MONTHLY', day_of_month=31, send_time=datetime.time(8, 30))
        self.assertTrue(schedule.matches_calendar_day(bot, datetime.date(2026, 2, 28)))
        self.assertFalse(schedule.matches_calendar_day(bot, datetime.date(2026, 2, 27)))
        self.assertTrue(schedule.matches_calendar_day(bot, datetime.date(2026, 3, 31)))

    def test_date_window_is_respected(self):
        bot = BotAutomation(
            frequency='DAILY',
            send_time=datetime.time(8, 30),
            start_date=datetime.date(2026, 8, 10),
            end_date=datetime.date(2026, 8, 12),
        )
        self.assertFalse(schedule.matches_calendar_day(bot, datetime.date(2026, 8, 9)))
        self.assertTrue(schedule.matches_calendar_day(bot, datetime.date(2026, 8, 11)))
        self.assertFalse(schedule.matches_calendar_day(bot, datetime.date(2026, 8, 13)))

    def test_next_run_is_later_today_when_send_time_has_not_passed(self):
        bot = BotAutomation(frequency='DAILY', send_time=datetime.time(8, 30))
        self.assertEqual(
            schedule.calculate_next_run(bot, local(2026, 8, 11, 7, 0)),
            local(2026, 8, 11, 8, 30),
        )

    def test_next_run_rolls_to_tomorrow_once_send_time_has_passed(self):
        bot = BotAutomation(frequency='DAILY', send_time=datetime.time(8, 30))
        self.assertEqual(
            schedule.calculate_next_run(bot, local(2026, 8, 11, 9, 0)),
            local(2026, 8, 12, 8, 30),
        )

    def test_next_run_skips_to_the_next_listed_weekday(self):
        bot = BotAutomation(frequency='WEEKLY', weekdays='4', send_time=datetime.time(8, 30))
        # Tuesday 11 Aug 2026 -> Friday 14 Aug 2026
        self.assertEqual(
            schedule.calculate_next_run(bot, local(2026, 8, 11, 9, 0)),
            local(2026, 8, 14, 8, 30),
        )

    def test_next_run_is_none_past_end_date(self):
        bot = BotAutomation(
            frequency='DAILY', send_time=datetime.time(8, 30), end_date=datetime.date(2026, 8, 10)
        )
        self.assertIsNone(schedule.calculate_next_run(bot, local(2026, 8, 11, 9, 0)))


class IsDueTests(TestCase):
    def test_inactive_bot_is_never_due(self):
        bot = make_bot(is_active=False)
        self.assertFalse(schedule.is_due(bot, local(2026, 8, 11, 9, 0)))

    def test_not_due_before_send_time(self):
        bot = make_bot()
        self.assertFalse(schedule.is_due(bot, local(2026, 8, 11, 8, 0)))

    def test_due_after_send_time_with_no_run_today(self):
        bot = make_bot()
        self.assertTrue(schedule.is_due(bot, local(2026, 8, 11, 9, 0)))
        self.assertIn(bot, schedule.due_bots(local(2026, 8, 11, 9, 0)))

    def test_not_due_again_after_a_successful_run_today(self):
        bot = make_bot()
        make_execution(bot, STATUS_SENT, local(2026, 8, 11, 8, 31))
        self.assertFalse(schedule.is_due(bot, local(2026, 8, 11, 9, 0)))

    def test_a_skipped_empty_run_also_closes_the_window(self):
        bot = make_bot()
        make_execution(bot, STATUS_SKIPPED, local(2026, 8, 11, 8, 31))
        self.assertFalse(schedule.is_due(bot, local(2026, 8, 11, 9, 0)))

    def test_yesterdays_run_does_not_block_today(self):
        bot = make_bot()
        make_execution(bot, STATUS_SENT, local(2026, 8, 10, 8, 31))
        self.assertTrue(schedule.is_due(bot, local(2026, 8, 11, 9, 0)))

    def test_not_due_while_a_run_is_in_flight(self):
        bot = make_bot()
        make_execution(bot, STATUS_PENDING, local(2026, 8, 11, 8, 31))
        self.assertFalse(schedule.is_due(bot, local(2026, 8, 11, 9, 0)))

    def test_failure_cooldown_blocks_an_immediate_retry(self):
        bot = make_bot(retry_count=2, retry_interval_minutes=15)
        make_execution(bot, STATUS_FAILED, local(2026, 8, 11, 8, 31))
        self.assertFalse(schedule.is_due(bot, local(2026, 8, 11, 8, 36)))
        # ...but allows one once the cooldown has elapsed.
        self.assertTrue(schedule.is_due(bot, local(2026, 8, 11, 8, 50)))

    def test_retries_stop_after_retry_count_failures(self):
        bot = make_bot(retry_count=1, retry_interval_minutes=15)
        make_execution(bot, STATUS_FAILED, local(2026, 8, 11, 8, 31))
        make_execution(bot, STATUS_FAILED, local(2026, 8, 11, 8, 50))
        self.assertFalse(schedule.is_due(bot, local(2026, 8, 11, 12, 0)))


class RecipientTests(TestCase):
    def setUp(self):
        self.planner = User.objects.create_user(
            'planner1', email='planner1@example.com', password='x'
        )
        self.planner.profile.role = 'planner'
        self.planner.profile.save()

        self.no_email = User.objects.create_user('planner2', email='', password='x')
        self.no_email.profile.role = 'planner'
        self.no_email.profile.save()

        self.inactive = User.objects.create_user(
            'planner3', email='planner3@example.com', password='x', is_active=False
        )
        self.inactive.profile.role = 'planner'
        self.inactive.profile.save()

    def test_roles_expand_to_live_user_emails(self):
        bot = make_bot(email_to='', recipient_roles='planner')
        to, cc, bcc = services.resolve_recipients(bot)
        self.assertEqual(to, ['planner1@example.com'])

    def test_users_without_email_or_inactive_are_excluded(self):
        bot = make_bot(email_to='', recipient_roles='planner')
        to, _, _ = services.resolve_recipients(bot)
        self.assertNotIn('', to)
        self.assertNotIn('planner3@example.com', to)

    def test_explicit_and_role_addresses_are_deduped_case_insensitively(self):
        bot = make_bot(email_to='PLANNER1@example.com, ops@example.com', recipient_roles='planner')
        to, _, _ = services.resolve_recipients(bot)
        self.assertEqual(to, ['PLANNER1@example.com', 'ops@example.com'])

    def test_an_address_in_to_is_not_repeated_in_cc_or_bcc(self):
        bot = make_bot(
            email_to='ops@example.com',
            email_cc='ops@example.com, boss@example.com',
            email_bcc='boss@example.com, audit@example.com',
        )
        to, cc, bcc = services.resolve_recipients(bot)
        self.assertEqual(to, ['ops@example.com'])
        self.assertEqual(cc, ['boss@example.com'])
        self.assertEqual(bcc, ['audit@example.com'])

    def test_newlines_and_semicolons_are_accepted_as_separators(self):
        bot = make_bot(email_to='a@example.com;b@example.com\nc@example.com')
        to, _, _ = services.resolve_recipients(bot)
        self.assertEqual(to, ['a@example.com', 'b@example.com', 'c@example.com'])


class TemplateRenderingTests(TestCase):
    def setUp(self):
        from core.models import Department

        self.user = User.objects.create_superuser('admin1', email='admin@example.com', password='x')
        self.user.first_name = 'Aisha'
        self.user.last_name = 'Khan'
        self.user.save()
        self.user.profile.department = Department.objects.create(name='Production Printing')
        self.user.profile.save()

    def _context(self, bot, rows=SAMPLE_ROWS):
        payload = make_payload(rows)
        headers, labels, rows = services.report_adapter.extract_rows(payload)
        return template_engine.build_context(
            bot, payload, headers, labels, rows, local(2026, 8, 11, 8, 30)
        )

    def test_every_documented_variable_renders(self):
        bot = make_bot(run_as=self.user)
        context = self._context(bot)
        for name, _description in template_engine.SUPPORTED_VARIABLES:
            key = name.strip('{} ')
            self.assertIn(key, context, f'{name} is documented but missing from the context')
            rendered = template_engine.render_template(name, context)
            self.assertNotEqual(rendered.strip(), '', f'{name} rendered empty')

    def test_subject_substitutes_the_date_and_stays_single_line(self):
        bot = make_bot(subject_template='Pending Work - {{date}}\nsecond line')
        subject = template_engine.render_subject(bot, self._context(bot))
        self.assertEqual(subject, 'Pending Work - 11 Aug 2026 second line')

    def test_body_table_contains_a_row_per_record(self):
        bot = make_bot()
        html = template_engine.render_body(bot, self._context(bot))
        self.assertIn('JC-001', html)
        self.assertIn('JC-002', html)
        self.assertIn('Job Card', html)  # the label, not the raw key

    def test_table_truncates_at_max_rows_and_says_so(self):
        bot = make_bot(max_rows_in_body=1)
        html = template_engine.render_body(bot, self._context(bot))
        self.assertIn('JC-001', html)
        self.assertNotIn('JC-002', html)
        self.assertIn('and 1 more row(s)', html)

    def test_empty_report_renders_a_no_records_message(self):
        bot = make_bot()
        html = template_engine.render_body(bot, self._context(bot, rows=[]))
        self.assertIn('No pending records', html)

    def test_text_body_has_no_html_tags(self):
        bot = make_bot()
        text = template_engine.render_text_body(bot, self._context(bot))
        self.assertNotIn('<table', text)
        self.assertNotIn('<p>', text)
        self.assertIn('JC-001', text)

    def test_html_is_escaped_in_cells(self):
        bot = make_bot()
        rows = [{'job_card_no': '<script>x</script>', 'sku': 'A', 'status': 'draft'}]
        html = template_engine.render_body(bot, self._context(bot, rows=rows))
        self.assertNotIn('<script>', html)
        self.assertIn('&lt;script&gt;', html)

    def test_filters_summary_is_human_readable(self):
        self.assertEqual(
            template_engine.summarize_filters({'stage': 'not_released'}), 'Stage: not_released'
        )
        self.assertEqual(template_engine.summarize_filters({}), 'None')


@override_settings(EMAIL_BACKEND='django.core.mail.backends.locmem.EmailBackend')
class RunBotTests(TestCase):
    def setUp(self):
        self.user = User.objects.create_superuser('admin1', email='admin@example.com', password='x')
        mail.outbox = []

    def _patch_report(self, rows=SAMPLE_ROWS):
        return mock.patch.object(
            services.report_adapter, 'fetch_report', return_value=make_payload(rows)
        )

    def test_happy_path_sends_and_logs_a_sent_execution(self):
        bot = make_bot(email_to='planning@example.com', email_cc='pm@example.com')
        with self._patch_report():
            execution = services.run_bot(bot)

        self.assertEqual(execution.status, STATUS_SENT)
        self.assertEqual(execution.record_count, 2)
        self.assertEqual(execution.recipients_to, 'planning@example.com')
        self.assertEqual(execution.recipients_cc, 'pm@example.com')
        self.assertIn('Pending Work', execution.rendered_subject)
        self.assertIn('JC-001', execution.rendered_body)

        self.assertEqual(len(mail.outbox), 1)
        message = mail.outbox[0]
        self.assertEqual(message.to, ['planning@example.com'])
        self.assertEqual(message.cc, ['pm@example.com'])
        self.assertIn('JC-001', message.body)  # plain-text part
        self.assertEqual(message.alternatives[0][1], 'text/html')
        self.assertIn('<table', message.alternatives[0][0])

    def test_bookkeeping_is_updated_after_a_run(self):
        bot = make_bot()
        with self._patch_report():
            services.run_bot(bot)
        bot.refresh_from_db()
        self.assertIsNotNone(bot.last_run_at)
        self.assertEqual(bot.last_status, STATUS_SENT)
        self.assertIsNotNone(bot.next_run_at)

    def test_attachment_is_built_when_enabled(self):
        bot = make_bot(attach_report=True, attachment_format='csv')
        with self._patch_report():
            execution = services.run_bot(bot)
        self.assertEqual(execution.status, STATUS_SENT)
        self.assertTrue(execution.attachment_name.endswith('.csv'))
        self.assertEqual(len(mail.outbox[0].attachments), 1)

    def test_empty_report_is_skipped_and_sends_nothing(self):
        bot = make_bot(send_when_empty=False)
        with self._patch_report(rows=[]):
            execution = services.run_bot(bot)
        self.assertEqual(execution.status, STATUS_SKIPPED)
        self.assertEqual(execution.record_count, 0)
        self.assertEqual(len(mail.outbox), 0)

    def test_empty_report_is_sent_when_send_when_empty_is_on(self):
        bot = make_bot(send_when_empty=True)
        with self._patch_report(rows=[]):
            execution = services.run_bot(bot)
        self.assertEqual(execution.status, STATUS_SENT)
        self.assertEqual(len(mail.outbox), 1)

    def test_no_recipients_fails_without_raising(self):
        bot = make_bot(email_to='', recipient_roles='')
        with self._patch_report():
            execution = services.run_bot(bot)
        self.assertEqual(execution.status, STATUS_FAILED)
        self.assertIn('No recipients', execution.error_message)
        self.assertEqual(len(mail.outbox), 0)

    def test_a_broken_report_is_recorded_as_failed_and_does_not_raise(self):
        bot = make_bot(report_slug='no-such-report')
        execution = services.run_bot(bot)  # real fetch_report -> Http404
        self.assertEqual(execution.status, STATUS_FAILED)
        self.assertTrue(execution.error_message)
        self.assertEqual(len(mail.outbox), 0)
        bot.refresh_from_db()
        self.assertEqual(bot.last_status, STATUS_FAILED)

    def test_test_send_targets_one_address_and_ignores_the_empty_skip(self):
        bot = make_bot(email_to='planning@example.com', send_when_empty=False)
        with self._patch_report(rows=[]):
            execution = services.send_test_email(bot, 'me@example.com', actor=self.user)
        self.assertEqual(execution.status, STATUS_SENT)
        self.assertEqual(execution.recipients_to, 'me@example.com')
        self.assertEqual(mail.outbox[0].to, ['me@example.com'])
        self.assertEqual(execution.triggered_by, self.user)

    def test_manual_run_does_not_move_next_run_at(self):
        bot = make_bot()
        services.refresh_next_run(bot)
        scheduled = bot.next_run_at
        with self._patch_report():
            services.run_bot_manually(bot, actor=self.user)
        bot.refresh_from_db()
        self.assertEqual(bot.next_run_at, scheduled)

    def test_refresh_next_run_clears_the_schedule_for_an_inactive_bot(self):
        bot = make_bot(is_active=False)
        self.assertIsNone(services.refresh_next_run(bot))


class SeedCommandTests(TestCase):
    def test_seed_is_idempotent_and_leaves_edits_alone(self):
        from django.core.management import call_command
        from io import StringIO

        call_command('seed_bots', stdout=StringIO())
        bot = BotAutomation.objects.get(code='PENDING_PRODUCTION_RELEASE')
        self.assertFalse(bot.is_active)
        self.assertEqual(bot.report_slug, 'pending-work')
        self.assertEqual(bot.report_filters, {'stage': 'not_released'})

        bot.email_to = 'planning@example.com'
        bot.save()

        call_command('seed_bots', stdout=StringIO())
        self.assertEqual(BotAutomation.objects.filter(code='PENDING_PRODUCTION_RELEASE').count(), 1)
        bot.refresh_from_db()
        self.assertEqual(bot.email_to, 'planning@example.com')


class EditFormRoundTripTests(TestCase):
    """Opening an existing bot's form and saving it unchanged must be a no-op.

    ModelForm fills self.initial from the instance, and BoundField.value() reads
    self.initial before field.initial — so any field whose stored shape differs
    from its widget's shape has to be corrected there, or the rendered form
    carries values its own clean_* methods reject (or silently drops them).
    """

    def _bound_data(self, form):
        """The POST an untouched browser form would produce."""
        data = {}
        for name, field in form.fields.items():
            value = form[name].value()
            if value is None or value is False:
                continue
            if isinstance(value, list):
                data[name] = [str(item) for item in value]
            elif value is True:
                data[name] = 'on'
            else:
                data[name] = str(value)
        return data

    def test_filters_render_as_json_not_python_repr(self):
        bot = make_bot(report_filters={'stage': 'not_released'})
        rendered = BotAutomationForm(instance=bot)['report_filters'].value()
        self.assertEqual(json.loads(rendered), {'stage': 'not_released'})
        self.assertNotIn("'", rendered, 'JSON textarea must not show a Python dict repr')

    def test_weekdays_and_roles_render_as_checkbox_lists(self):
        bot = make_bot(weekdays='0,2,4', recipient_roles='planner,manager')
        form = BotAutomationForm(instance=bot)
        self.assertEqual(form['weekdays'].value(), ['0', '2', '4'])
        self.assertEqual(form['recipient_roles'].value(), ['planner', 'manager'])

    def test_saving_an_untouched_form_changes_nothing(self):
        bot = make_bot(
            report_filters={'stage': 'not_released'},
            weekdays='0,2,4',
            recipient_roles='planner,manager',
        )
        form = BotAutomationForm(self._bound_data(BotAutomationForm(instance=bot)), instance=bot)
        self.assertTrue(form.is_valid(), form.errors.as_json())
        form.save()

        bot.refresh_from_db()
        self.assertEqual(bot.report_filters, {'stage': 'not_released'})
        self.assertEqual(bot.weekdays, '0,2,4')
        self.assertEqual(bot.recipient_roles, 'planner,manager')
