"""Seeds Bot 001 and the nav.bot permission.

Idempotent: re-running it never clobbers configuration an admin has edited from
the UI — the automation row is only created, never overwritten.
"""
import datetime

from django.core.management.base import BaseCommand
from django.db import transaction

from bot.models import DEFAULT_BODY_TEMPLATE, DEFAULT_SUBJECT_TEMPLATE, BotAutomation
from core import navigation as nav

BOT_001_CODE = 'PENDING_PRODUCTION_RELEASE'


class Command(BaseCommand):
    help = (
        'Creates the "Pending Work Release to Production" automation (inactive) '
        'and grants the nav.bot permission to the default roles.'
    )

    @transaction.atomic
    def handle(self, *args, **options):
        self._seed_permission()
        self._seed_bot_001()
        self.stdout.write(self.style.SUCCESS('Bot seed complete.'))

    def _seed_permission(self):
        from core.models import Permission, Role

        permission, created = Permission.objects.get_or_create(
            code=nav.NAV_PERMISSION_CODES['bot'],
            defaults={
                'name': 'Bot / Automation',
                'description': 'Configure and run scheduled report automations',
                'category': 'Navigation & Modules',
            },
        )
        self.stdout.write(f"{'Created' if created else 'Found'} permission: {permission.code}")

        for slug in sorted(nav.BOT_NAV_ROLES):
            role = Role.objects.filter(slug=slug).first()
            if role is None:
                self.stdout.write(self.style.WARNING(f'  role "{slug}" not found — skipped'))
                continue
            role.permissions.add(permission)
            self.stdout.write(f'  granted to role: {slug}')

    def _seed_bot_001(self):
        if BotAutomation.objects.filter(code=BOT_001_CODE).exists():
            self.stdout.write('Bot 001 already exists — left untouched.')
            return

        bot = BotAutomation.objects.create(
            code=BOT_001_CODE,
            name='Pending Work Release to Production',
            description=(
                'Daily email to Planning listing job cards that are not yet released '
                'to production, so the backlog is visible before the shift starts.'
            ),
            # Ships OFF: preview and test-send it before it mails the Planning team.
            is_active=False,
            report_slug='pending-work',
            report_filters={'stage': 'not_released'},
            frequency='DAILY',
            send_time=datetime.time(8, 30),
            subject_template=DEFAULT_SUBJECT_TEMPLATE,
            body_template=DEFAULT_BODY_TEMPLATE,
            attach_report=True,
            attachment_format='xlsx',
            send_when_empty=False,
        )
        self.stdout.write(self.style.SUCCESS(f'Created bot: {bot}'))
        self.stdout.write(
            '  Inactive by design — open /bot/, preview it, send a test, then activate.'
        )
