"""Seeds the built-in automations and the nav.bot permission.

Idempotent: re-running it never clobbers configuration an admin has edited from
the UI — an automation row is only created, never overwritten.

Every bot here ships inactive and with no recipients. Recipients are business
config, so they are filled in from the ERP UI, not baked into the repo.
"""
import datetime

from django.core.management.base import BaseCommand
from django.db import transaction

from bot.models import (
    DEFAULT_BODY_TEMPLATE,
    DEFAULT_SUBJECT_TEMPLATE,
    BotAutomation,
    stage_body_template,
)
from core import navigation as nav

BOT_001_CODE = 'PENDING_PRODUCTION_RELEASE'

# The three production backlog bots. Same report, different `stage` slice —
# which is exactly the "new automation = new configuration row, not new code"
# property the app was built for.
STAGE_BOTS = [
    {
        'code': 'PRODUCTION_PENDING_PRINTING',
        'stage': 'printing',
        'label': 'Printing',
        'name': 'Production Pending - Printing',
        'description': (
            'Daily email listing job cards with printing still outstanding '
            '(order quantity not yet printed).'
        ),
        'send_time': datetime.time(8, 0),
    },
    {
        'code': 'PRODUCTION_PENDING_PACKING',
        'stage': 'packing',
        'label': 'Packing',
        'name': 'Production Pending - Packing',
        'description': (
            'Daily email listing job cards printed but not yet packed '
            '(including Cut & Pack jobs measured against order quantity).'
        ),
        'send_time': datetime.time(8, 10),
    },
    {
        'code': 'PRODUCTION_PENDING_DISPATCH',
        'stage': 'dispatch',
        'label': 'Dispatch',
        'name': 'Production Pending - Dispatch',
        'description': (
            'Daily email listing job cards packed but not yet dispatched.'
        ),
        'send_time': datetime.time(8, 20),
    },
]


class Command(BaseCommand):
    help = (
        'Creates the built-in automations (all inactive, no recipients) and '
        'grants the nav.bot permission to the default roles.'
    )

    @transaction.atomic
    def handle(self, *args, **options):
        self._seed_permission()
        self._seed_bot_001()
        for spec in STAGE_BOTS:
            self._seed_stage_bot(spec)
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
            report_period='month',
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

    def _seed_stage_bot(self, spec):
        if BotAutomation.objects.filter(code=spec['code']).exists():
            self.stdout.write(f"{spec['code']} already exists — left untouched.")
            return

        bot = BotAutomation.objects.create(
            code=spec['code'],
            name=spec['name'],
            description=spec['description'],
            is_active=False,
            report_slug='pending-work',
            report_filters={'stage': spec['stage']},
            # Same default the report screen opens with. Switch to "All Time"
            # from the UI to cover the full backlog instead of this month.
            report_period='month',
            frequency='DAILY',
            send_time=spec['send_time'],
            subject_template=f"Production Pending - {spec['label']} - {{{{date}}}}",
            body_template=stage_body_template(spec['label'].lower()),
            attach_report=True,
            attachment_format='xlsx',
            send_when_empty=False,
        )
        self.stdout.write(self.style.SUCCESS(f'Created bot: {bot}'))
        self.stdout.write('  Inactive, no recipients — add recipients in /bot/ then activate.')
