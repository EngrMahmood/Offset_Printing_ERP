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
DAILY_PRODUCTION_CODE = 'DAILY_PRODUCTION_REPORT'
STOCK_REPORT_CODE = 'STOCK_REPORT_EXCESS_INVENTORY'

STOCK_REPORT_BODY = """<p>Dear Planning Team,</p>

<p>Below is the current finished-goods excess stock on hand as of {{date}} —
leftover from over-packed runs, or entered manually from a physical stock
check. Carried forward automatically the next time each SKU is planned.</p>

{{report_table}}

<p>Total SKUs/jobs holding stock: {{total_records}}</p>

<p>Regards,<br>
Production Printing</p>
"""

DAILY_PRODUCTION_BODY = """<p>Dear Team,</p>

<p>Below is the daily production summary for {{period_label}} ({{period_from}}) —
jobs released, printing, packing and dispatch for the day.</p>

{{report_table}}

<p>Kindly review and address any shortfall against plan.</p>

<p>Regards,<br>
Production Printing</p>
"""

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
        self._seed_daily_production()
        self._seed_stock_report()
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

    def _seed_daily_production(self):
        if BotAutomation.objects.filter(code=DAILY_PRODUCTION_CODE).exists():
            self.stdout.write(f'{DAILY_PRODUCTION_CODE} already exists — left untouched.')
            return

        bot = BotAutomation.objects.create(
            code=DAILY_PRODUCTION_CODE,
            name='Daily Production Report',
            description=(
                'Every morning, emails yesterday\'s Daily Production overview — jobs '
                'released, impressions, printed sheets, packed pcs, dispatch and '
                'process wastage for the day.'
            ),
            is_active=False,
            report_slug='daily-production',
            # `tab` picks which of the report's six tables gets exported; the other
            # tabs (printing / packing / dispatch / released / wastage) are reachable
            # by cloning this bot and changing this one value.
            report_filters={'tab': 'overview'},
            report_period='yesterday',
            frequency='DAILY',
            # Ahead of the 08:00-08:30 backlog bots — yesterday's numbers first,
            # then today's outstanding work.
            send_time=datetime.time(7, 0),
            subject_template='Daily Production Report - {{period_label}} ({{period_from}})',
            body_template=DAILY_PRODUCTION_BODY,
            attach_report=True,
            attachment_format='xlsx',
            send_when_empty=False,
        )
        self.stdout.write(self.style.SUCCESS(f'Created bot: {bot}'))
        self.stdout.write('  Inactive, no recipients — add recipients in /bot/ then activate.')

    def _seed_stock_report(self):
        if BotAutomation.objects.filter(code=STOCK_REPORT_CODE).exists():
            self.stdout.write(f'{STOCK_REPORT_CODE} already exists — left untouched.')
            return

        bot = BotAutomation.objects.create(
            code=STOCK_REPORT_CODE,
            name='Stock Report - Excess Inventory',
            description=(
                'Daily email listing finished-goods excess stock currently on hand '
                'per SKU/job — leftover from over-packed runs, or entered manually '
                'from a physical stock check.'
            ),
            is_active=False,
            report_slug='stock-report',
            # Point-in-time snapshot query — the report ignores date filters.
            report_filters={},
            report_period='',
            frequency='DAILY',
            send_time=datetime.time(7, 30),
            subject_template='Stock Report - Excess Inventory - {{date}}',
            body_template=STOCK_REPORT_BODY,
            attach_report=True,
            attachment_format='xlsx',
            send_when_empty=False,
        )
        self.stdout.write(self.style.SUCCESS(f'Created bot: {bot}'))
        self.stdout.write('  Inactive, no recipients — add recipients in /bot/ then activate.')

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
