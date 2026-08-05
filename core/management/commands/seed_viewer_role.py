from django.contrib.auth.models import Group, Permission as DjangoPermission
from django.core.management.base import BaseCommand
from django.db import transaction

from core.models import Role, Permission

VIEWER_ROLE_SLUG = 'viewer'
VIEWER_ROLE_NAME = 'Viewer (Read-Only)'
VIEWER_ROLE_DESCRIPTION = (
    "Read-only access to every module — for an outside reviewer or the CEO to browse "
    "the ERP without being able to create, edit, approve, or delete anything. Full "
    "normal chat access (send messages, create groups, start calls)."
)

# Every nav.* permission — full page visibility across all modules.
NAV_CODES = [
    'nav.planning', 'nav.qc', 'nav.production', 'nav.dispatch', 'nav.master_data',
    'nav.migration', 'nav.reports', 'nav.job_summary', 'nav.printing_plates', 'nav.guides',
    'nav.supply_chain', 'nav.audit', 'nav.item_request', 'nav.maintenance', 'nav.chat',
]

# Read-only action.* permissions. action.manage_masters is included because Master Data
# has no dedicated read-only permission — its write branches separately hardcode
# role == 'admin' / is_superuser in core/views.py, so granting this only unlocks
# viewing the page, not editing it.
VIEW_ACTION_CODES = [
    'action.view_plate_queue', 'action.view_jobcard', 'action.view_planning_queue',
    'action.view_approval_queue', 'action.view_production_records', 'action.view_qc_queue',
    'action.view_pm_queue', 'action.view_dispatch_records', 'action.view_analytics',
    'action.view_reports', 'action.view_job_summary', 'action.view_sku_master_review_queue',
    'action.manage_masters',
]

# Normal (non-moderation) chat access — same as any regular team member.
CHAT_ACTION_CODES = [
    'action.chat_create_group', 'action.chat_initiate_call', 'action.chat_manage_group_members',
]

# The Migration module uses Django's built-in auth permissions (not this app's
# Role/Permission model) — see migration/models.py Meta.permissions. Viewer users
# are added to this Group (see core.views.access_user_role_update/user_create) so
# they can browse migration.view_import without ever getting migration.run_import.
VIEWER_DJANGO_GROUP = 'Viewer'
VIEWER_DJANGO_PERMS = ['migration.view_import']


class Command(BaseCommand):
    help = "Creates/updates the 'viewer' role: read-only access to every page, full chat access."

    @transaction.atomic
    def handle(self, *args, **options):
        role, created = Role.objects.get_or_create(
            slug=VIEWER_ROLE_SLUG,
            defaults={'display_name': VIEWER_ROLE_NAME, 'description': VIEWER_ROLE_DESCRIPTION},
        )
        if not created:
            role.display_name = VIEWER_ROLE_NAME
            role.description = VIEWER_ROLE_DESCRIPTION
            role.save()
        self.stdout.write(f"{'Created' if created else 'Updated'} role: {VIEWER_ROLE_SLUG}")

        codes = NAV_CODES + VIEW_ACTION_CODES + CHAT_ACTION_CODES
        permissions = Permission.objects.filter(code__in=codes)
        missing = set(codes) - set(permissions.values_list('code', flat=True))
        if missing:
            self.stdout.write(self.style.WARNING(f"Permission codes not found in DB (skipped): {sorted(missing)}"))

        role.permissions.set(permissions)
        self.stdout.write(self.style.SUCCESS(f"Granted {permissions.count()} permissions to '{VIEWER_ROLE_SLUG}'."))

        group, _ = Group.objects.get_or_create(name=VIEWER_DJANGO_GROUP)
        django_perms = []
        for perm_string in VIEWER_DJANGO_PERMS:
            app_label, codename = perm_string.split('.')
            try:
                django_perms.append(DjangoPermission.objects.get(content_type__app_label=app_label, codename=codename))
            except DjangoPermission.DoesNotExist:
                self.stdout.write(self.style.WARNING(f"Django permission not found (skipped): {perm_string}"))
        group.permissions.set(django_perms)
        self.stdout.write(self.style.SUCCESS(f"Django group '{VIEWER_DJANGO_GROUP}' has {len(django_perms)} permission(s)."))
