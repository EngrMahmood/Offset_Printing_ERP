from django.core.management.base import BaseCommand
from django.db import transaction

from core import navigation as nav
from core.models import Permission, Role

ALL_ROLE_SLUGS = set(getattr(nav, 'BOM_NAV_ROLES', {'admin', 'manager', 'planner', 'production_manager', 'supply_chain'}))
MANAGE_ROLE_SLUGS = {'admin', 'manager', 'production_manager', 'planner'}
APPROVE_ROLE_SLUGS = {'admin', 'manager'}

BOM_PERMISSIONS = {
    'nav.bom': ('BOM', 'Access the Bill of Materials module', 'Navigation & Modules', ALL_ROLE_SLUGS),
    'action.manage_bom': ('Manage BOM Masters', 'Create/edit raw items, templates and spec attributes', 'BOM', MANAGE_ROLE_SLUGS),
    'action.approve_bom': ('Approve BOM', 'Review and approve generated/imported BOMs', 'BOM', APPROVE_ROLE_SLUGS),
}


class Command(BaseCommand):
    help = 'Seeds BOM module permissions and grants default role access.'

    @transaction.atomic
    def handle(self, *args, **options):
        role_by_slug = {role.slug: role for role in Role.objects.all()}

        for code, (name, description, category, allowed_roles) in BOM_PERMISSIONS.items():
            permission, created = Permission.objects.get_or_create(
                code=code, defaults={'name': name, 'description': description, 'category': category},
            )
            self.stdout.write(f"{'Created' if created else 'Found'} permission: {code}")
            for slug in allowed_roles:
                role = role_by_slug.get(slug)
                if role is None:
                    continue
                role.permissions.add(permission)

        from core.permissions import bump_cache_version
        bump_cache_version()
        self.stdout.write(self.style.SUCCESS('BOM permission seed complete.'))
