from django.core.management.base import BaseCommand
from django.db import transaction

from core import navigation as nav
from core.models import Permission, Role

ALL_ROLE_SLUGS = set(getattr(nav, 'FLEXO_NAV_ROLES', {'admin', 'manager', 'planner', 'production_manager', 'production'}))

FLEXO_PERMISSIONS = {
    'nav.flexo': ('Flexo', 'Access the Flexo Planning / Job Card module', 'Navigation & Modules', ALL_ROLE_SLUGS),
}


class Command(BaseCommand):
    help = 'Seeds Flexo module permissions and grants default role access.'

    @transaction.atomic
    def handle(self, *args, **options):
        role_by_slug = {role.slug: role for role in Role.objects.all()}

        for code, (name, description, category, allowed_roles) in FLEXO_PERMISSIONS.items():
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
        self.stdout.write(self.style.SUCCESS('Flexo permission seed complete.'))
