from django.core.management.base import BaseCommand
from django.db import transaction

from core import navigation as nav
from core.models import Permission, Role, UserProfile

# All roles get base chat access + call-initiation (per product decision).
ALL_ROLE_SLUGS = set(nav.CHAT_NAV_ROLES)

# Group creation / member management restricted to office & supervisory roles —
# floor-only "operator" role can chat/DM but not spin up new group rooms.
GROUP_ADMIN_ROLE_SLUGS = ALL_ROLE_SLUGS - {'operator'}

# Moderation override (delete any message) — admins/managers only.
MODERATION_ROLE_SLUGS = {'admin', 'manager'}

CHAT_PERMISSIONS = {
    'nav.chat': (
        'Chat', 'Access the Chat module', 'Navigation & Modules', ALL_ROLE_SLUGS,
    ),
    'action.chat_create_group': (
        'Create Group Chat', 'Can create a new group chat room', 'Chat', GROUP_ADMIN_ROLE_SLUGS,
    ),
    'action.chat_manage_group_members': (
        'Manage Group Members', 'Can add/remove members and promote group admins', 'Chat', GROUP_ADMIN_ROLE_SLUGS,
    ),
    'action.chat_initiate_call': (
        'Start Voice/Video Call', 'Can start a voice or video call', 'Chat', ALL_ROLE_SLUGS,
    ),
    'action.chat_delete_any_message': (
        'Delete Any Chat Message', 'Moderation override to delete other users\' messages', 'Chat', MODERATION_ROLE_SLUGS,
    ),
}


class Command(BaseCommand):
    help = "Seeds Chat module permissions and grants default role access."

    @transaction.atomic
    def handle(self, *args, **options):
        role_by_slug = {role.slug: role for role in Role.objects.filter(slug__in=ALL_ROLE_SLUGS)}

        for code, (name, description, category, allowed_roles) in CHAT_PERMISSIONS.items():
            permission, created = Permission.objects.get_or_create(
                code=code,
                defaults={'name': name, 'description': description, 'category': category},
            )
            self.stdout.write(f"{'Created' if created else 'Found'} permission: {code}")

            for slug in allowed_roles:
                role = role_by_slug.get(slug)
                if role is None:
                    continue
                role.permissions.add(permission)

        self.stdout.write(self.style.SUCCESS('Chat permission seed complete.'))
