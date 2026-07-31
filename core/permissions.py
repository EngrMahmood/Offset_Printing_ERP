"""Central permission-check API for the soft-coded access control system.

Roles and Permissions are DB rows (core.models.Role / Permission) editable from
Settings -> Roles & Permissions. A user's effective permission set is their
role's permissions, plus/minus any per-user UserPermissionOverride rows.
"""
from __future__ import annotations

from django.core.cache import cache

_VERSION_KEY = 'access_control:cache_version'
_CODES_KEY = 'access_control:codes:{version}:{user_id}'


def _cache_version() -> int:
    version = cache.get(_VERSION_KEY)
    if version is None:
        version = 1
        cache.set(_VERSION_KEY, version, None)
    return version


def bump_cache_version() -> None:
    """Invalidate all cached permission sets (called on Role/Permission/override changes)."""
    try:
        cache.incr(_VERSION_KEY)
    except ValueError:
        cache.set(_VERSION_KEY, 1, None)


def get_granted_permission_codes(user) -> set[str]:
    if not getattr(user, 'is_authenticated', False):
        return set()

    if getattr(user, 'is_superuser', False):
        from core.models import Permission
        return set(Permission.objects.filter(is_active=True).values_list('code', flat=True))

    cache_key = _CODES_KEY.format(version=_cache_version(), user_id=user.id)
    cached = cache.get(cache_key)
    if cached is not None:
        return cached

    from core.models import Role, UserPermissionOverride

    profile = getattr(user, 'profile', None)
    role_slug = (getattr(profile, 'role', '') or '').strip().lower()

    codes: set[str] = set()
    role = Role.objects.filter(slug=role_slug).prefetch_related('permissions').first()
    if role is not None:
        codes = set(role.permissions.filter(is_active=True).values_list('code', flat=True))

    overrides = UserPermissionOverride.objects.filter(user=user).select_related('permission')
    for override in overrides:
        if not override.permission.is_active:
            continue
        if override.granted:
            codes.add(override.permission.code)
        else:
            codes.discard(override.permission.code)

    cache.set(cache_key, codes, 300)
    return codes


def user_has_permission(user, code: str) -> bool:
    return code in get_granted_permission_codes(user)
