from __future__ import annotations

from hashlib import sha256

from django.core.cache import cache
from django.core.exceptions import PermissionDenied
from django.http import Http404
from django.utils import timezone
import zoneinfo

from core.navigation import REPORTS_NAV_ROLES
from core.permissions import user_has_permission
from reports.filters import parse_universal_filters
from reports.report_engine.serializer import to_json_safe
from reports.report_registry import registry


def _has_access(request, permissions: tuple[str, ...]) -> bool:
    if not permissions:
        return True
    user = getattr(request, 'user', None)
    if not user or not user.is_authenticated:
        return False
    if user.is_superuser:
        return True
    profile = getattr(user, 'profile', None)
    if profile is not None and getattr(profile, 'role', None) in REPORTS_NAV_ROLES:
        return True
    # `permissions` entries are Django-style codes (e.g. 'core.view_reports'), but
    # this app's actual role/permission grants (Settings -> Roles & Permissions,
    # incl. per-user overrides) live in the soft-coded core.permissions system
    # under 'action.<name>' codes, not Django's built-in auth permissions. Check
    # both so grants made through Settings actually take effect.
    if any(user_has_permission(user, f'action.{code.rsplit(".", 1)[-1]}') for code in permissions):
        return True
    return any(user.has_perm(code) for code in permissions)


CACHE_VERSION_KEY = 'reports:engine:cache_version'


def get_cache_version() -> int:
    version = cache.get(CACHE_VERSION_KEY)
    if version is None:
        version = 1
        cache.set(CACHE_VERSION_KEY, version, timeout=None)
    return version


def bump_cache_version() -> None:
    try:
        cache.incr(CACHE_VERSION_KEY)
    except ValueError:
        cache.set(CACHE_VERSION_KEY, 1, timeout=None)


def _cache_key(slug: str, user_id: int | None, filters: dict) -> str:
    key_input = f'{get_cache_version()}|{slug}|{user_id}|{filters}'
    return f'reports:engine:{sha256(key_input.encode("utf-8")).hexdigest()}'


def run_report(slug: str, request, filters: dict | None = None) -> dict:
    definition = registry.get(slug)
    if definition is None:
        raise Http404('Report not found')

    if not _has_access(request, definition.permissions):
        raise PermissionDenied('Access denied to this report')

    if filters is None:
        filters = parse_universal_filters(request)

    key = _cache_key(slug, getattr(request.user, 'id', None) if request.user.is_authenticated else None, filters)
    cached = cache.get(key)
    if cached is not None:
        return cached

    raw_data = definition.executor(request, filters)
    payload = {
        'report': {
            'slug': definition.slug,
            'title': definition.title,
            'description': definition.description,
            'department': definition.department,
            'filters': definition.filters,
            'supported_exports': definition.supported_exports,
            'supported_charts': definition.supported_charts,
            'drilldown_support': definition.drilldown_support,
            'category': definition.category,
            'navigation_group': definition.navigation_group,
            'icon': definition.icon,
        },
        'filters': filters,
        'generated_at': timezone.now().astimezone(zoneinfo.ZoneInfo('Asia/Karachi')).strftime('%Y-%m-%d %I:%M:%S %p PKT'),
        'data': to_json_safe(raw_data),
    }
    cache.set(key, payload, timeout=max(30, int(definition.cache_timeout or 300)))
    return payload
