from __future__ import annotations

from hashlib import sha256

from django.core.cache import cache
from django.core.exceptions import PermissionDenied
from django.utils import timezone

from reports.filters import parse_universal_filters
from reports.report_engine.serializer import to_json_safe
from reports.report_registry import registry


def _has_access(request, permissions: tuple[str, ...]) -> bool:
    if not permissions:
        return True
    user = getattr(request, 'user', None)
    if not user or not user.is_authenticated:
        return False
    return any(user.has_perm(code) for code in permissions)


def _cache_key(slug: str, user_id: int | None, filters: dict) -> str:
    key_input = f'{slug}|{user_id}|{filters}'
    return f'reports:engine:{sha256(key_input.encode("utf-8")).hexdigest()}'


def run_report(slug: str, request) -> dict:
    definition = registry.get(slug)
    if definition is None:
        raise KeyError(f'Report not found: {slug}')

    if not _has_access(request, definition.permissions):
        raise PermissionDenied('You do not have permission to view this report.')

    filters = parse_universal_filters(request)
    key = _cache_key(definition.slug, getattr(getattr(request, 'user', None), 'id', None), filters)
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
        'generated_at': timezone.now().isoformat(),
        'data': to_json_safe(raw_data),
    }
    cache.set(key, payload, timeout=max(30, int(definition.cache_timeout or 300)))
    return payload
