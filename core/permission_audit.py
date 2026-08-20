"""Shared logic for auditing which URL-routed views have no soft-coded
permission gating — used by both `manage.py audit_view_permissions` (CLI)
and the Settings -> Roles & Access Control -> Permission Audit screen (UI).

See core/views.py (permission_required, require_role) for how the
_rbac_configurable / _rbac_hardcoded markers referenced below get set.
"""
from __future__ import annotations

from django.urls import URLPattern, URLResolver, get_resolver

# URL names that are legitimately open to any authenticated user, or are the
# login/logout/password-reset plumbing itself — not real gaps.
EXEMPT_VIEW_NAMES = {
    'login', 'logout', 'password_reset', 'password_reset_done',
    'password_reset_confirm', 'password_reset_complete', 'home',
    'permission_audit_view',
}

# Namespaces that are intentionally open to every authenticated user.
EXEMPT_NAMESPACES = {'tasks', 'audit', 'chat', 'manual_working', 'floor_dashboard'}

OWN_APP_PREFIXES = (
    'core.', 'production.', 'planning.', 'qc.', 'dispatch.', 'supply_chain.',
    'printing_plates.', 'job_summary.', 'maintenance.', 'reports.', 'migration.',
)


def _iter_view_funcs(patterns, namespace=None):
    for pattern in patterns:
        if isinstance(pattern, URLResolver):
            yield from _iter_view_funcs(pattern.url_patterns, namespace=pattern.namespace or namespace)
        elif isinstance(pattern, URLPattern):
            yield namespace, pattern.name, pattern.callback


def run_permission_audit():
    """Returns {'unguarded': [...], 'hardcoded': [...], 'configurable_count': int, 'total': int},
    each list of labels sorted alphabetically."""
    unguarded = []
    hardcoded = []
    seen = set()

    for namespace, name, view_func in _iter_view_funcs(get_resolver().url_patterns):
        if not name or name in EXEMPT_VIEW_NAMES or namespace in EXEMPT_NAMESPACES:
            continue
        module = getattr(view_func, '__module__', '') or ''
        if not module.startswith(OWN_APP_PREFIXES) or module.endswith('.admin'):
            continue
        if getattr(view_func, 'view_class', None):
            continue  # class-based view — not decorator-wrapped the same way, skip rather than guess

        key = (namespace, name)
        if key in seen:
            continue
        seen.add(key)

        label = f"{namespace}:{name}" if namespace else name
        if getattr(view_func, '_rbac_configurable', False):
            continue
        if getattr(view_func, '_rbac_hardcoded', False):
            hardcoded.append(label)
        else:
            unguarded.append(label)

    unguarded.sort()
    hardcoded.sort()

    return {
        'unguarded': unguarded,
        'hardcoded': hardcoded,
        'configurable_count': len(seen) - len(unguarded) - len(hardcoded),
        'total': len(seen),
    }
