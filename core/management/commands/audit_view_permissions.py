"""On-demand audit: which URL-routed views have no soft-coded permission
gating, so a newly added screen doesn't silently go missing from Settings ->
Roles & Access Control the way Production WIP and Job Card Finalization did.

Run it any time after adding new screens/URLs:

    python manage.py audit_view_permissions

Two tiers, reported separately:
  - UNGUARDED: no recognized access-control decorator at all (not even a
    hardcoded role check) — genuinely open to any logged-in user. Always
    worth a look.
  - HARDCODED (not UI-configurable): gated, but via @require_role(...) or an
    equivalent hand-rolled role/permission check rather than
    @permission_required('can_xxx') — real access control, just not
    something an admin can adjust per role from the Roles & Access Control
    screen. Often fine (e.g. the endpoints that manage the access-control
    system itself deliberately stay hardcoded-admin-only), but worth a
    second look for any screen that should differ by role.

This can only see which decorator was used, not judge whether it's correct.
Class-based views and third-party/admin URLs are skipped rather than guessed
at. See core/views.py (permission_required, require_role) for how the
_rbac_configurable / _rbac_hardcoded markers get set.
"""
from __future__ import annotations

from django.core.management.base import BaseCommand
from django.urls import URLPattern, URLResolver, get_resolver

# URL names that are legitimately open to any authenticated user, or are the
# login/logout/password-reset plumbing itself — not real gaps.
EXEMPT_VIEW_NAMES = {
    'login', 'logout', 'password_reset', 'password_reset_done',
    'password_reset_confirm', 'password_reset_complete', 'home',
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


class Command(BaseCommand):
    help = "Lists URL-routed views with no soft-coded (Roles & Access Control) permission gating."

    def handle(self, *args, **options):
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

        self.stdout.write(self.style.ERROR(f"\nUNGUARDED ({len(unguarded)}) — no access-control decorator found:"))
        for label in unguarded:
            self.stdout.write(f"  {label}")

        self.stdout.write(self.style.WARNING(f"\nHARDCODED, not UI-configurable ({len(hardcoded)}):"))
        for label in hardcoded:
            self.stdout.write(f"  {label}")

        self.stdout.write(self.style.SUCCESS(
            f"\n{len(seen) - len(unguarded) - len(hardcoded)} view(s) already configurable from "
            f"Settings -> Roles & Access Control."
        ))
