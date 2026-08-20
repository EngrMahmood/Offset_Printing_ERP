"""On-demand audit: which URL-routed views have no soft-coded permission
gating, so a newly added screen doesn't silently go missing from Settings ->
Roles & Access Control the way Production WIP and Job Card Finalization did.

Run it any time after adding new screens/URLs:

    python manage.py audit_view_permissions

Same report is also available in-app at Settings -> Roles & Access Control
-> Permission Audit (superuser only) — see core/permission_audit.py, which
both share.

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
at.
"""
from __future__ import annotations

from django.core.management.base import BaseCommand

from core.permission_audit import run_permission_audit


class Command(BaseCommand):
    help = "Lists URL-routed views with no soft-coded (Roles & Access Control) permission gating."

    def handle(self, *args, **options):
        result = run_permission_audit()

        self.stdout.write(self.style.ERROR(f"\nUNGUARDED ({len(result['unguarded'])}) — no access-control decorator found:"))
        for label in result['unguarded']:
            self.stdout.write(f"  {label}")

        self.stdout.write(self.style.WARNING(f"\nHARDCODED, not UI-configurable ({len(result['hardcoded'])}):"))
        for label in result['hardcoded']:
            self.stdout.write(f"  {label}")

        self.stdout.write(self.style.SUCCESS(
            f"\n{result['configurable_count']} view(s) already configurable from "
            f"Settings -> Roles & Access Control."
        ))
