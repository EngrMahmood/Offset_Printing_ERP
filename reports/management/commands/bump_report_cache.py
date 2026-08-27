"""Invalidates every cached report payload.

Report payloads are cached by (cache_version, slug, user, filters) — see
reports/report_engine/engine.py. That version only bumps on data changes
(PlanningJob/Machine saves etc.), not on code changes, so a deploy that
changes what a report builder returns can serve stale cached payloads
(missing/wrong fields) for up to each report's cache_timeout. Run this
after every deploy to force a clean slate immediately instead of waiting
out the TTL.
"""
from django.core.management.base import BaseCommand

from reports.report_engine.engine import bump_cache_version


class Command(BaseCommand):
    help = "Invalidate all cached report payloads (run after every deploy)."

    def handle(self, *args, **options):
        bump_cache_version()
        self.stdout.write(self.style.SUCCESS('Report cache invalidated.'))
