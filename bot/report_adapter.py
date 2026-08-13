"""Bridge between a BotAutomation row and the reports app.

Reuses reports.report_engine.run_report so a bot can email any registered
report without duplicating its query.

IMPORTANT — why the stub request carries a populated GET:
    run_report(slug, request, filters) passes `filters` to the executor, but
    several executors ignore that argument and read request.GET directly
    (e.g. _pending_work_executor -> build_pending_work_context(request), which
    calls _parse_period_filter(request) and request.GET.get('stage')).
    So the bot's report_filters must be written into request.GET *and* passed
    as `filters` — the former drives the query, the latter drives the cache key.
"""
from __future__ import annotations

from urllib.parse import urlencode

from django.http import QueryDict

from reports.export.services import (
    build_export_filename,
    build_report_title,
    export_as_csv,
    export_as_pdf,
    export_as_xlsx,
)
from reports.report_engine import run_report
from reports.report_registry import registry


class StubRequest:
    """The minimal request surface the reports engine and its executors touch:
    `.GET` (a QueryDict), `.user`, and `.build_absolute_uri` for the odd
    executor that links back to a page."""

    def __init__(self, user, filters: dict | None = None):
        pairs = [
            (key, str(value))
            for key, value in (filters or {}).items()
            if value is not None
        ]
        self.GET = QueryDict(urlencode(pairs))
        self.POST = QueryDict('')
        self.user = user
        self.method = 'GET'

    def build_absolute_uri(self, location=''):
        return location or '/'


def available_report_choices():
    """(slug, label) pairs for every registered report — populates the bot form
    so the report list is never hard-coded."""
    return [(definition.slug, definition.title) for definition in registry.all()]


def build_stub_request(user, filters: dict | None = None) -> StubRequest:
    return StubRequest(user, filters)


def fetch_report(bot, user=None) -> dict:
    """Run the bot's report and return the engine payload.

    Raises whatever run_report raises (Http404 for an unknown slug,
    PermissionDenied when run_as lacks report access) — the caller logs it
    onto the BotExecution row.
    """
    run_as = user or bot.resolve_run_as_user()
    # effective_filters(), not report_filters — it layers the bot's period
    # control on top of the raw JSON.
    filters = bot.effective_filters()
    request = build_stub_request(run_as, filters)
    return run_report(bot.report_slug, request, filters=filters)


def extract_rows(payload: dict) -> tuple[list[str], dict[str, str], list[dict]]:
    """Pull (headers, header_labels, rows) out of an engine payload.

    Mirrors how reports/export/services.py reads the same payload, so the
    email body and the attachment always agree.
    """
    data = payload.get('data') or {}
    if not isinstance(data, dict):
        data = {}

    rows = data.get('export_rows')
    if not isinstance(rows, list):
        rows = []
    rows = [row for row in rows if isinstance(row, dict)]

    headers = payload.get('headers') or data.get('headers')
    if not headers:
        headers = sorted({key for row in rows for key in row.keys()})
    headers = [str(header) for header in headers]

    labels = payload.get('header_labels') or data.get('header_labels') or {}
    if not isinstance(labels, dict):
        labels = {}

    return headers, labels, rows


_EXPORTERS = {
    'xlsx': export_as_xlsx,
    'csv': export_as_csv,
    'pdf': export_as_pdf,
}

_MIME_TYPES = {
    'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'csv': 'text/csv',
    'pdf': 'application/pdf',
}


def build_attachment(payload: dict, slug: str, fmt: str) -> tuple[str, bytes, str]:
    """(filename, content, mimetype) using the reports app's own exporters."""
    fmt = (fmt or 'xlsx').lower()
    exporter = _EXPORTERS.get(fmt)
    if exporter is None:
        raise ValueError(f'Unsupported attachment format: {fmt}')
    content = exporter(payload)
    filename = f'{build_export_filename(payload, slug)}.{fmt}'
    return filename, content, _MIME_TYPES[fmt]


def period_info(payload: dict) -> tuple[str, str, str]:
    """(label, date_from, date_to) for the window the report actually covered.

    Reports agree on the key names but not on where they sit: daily-production
    puts period_label/date_from/date_to at the top of its data dict, while
    pending-work nests the same three under data['filters']. Check both, and
    return blanks rather than guessing — All Time legitimately has no dates.
    """
    data = payload.get('data') or {}
    if not isinstance(data, dict):
        return '', '', ''

    sources = [data]
    nested = data.get('filters')
    if isinstance(nested, dict):
        sources.append(nested)

    def pick(key):
        for source in sources:
            value = source.get(key)
            if value:
                return str(value)
        return ''

    return pick('period_label'), pick('date_from'), pick('date_to')


def report_title(payload: dict) -> str:
    """Report title including the active filter slice (e.g. "... - Not Yet
    Released"), matching what the exports name themselves."""
    return build_report_title(payload)
