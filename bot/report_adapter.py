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


def _as_number(value):
    """Parse a report cell into a float, or None when it isn't numeric.

    Report cells reach here already serialised to strings, and quantity
    columns are commonly formatted with thousands separators, so strip those
    before parsing or every such column would look non-numeric.
    """
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(',', '').replace('%', '')
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def _numeric_column_totals(headers, labels, rows):
    """Sum and average of every numeric column, computed over ALL rows.

    Exists because the row list handed to the model is only a sample (see
    build_narration_facts) — without real aggregates in front of it, a model
    asked to summarise a 287-row report will happily invent a plausible-looking
    total from the 25 rows it can see. Giving it the true figures is what keeps
    the emailed narration honest; the prompt rule against inventing numbers is
    a backstop, not a substitute.
    """
    totals = []
    for header in headers:
        numbers = [
            number
            for number in (_as_number(row.get(header)) for row in rows)
            if number is not None
        ]
        # Mostly-numeric only: a column of job card numbers or PO refs must
        # never be summed just because a few of them happen to parse.
        if not numbers or len(numbers) < len(rows) * 0.6:
            continue
        label = labels.get(header, header)
        total = sum(numbers)
        average = total / len(numbers)
        totals.append(f'{label}: total {total:,.0f} across {len(numbers)} record(s), average {average:,.1f}')
    return totals


def build_narration_facts(payload, headers, labels, rows, fallback_title='', max_rows=25) -> str:
    """Row-capped, LLM-prompt-ready facts block for a report payload. Shared
    by bot email narration (bot/services.py) and the chat AI assistant
    (chat/ai_assistant.py) so both phrase the same underlying data the same
    way. Capped — this is a slow, small-context local model; a 500-row table
    would blow the context window and response-time budget for no benefit
    (the full table is available elsewhere: the email attachment, or the
    report screen itself).

    The row cap is why the totals block below is computed here in Python over
    every row rather than left to the model to work out from the sample.
    """
    sample = rows[:max_rows]
    lines = [
        f"Report: {report_title(payload) if payload else fallback_title}",
        f"Total records: {len(rows)}",
    ]

    totals = _numeric_column_totals(headers, labels, rows)
    if totals:
        lines.append('')
        lines.append('Verified totals for ALL records (use these exact figures for any total or average):')
        lines.extend(f'- {line}' for line in totals)
        lines.append('')

    if sample:
        lines.append(
            f'Sample of {len(sample)} record(s) below'
            + (' — NOT the full list, never add these up:' if len(rows) > len(sample) else ':')
        )
    for row in sample:
        lines.append(', '.join(f'{labels.get(h, h)}: {row.get(h)}' for h in headers))
    if len(rows) > len(sample):
        lines.append(f'... and {len(rows) - len(sample)} more row(s) not shown here.')
    return '\n'.join(lines)


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
