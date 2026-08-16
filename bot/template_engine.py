"""Renders a bot's subject/body templates.

Uses django.template with an explicitly-built context — the templates are
editable by admins from the ERP UI, so they get a fixed, documented set of
variables and no access to the ORM.
"""
from __future__ import annotations

import html
from datetime import datetime

from django.template import Context, Template
from django.utils.safestring import mark_safe


# Documented in the UI so the editor shows exactly what is available.
SUPPORTED_VARIABLES = [
    ('{{today}}', 'Today\'s date, e.g. 11 Aug 2026'),
    ('{{date}}', 'Same as {{today}} — the report date'),
    ('{{time}}', 'Time the bot ran, e.g. 08:30 AM'),
    ('{{report_title}}', 'Title of the report being sent'),
    ('{{report_table}}', 'The report data as an HTML table'),
    ('{{total_records}}', 'Number of rows in the report'),
    ('{{bot_name}}', 'Name of this automation'),
    ('{{user_name}}', 'Full name of the user the report runs as'),
    ('{{department}}', 'Department of that user, if set'),
    ('{{filters_summary}}', 'Human-readable summary of the applied filters'),
    ('{{period_label}}', 'Period the report covers, e.g. Yesterday or This Month'),
    ('{{period_from}}', 'First day of that period, e.g. 12 Aug 2026 (blank for All Time)'),
    ('{{period_to}}', 'Last day of that period'),
    ('{{ai_summary}}', 'AI-generated plain-English summary of this report (blank if AI summary is off or unavailable)'),
]

_TABLE_STYLE = (
    'border-collapse:collapse;width:100%;font-family:Segoe UI,Arial,sans-serif;'
    'font-size:13px;margin:12px 0;'
)
_TH_STYLE = (
    'background:#2f6ea7;color:#ffffff;text-align:left;padding:8px 10px;'
    'border:1px solid #255986;white-space:nowrap;'
)
_TD_STYLE = 'padding:7px 10px;border:1px solid #d7dee6;'
_TD_ALT_STYLE = _TD_STYLE + 'background:#f5f8fb;'


def _cell(value) -> str:
    if value is None:
        return ''
    return html.escape(str(value))


def build_html_table(headers, labels, rows, max_rows=None) -> str:
    """Inline-styled HTML table — email clients strip <style> blocks, so every
    rule has to live on the element."""
    if not rows:
        return '<p style="font-family:Segoe UI,Arial,sans-serif;color:#64748b;">No pending records.</p>'

    visible = rows if not max_rows else rows[:max_rows]
    hidden = len(rows) - len(visible)

    parts = [f'<table style="{_TABLE_STYLE}" cellspacing="0" cellpadding="0" border="0">']
    parts.append('<thead><tr>')
    parts.append(f'<th style="{_TH_STYLE}">#</th>')
    for header in headers:
        parts.append(f'<th style="{_TH_STYLE}">{_cell(labels.get(header, header))}</th>')
    parts.append('</tr></thead><tbody>')

    for index, row in enumerate(visible, start=1):
        style = _TD_ALT_STYLE if index % 2 == 0 else _TD_STYLE
        parts.append('<tr>')
        parts.append(f'<td style="{style}">{index}</td>')
        for header in headers:
            parts.append(f'<td style="{style}">{_cell(row.get(header))}</td>')
        parts.append('</tr>')

    parts.append('</tbody>')
    if hidden > 0:
        colspan = len(headers) + 1
        parts.append(
            f'<tfoot><tr><td colspan="{colspan}" '
            f'style="{_TD_STYLE}color:#64748b;font-style:italic;">'
            f'… and {hidden} more row(s) — see the attached report.</td></tr></tfoot>'
        )
    parts.append('</table>')
    return ''.join(parts)


def build_text_table(headers, labels, rows, max_rows=None) -> str:
    """Fixed-width rendering for the plain-text alternative part."""
    if not rows:
        return 'No pending records.'

    visible = rows if not max_rows else rows[:max_rows]
    hidden = len(rows) - len(visible)

    columns = ['#'] + [str(labels.get(header, header)) for header in headers]
    table = [columns]
    for index, row in enumerate(visible, start=1):
        table.append([str(index)] + ['' if row.get(h) is None else str(row.get(h)) for h in headers])

    widths = [max(len(row[i]) for row in table) for i in range(len(columns))]
    lines = [
        '  '.join(value.ljust(widths[i]) for i, value in enumerate(table[0])),
        '  '.join('-' * width for width in widths),
    ]
    lines.extend('  '.join(value.ljust(widths[i]) for i, value in enumerate(row)) for row in table[1:])
    if hidden > 0:
        lines.append(f'... and {hidden} more row(s) - see the attached report.')
    return '\n'.join(lines)


def summarize_filters(filters: dict) -> str:
    if not filters:
        return 'None'
    parts = [
        f"{key.replace('_', ' ').title()}: {value}"
        for key, value in sorted(filters.items())
        if value not in (None, '', [], {})
    ]
    return ', '.join(parts) or 'None'


def _friendly_date(iso_date: str) -> str:
    """'2026-08-12' -> '12 Aug 2026', matching {{date}}. Anything else passes
    through untouched — a report is free to hand back a label of its own."""
    if not iso_date:
        return ''
    try:
        return datetime.strptime(iso_date, '%Y-%m-%d').strftime('%d %b %Y')
    except ValueError:
        return iso_date


def build_context(bot, payload, headers, labels, rows, now, table_html=None, table_text=None, ai_summary='') -> dict:
    """The full variable set available to subject/body templates."""
    run_as = bot.resolve_run_as_user()
    profile = getattr(run_as, 'profile', None)
    department = getattr(getattr(profile, 'department', None), 'name', '') or ''

    from bot.report_adapter import period_info, report_title

    period_label, period_from, period_to = period_info(payload or {})
    if not period_label and bot.report_period:
        # No payload (the form's validation probe) or a report that doesn't
        # report its own window — the bot still knows what it asked for.
        period_label = bot.get_report_period_display()

    if table_html is None:
        table_html = build_html_table(headers, labels, rows, bot.max_rows_in_body)
    if table_text is None:
        table_text = build_text_table(headers, labels, rows, bot.max_rows_in_body)

    date_str = now.strftime('%d %b %Y')
    return {
        'today': date_str,
        'date': date_str,
        'time': now.strftime('%I:%M %p'),
        'report_title': report_title(payload) if payload else '',
        'report_table': mark_safe(table_html),
        'report_table_text': table_text,
        'total_records': len(rows),
        'bot_name': bot.name,
        'user_name': (run_as.get_full_name() or run_as.username) if run_as else '',
        'department': department,
        'filters_summary': summarize_filters(bot.effective_filters()),
        'period_label': period_label,
        'period_from': _friendly_date(period_from),
        'period_to': _friendly_date(period_to),
        'ai_summary': ai_summary,
    }


def render_template(template_text: str, context: dict) -> str:
    """Render one of the admin-editable templates. Django's template engine
    raises TemplateSyntaxError on a malformed template — the caller records
    that on the execution row rather than letting it reach the scheduler."""
    return Template(template_text or '').render(Context(context))


def render_subject(bot, context: dict) -> str:
    # Newlines in a Subject header would split it — collapse them.
    rendered = render_template(bot.subject_template, context)
    return ' '.join(rendered.split())


def render_body(bot, context: dict) -> str:
    return render_template(bot.body_template, context)


def render_text_body(bot, context: dict) -> str:
    """Plain-text alternative: same template with the fixed-width table
    swapped in for the HTML one, then tags stripped."""
    from django.utils.html import strip_tags

    text_context = dict(context)
    text_context['report_table'] = f'\n{context.get("report_table_text", "")}\n'
    rendered = render_template(bot.body_template, text_context)
    return html.unescape(strip_tags(rendered)).strip()
