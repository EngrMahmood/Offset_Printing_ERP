"""Coating (UV / varnish) helpers: which jobs need it, how much is done, what is still pending."""
from __future__ import annotations

from django.db.models import Count, Q, Sum

from core.models import JOB_CARD_PRODUCTION_CONTINUE_STATUSES, JobCard

from .models import CoatingRun

COATING_KEYWORDS = ('uv', 'varnish')
# Coated sheets may run this far past the printed good sheets (counter drift) before entry is refused.
OVERRUN_PCT = 1


def coating_kind(application):
    """'uv' | 'matt_varnish' | 'gloss_varnish' | 'other' for an application label, '' when it is no coating
    (lamination and NO are separate processes and stay out)."""
    text = (application or '').strip().lower()
    if not text or 'lamination' in text:
        return ''
    if 'uv' in text.replace('-', ' ').split() or text.startswith('uv'):
        return 'uv'
    if 'varnish' in text:
        return 'matt_varnish' if 'matt' in text or 'matte' in text else 'gloss_varnish' if 'gloss' in text else 'other'
    return ''


def job_application(job_card):
    planning_job = getattr(job_card, 'planning_job', None)
    return planning_job.application_display if planning_job else ''


def needs_coating(job_card):
    return bool(coating_kind(job_application(job_card)))


BOTH_SIDE_WORDS = ('front', 'back', 'both', 'f/b', 'f&b', 'fb', '2 side', 'two side', 'double')


def sides_from_application(application):
    """2 when the application says the sheet is coated on both sides (front and back), else 1."""
    text = (application or '').lower()
    return 2 if any(word in text for word in BOTH_SIDE_WORDS) else 1


def summary(job_card, exclude_run_id=None, new_coats=1):
    """Printed good sheets vs coated so far for one job card.

    A sheet coated on both sides (front and back) is two passes through the press, so the limit is counted in
    passes (sheets x coats): printed sheets x sides, where sides is 2 for a front/back job, otherwise the most
    coats any run of the job used (and the one being entered)."""
    runs = job_card.coating_runs.filter(is_active=True)
    if exclude_run_id:
        runs = runs.exclude(pk=exclude_run_id)
    rows = list(runs.values_list('sheets_coated', 'waste_sheets', 'coats', 'status'))
    printed = int(job_card.total_production or 0)
    passes = sum(r[0] * r[2] for r in rows)
    waste = sum(r[1] for r in rows)
    sides = max([sides_from_application(job_application(job_card)), int(new_coats or 1)] + [r[2] for r in rows])
    coated = -(-passes // sides) if passes else 0                      # sheets fully coated, whatever the sides
    overrun = (printed * OVERRUN_PCT + 99) // 100
    pass_limit = printed * sides + overrun * sides
    return {
        'printed': printed,
        'coated': coated,
        'waste': waste,
        'sides': sides,
        'passes': passes,
        'pass_limit': pass_limit,
        'remaining': max(printed - coated, 0),
        'allowed_more': max(pass_limit - passes, 0),                  # passes (sheets x coats) that may still be entered
        'sheet_cap': printed + overrun,                               # one run can never exceed the printed sheets
        'pct': round(coated * 100 / printed, 1) if printed else 0,
        'marked_done': any(r[3] == 'completed' for r in rows),
        'runs': len(rows),
    }


def coating_jobs(include_all=False):
    """Job cards that can take a coating entry: printing has started and the job's application is a coating."""
    qs = JobCard.objects.filter(is_active=True, status__in=JOB_CARD_PRODUCTION_CONTINUE_STATUSES, productions__entry_type='printing', productions__is_active=True)
    if not include_all:
        cond = Q()
        for word in COATING_KEYWORDS:
            cond |= Q(planning_job__application__icontains=word)
        qs = qs.filter(cond).exclude(planning_job__application__icontains='lamination')
    return qs.select_related('planning_job', 'material').distinct()


def pending_rows(limit=None):
    """Coating jobs with printed sheets not coated yet, oldest first."""
    rows = []
    for card in coating_jobs().order_by('created_at'):
        info = summary(card)
        if info['marked_done'] or info['remaining'] <= 0:
            continue
        rows.append({'job_card': card, 'kind': coating_kind(job_application(card)), **info})
        if limit and len(rows) >= limit:
            break
    return rows
