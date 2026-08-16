"""Resolves a staff chat message into an ORM-backed answer, then asks the
local LLM to phrase it. Intent resolution is pattern-based, not LLM-driven —
the LLM's only job is phrasing already-fetched data, never deciding what to
query, to keep this reliable on a small, slow local model.

Every identifier type below (JC, PO/WO/PR, SKU, Set no, AWC no) resolves back
to the same JobCard entity, so there is one shared facts/LLM path — per the
business rule confirmed with the user: PO, WO, and PR are the same
JobCard.PO_No field (planners fill in whichever number they have at entry
time — a PO if approved, a WO or PR indent number if not).
"""
from __future__ import annotations

import re

# Unanchored, unlike core.jc_numbering._JC_PATTERN (which is anchored with
# ^...$ for whole-string validation elsewhere and would never match a JC
# number embedded in a sentence like "what's the status of JC-07-26-PP-0701").
JC_EXTRACT_PATTERN = re.compile(r'JC-\d{2}-\d{2}-(?:PP-)?\d+(?:\.\d+)?', re.IGNORECASE)

# Fallback for "jc no 105" / "jc 105" / "jc#105" style mentions — the trailing
# number in a full JC number (JC-MM-YY-PP-####) comes from a single global
# counter (see core/jc_numbering.py: allocate_next_jc_number) that never
# resets, so the bare serial alone is enough to identify a job card uniquely;
# no month/year prefix is needed. Requires the word "jc" nearby so we don't
# treat an arbitrary number elsewhere in the message as a job card reference.
SHORT_JC_PATTERN = re.compile(
    r'\bjc\b\.?\s*(?:no\.?|number|#)?\s*[:\-]?\s*(\d{1,6})\b', re.IGNORECASE
)

# PO/WO/PR numbers are free text (not purely numeric), so capture a fairly
# permissive token — letters, digits, dashes/slashes/dots — after the keyword.
# Capped at 120 to comfortably cover the longest field this matches against
# (JobCard.SKU, max_length=100) — a lower cap here (31) previously truncated
# real SKUs mid-string, e.g. "INSERTCARD-UBPILLOWENCASEMENTBAMBOO4PACK-0426"
# (45 chars) got cut to "INSERTCARD-UBPILLOWENCASEMENTBA" and predictably
# didn't match anything.
_VALUE = r'([A-Za-z0-9][A-Za-z0-9\-/\.]{1,120})'

PO_WO_PR_PATTERN = re.compile(
    r'\b(?:po|wo|pr)\b\.?\s*(?:no\.?|number|#)?\s*[:\-]?\s*' + _VALUE, re.IGNORECASE
)
SKU_PATTERN = re.compile(
    r'\bsku\b\.?\s*(?:no\.?|number|#)?\s*[:\-]?\s*' + _VALUE, re.IGNORECASE
)
# "set" alone is too common an English word to trigger on — require the
# no/number/# suffix so casual mentions of "set" don't misfire.
SET_NO_PATTERN = re.compile(
    r'\b(?:plate\s+)?set\s*(?:no\.?|number|#)\s*[:\-]?\s*' + _VALUE, re.IGNORECASE
)
AWC_PATTERN = re.compile(
    r'\bawc\b\.?\s*(?:no\.?|number|#)?\s*[:\-]?\s*' + _VALUE, re.IGNORECASE
)

NO_MATCH_REPLY = (
    "I can look up job cards right now by JC number (e.g. JC-07-26-PP-0701, "
    "or just \"jc 105\"), PO/WO/PR number, SKU, plate set no, or AWC no — or "
    "ask for a report by name (e.g. \"stock report\", \"daily production "
    "report\"). I didn't find any of those in your message."
)


def _find_report_match(question: str):
    """Fuzzy-match a question against every registered report's title (the
    same list the bot email editor's report dropdown uses — see
    bot/report_adapter.py: available_report_choices()). Requires the word
    "report" so a casual mention of e.g. "stock" alone doesn't misfire into
    running a report. Returns (slug, title) for the best keyword-overlap
    match, or None."""
    q_words = set(re.findall(r'[a-z0-9]+', question.lower()))
    if 'report' not in q_words and 'reports' not in q_words:
        return None

    from bot.report_adapter import available_report_choices

    best = None
    best_score = 0
    for slug, title in available_report_choices():
        title_words = set(re.findall(r'[a-z0-9]+', title.lower())) - {'report', 'reports'}
        score = len(q_words & title_words)
        if score > best_score:
            best, best_score = (slug, title), score
    return best


def _find_by_serial(serial: int):
    """Look up a job card by its trailing serial alone, ignoring the
    month/year prefix — the serial is globally unique, so this is
    unambiguous. Matches on the zero-padded 4-digit form the allocator
    writes (####), with a plain fallback for serials that overflow it."""
    from core.models import JobCard

    padded = str(serial).zfill(4)
    jc = JobCard.objects.filter(job_card_no__endswith=f'-{padded}').select_related('machine_name', 'production_wip_status__status', 'planning_job').first()
    if jc is None and len(padded) != len(str(serial)):
        jc = JobCard.objects.filter(job_card_no__endswith=f'-{serial}').select_related('machine_name', 'production_wip_status__status', 'planning_job').first()
    return jc


def _find_by_po(value: str):
    from core.models import JobCard

    return (
        JobCard.objects.filter(PO_No__iexact=value)
        .select_related('machine_name', 'production_wip_status__status', 'planning_job')
        .order_by('-id')
        .first()
    )


def _find_by_sku(value: str):
    from core.models import JobCard

    return (
        JobCard.objects.filter(SKU__iexact=value)
        .select_related('machine_name', 'production_wip_status__status', 'planning_job')
        .order_by('-id')
        .first()
    )


def _find_by_set_no(value: str):
    from core.models import JobCard

    return (
        JobCard.objects.filter(plate_set_no__iexact=value)
        .select_related('machine_name', 'production_wip_status__status', 'planning_job')
        .order_by('-id')
        .first()
    )


def _find_by_awc(value: str):
    """AWC lives on the SKU master, not the job card — resolve to the SKU's
    most recent job card. Falls back to a Plate Request carrying the same
    AWC no (which links to a job card directly) for AWC codes that were
    only ever recorded on the plate side. Mirrors
    planning.models.PlanningJob.awc_no_display's own fallback order."""
    from planning.services import normalize_awc_no

    normalized = normalize_awc_no(value)
    if not normalized:
        return None

    from planning.models import SkuRecipe

    recipe = SkuRecipe.objects.filter(awc_no__iexact=normalized).first()
    if recipe:
        jc = _find_by_sku(recipe.sku)
        if jc:
            return jc

    from printing_plates.models import PlateRequest

    plate = (
        PlateRequest.objects.filter(awc_no__iexact=normalized, job_card__isnull=False)
        .select_related('job_card', 'job_card__machine_name')
        .order_by('-updated_at')
        .first()
    )
    return plate.job_card if plate else None


# (pattern, resolver, label used in the "couldn't find" message)
_LOOKUPS = [
    (PO_WO_PR_PATTERN, _find_by_po, 'PO/WO/PR number'),
    (SKU_PATTERN, _find_by_sku, 'SKU'),
    (SET_NO_PATTERN, _find_by_set_no, 'plate set no'),
    (AWC_PATTERN, _find_by_awc, 'AWC no'),
]


def _awc_no_for(jc) -> str:
    """The reverse direction of _find_by_awc: given a job, what's its AWC
    no. Mirrors PlanningJob.awc_no_display's own fallback (SKU master first,
    then a matching Plate Request) for jobs without a planning_job link."""
    if jc.planning_job_id and jc.planning_job:
        value = jc.planning_job.awc_no_display
        if value:
            return value

    from printing_plates.models import PlateRequest

    plate = PlateRequest.objects.filter(job_card=jc).exclude(awc_no='').order_by('-updated_at').first()
    if plate:
        from planning.services import normalize_awc_no
        return normalize_awc_no(plate.awc_no)
    return ''


def _facts_for(jc) -> str:
    # wip_status_name is the shop-floor stage (Printing/Packing/Dispatch/QC
    # Hold/etc, from the same ProductionWipStatus lookup table the floor
    # dashboards use) — this is "where the job is stuck," distinct from
    # workflow_status_label which is the higher-level approval-chain status.
    # total_waste/waste_percentage come from actual logged production
    # entries, not the static planning-time `wastage` estimate field.
    #
    # AWC No must be included here even though _find_by_awc already resolved
    # a job *from* an AWC — without this line, a lookup that started from
    # "awc 3382" gives the LLM a facts block with no AWC field in it at all,
    # and it will (correctly, from what it's given) say the AWC "wasn't
    # found" even though the job card was found by that exact AWC. Same gap
    # broke "jc 346 awc no?" — the LLM had no AWC fact to answer with, so it
    # answered a different field (Plate Set No) instead.
    return (
        f"Job Card: {jc.job_card_no}\n"
        f"SKU: {jc.SKU}\n"
        f"PO/WO/PR No: {jc.PO_No or 'Not set'}\n"
        f"Plate Set No: {jc.plate_set_no or 'Not set'}\n"
        f"AWC No: {_awc_no_for(jc) or 'Not set'}\n"
        f"Order Qty: {jc.order_qty}\n"
        f"Status: {jc.workflow_status_label}\n"
        f"Current Stage (shop floor): {jc.wip_status_name}\n"
        f"Machine: {jc.machine_name.name if jc.machine_name else 'Not assigned'}\n"
        f"Wastage: {jc.total_waste} sheets ({jc.waste_percentage}% of total production)\n"
        f"Dispatched: {jc.total_dispatch} of {jc.order_qty} ({jc.dispatch_completion_percent}%)\n"
    )


def _reply_for(jc, question: str) -> str:
    facts = _facts_for(jc)
    from core.llm.client import _ai_enabled
    if not _ai_enabled():
        return facts

    from core.llm.client import call_chat
    from core.llm.prompts import CHAT_ASSISTANT_SYSTEM_PROMPT

    # No lock_wait: block on core.llm.client._LLM_LOCK until any other
    # in-flight call finishes, rather than racing it and burning our own
    # LLM_TIMEOUT_SECONDS while merely queued. This runs on a background
    # thread already (see chat/api.py: _run_ai_assistant_reply), and the
    # typing indicator stays up the whole time, so a longer wait behind
    # someone else's question is fine — a silent timeout-then-raw-facts
    # fallback would be worse UX than just waiting one's turn.
    reply = call_chat([
        {'role': 'system', 'content': CHAT_ASSISTANT_SYSTEM_PROMPT},
        {'role': 'user', 'content': f'Question: {question}\n\nData:\n{facts}'},
    ])
    return reply or facts


def _reply_for_report(slug: str, title: str, question: str, user) -> str:
    """Runs the report live (as `user`, so normal report permissions apply —
    same access check as running it from the Reports screen) and phrases it
    the same way the bot email narration does (build_narration_facts +
    NARRATION_SYSTEM_PROMPT), just delivered in chat instead of an email."""
    from django.core.exceptions import PermissionDenied
    from django.http import Http404

    from bot import report_adapter

    try:
        request = report_adapter.build_stub_request(user, filters={})
        payload = report_adapter.run_report(slug, request, filters={})
    except PermissionDenied:
        return f"You don't have access to the {title} report."
    except Http404:
        return f"I couldn't find a report called {title!r} anymore — it may have been removed."

    headers, labels, rows = report_adapter.extract_rows(payload)
    facts = report_adapter.build_narration_facts(payload, headers, labels, rows, fallback_title=title)

    from core.llm.client import _ai_enabled
    if not _ai_enabled():
        return facts

    from core.llm.client import call_chat
    from core.llm.prompts import NARRATION_SYSTEM_PROMPT

    reply = call_chat([
        {'role': 'system', 'content': NARRATION_SYSTEM_PROMPT},
        {'role': 'user', 'content': f'Summarize this report:\n{facts}'},
    ])
    return reply or facts


def resolve_and_reply(question: str, user=None) -> str:
    """Always returns a string reply — never raises, never blank.

    user: the staff member asking, used to run report requests with their
    actual permissions (same access check as the Reports screen). Job card
    lookups don't need it — those aren't permission-scoped."""
    report_match = _find_report_match(question)
    if report_match:
        slug, title = report_match
        return _reply_for_report(slug, title, question, user)

    from core.models import JobCard

    match = JC_EXTRACT_PATTERN.search(question)
    if match:
        jc_no = match.group(0).upper()
        jc = JobCard.objects.filter(job_card_no=jc_no).select_related('machine_name', 'production_wip_status__status', 'planning_job').first()
        if not jc:
            return f"I couldn't find a job card numbered {jc_no}."
        return _reply_for(jc, question)

    short_match = SHORT_JC_PATTERN.search(question)
    if short_match:
        serial = int(short_match.group(1))
        jc = _find_by_serial(serial)
        if not jc:
            return f"I couldn't find a job card with number {serial}."
        return _reply_for(jc, question)

    for pattern, resolver, label in _LOOKUPS:
        found = pattern.search(question)
        if not found:
            continue
        value = found.group(1)
        jc = resolver(value)
        if not jc:
            return f"I couldn't find a job card with {label} {value}."
        return _reply_for(jc, question)

    return NO_MATCH_REPLY
