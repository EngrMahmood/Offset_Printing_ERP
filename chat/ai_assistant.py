"""Resolves a staff chat message into an ORM-backed answer, then asks the
local LLM to phrase it. Intent resolution is pattern-based, not LLM-driven —
the LLM's only job is phrasing already-fetched data, never deciding what to
query, to keep this reliable on a small, slow local model.

The JobCard identifier types (JC, PO/WO/PR, SKU, Set no, AWC no) all resolve
back to the same JobCard entity, so they share one facts/LLM path
(_reply_for) — per the business rule confirmed with the user: PO, WO, and PR
are the same JobCard.PO_No field (planners fill in whichever number they
have at entry time — a PO if approved, a WO or PR indent number if not).

Dispatch (by DC no), Task (by number), and raw material stock (by material
SKU) are separate entities with their own facts-building + reply functions,
sharing only the LLM-phrasing tail (_phrase_answer) with the JobCard path —
a DC no in particular can span multiple JobCards, so that one resolves to a
list rather than a single record.
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
# "dc" alone is too short/ambiguous to trigger on — require the no/number/#
# suffix, same reasoning as SET_NO_PATTERN.
DC_NO_PATTERN = re.compile(
    r'\bdc\s*(?:no\.?|number|#)\s*[:\-]?\s*' + _VALUE, re.IGNORECASE
)
# Numeric-id style, identical shape to SHORT_JC_PATTERN — Task has no
# formatted code, only a free-text title and the Django pk.
TASK_ID_PATTERN = re.compile(
    r'\btask\b\.?\s*(?:no\.?|number|#)?\s*[:\-]?\s*(\d{1,6})\b', re.IGNORECASE
)
# Deliberately distinct wording ("material"/"raw material") from SKU_PATTERN
# so the two never collide on the same message — RawMaterialSku is a
# separate inventory domain from JobCard.SKU (finished goods). Its own value
# capture, not _VALUE — unlike every other identifier here, real
# RawMaterialSku.sku values routinely contain spaces (e.g. "RUBBER COVERING
# OF SM-74 ROLLER SIZE DIA 75MM"), so this captures to the end of the
# message rather than stopping at whitespace; resolve_and_reply strips
# trailing punctuation before using it as a lookup value.
MATERIAL_SKU_PATTERN = re.compile(
    r'\b(?:raw\s+)?material\b\.?\s*(?:no\.?|number|#)?\s*[:\-]?\s*(.+)$', re.IGNORECASE
)

NO_MATCH_REPLY = (
    "I can look up job cards right now by JC number (e.g. JC-07-26-PP-0701, "
    "or just \"jc 105\"), PO/WO/PR number, SKU, plate set no, or AWC no. I "
    "can also look up a dispatch by DC number, a task by number (e.g. "
    "\"task 5\"), or raw material stock by material SKU — or ask for a "
    "report by name (e.g. \"stock report\", \"daily production report\"). "
    "I didn't find any of those in your message."
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


def _find_dispatches_by_dc(value: str):
    """A DC number is not globally unique — the same DC can legitimately
    cover multiple job cards (per Dispatch's own docstring) — so this
    returns a list, not a single record, unlike every other resolver here."""
    from core.models import Dispatch

    return list(
        Dispatch.objects.filter(dc_no__iexact=value, is_active=True)
        .select_related('job_card')
        .order_by('-dispatch_date')
    )


def _find_task_by_id(task_id: int):
    from tasks.models import Task

    return Task.objects.filter(pk=task_id).select_related('assignee', 'assigned_team').first()


def _find_raw_material_sku(value: str):
    from supply_chain.models import RawMaterialSku

    return (
        RawMaterialSku.objects.filter(sku__iexact=value, is_active=True)
        .select_related('material')
        .first()
    )


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


def _phrase_answer(facts: str, question: str) -> str:
    """Shared LLM-phrasing tail for every conversational (question-and-
    answer, as opposed to report-narration) reply in this module — checks
    the AI enable switch, calls the LLM, falls back to raw facts on failure
    or when disabled. Never raises, never blank.

    No lock_wait: blocks on core.llm.client._LLM_LOCK until any other
    in-flight call finishes, rather than racing it and burning our own
    LLM_TIMEOUT_SECONDS while merely queued. This runs on a background
    thread already (see chat/api.py: _run_ai_assistant_reply), and the
    typing indicator stays up the whole time, so a longer wait behind
    someone else's question is fine — a silent timeout-then-raw-facts
    fallback would be worse UX than just waiting one's turn."""
    from core.llm.client import _ai_enabled
    if not _ai_enabled():
        return facts

    from core.llm.client import call_chat
    from core.llm.prompts import CHAT_ASSISTANT_SYSTEM_PROMPT

    reply = call_chat([
        {'role': 'system', 'content': CHAT_ASSISTANT_SYSTEM_PROMPT},
        {'role': 'user', 'content': f'Question: {question}\n\nData:\n{facts}'},
    ])
    return reply or facts


def _reply_for(jc, question: str) -> str:
    return _phrase_answer(_facts_for(jc), question)


def _facts_for_dispatches(dispatches: list, dc_no: str) -> str:
    sample = dispatches[:25]
    lines = [f"DC No: {dc_no}", f"Total dispatch records: {len(dispatches)}"]
    for d in sample:
        lines.append(
            f"Job Card: {d.job_card.job_card_no}, SKU: {d.job_card.SKU}, "
            f"Dispatch Qty: {d.dispatch_qty}, Date: {d.dispatch_date}"
        )
    if len(dispatches) > len(sample):
        lines.append(f'... and {len(dispatches) - len(sample)} more record(s) not shown here.')
    return '\n'.join(lines)


def _reply_for_dispatches(dispatches: list, dc_no: str, question: str) -> str:
    return _phrase_answer(_facts_for_dispatches(dispatches, dc_no), question)


def _facts_for_task(task) -> str:
    if task.assignee:
        assignee = task.assignee.get_full_name() or task.assignee.username
    else:
        assignee = 'Not assigned'
    return (
        f"Task: {task.title}\n"
        f"Description: {task.description or 'Not set'}\n"
        f"Assignee: {assignee}\n"
        f"Team: {task.assigned_team.name if task.assigned_team else 'Not assigned'}\n"
        f"Priority: {task.get_priority_display()}\n"
        f"Status: {task.get_status_display()}\n"
        f"Due Date: {task.due_date}\n"
        f"Score: {task.score if task.score is not None else 'Not scored yet'}\n"
    )


def _reply_for_task(task, question: str) -> str:
    return _phrase_answer(_facts_for_task(task), question)


def _facts_for_raw_material(sku_obj) -> str:
    from supply_chain.demand_gap import _on_hand_by_sku

    on_hand = _on_hand_by_sku([sku_obj.pk]).get(sku_obj.pk, 0)
    return (
        f"Material SKU: {sku_obj.sku}\n"
        f"Material: {sku_obj.material.name}\n"
        f"Purchase Sheet Size: {sku_obj.purchase_sheet_size}\n"
        f"On-hand Stock: {on_hand} {sku_obj.uom}\n"
        f"Safety Stock: {sku_obj.safety_stock}\n"
        f"Max Stock Level: {sku_obj.max_stock_level}\n"
        f"Unit Cost: {sku_obj.unit_cost}\n"
    )


def _reply_for_material(sku_obj, question: str) -> str:
    return _phrase_answer(_facts_for_raw_material(sku_obj), question)


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

    # These three don't resolve to a JobCard, so they can't reuse _LOOKUPS
    # (which assumes a single JobCard result) — each gets its own branch.
    dc_match = DC_NO_PATTERN.search(question)
    if dc_match:
        dc_no = dc_match.group(1)
        dispatches = _find_dispatches_by_dc(dc_no)
        if not dispatches:
            return f"I couldn't find any dispatch records with DC no {dc_no}."
        return _reply_for_dispatches(dispatches, dc_no, question)

    task_match = TASK_ID_PATTERN.search(question)
    if task_match:
        task_id = int(task_match.group(1))
        task = _find_task_by_id(task_id)
        if not task:
            return f"I couldn't find a task with number {task_id}."
        return _reply_for_task(task, question)

    material_match = MATERIAL_SKU_PATTERN.search(question)
    if material_match:
        value = material_match.group(1).strip().rstrip('?.!,')
        sku_obj = _find_raw_material_sku(value)
        if not sku_obj:
            return f"I couldn't find a raw material with SKU {value}."
        return _reply_for_material(sku_obj, question)

    return NO_MATCH_REPLY
