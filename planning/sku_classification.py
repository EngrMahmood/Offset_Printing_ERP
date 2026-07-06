"""Repeat vs New SKU classification (single source of truth).

Forward-only: existing planning jobs keep their stored repeat_flag unless still
unset on an unlocked draft job. See is_job_repeat_classification_locked().
"""

from __future__ import annotations

from planning.models import PlanningJob, PoDocument, SkuRecipe
from workflow.services import (
    _normalize_status,
    _po_payload_items,
    _sku_key,
)


def _normalize_po_number(raw_value):
    return (raw_value or '').strip().upper()

# Job has meaningfully entered execution — do not recalculate repeat_flag.
REPEAT_CLASSIFICATION_LOCKED_STATUSES = {
    'pending_qc',
    'qc_approved',
    'released',
    'in_production',
    'completed',
    'planning_approved',
}
REPEAT_CLASSIFICATION_LOCKED_STAGES = {
    'new_plate_making',
    'repeat_plate_making',
    'plate_received',
    'planning_done',
    'in_production',
}

LEGACY_BULK_CORE_FIELDS = (
    'material',
    'color_spec',
    'ups',
    'print_sheet_size',
    'job_name',
)


def recipe_is_bulk_like(recipe) -> bool:
    """Heuristic for Google Sheet / bulk-uploaded master rows."""
    if not recipe:
        return False
    if getattr(recipe, 'legacy_produced', False):
        return True
    filled = sum(
        1
        for field in LEGACY_BULK_CORE_FIELDS
        if str(getattr(recipe, field, '') or '').strip()
    )
    return filled >= 3


def is_job_repeat_classification_locked(job) -> bool:
    """True when repeat_flag must not be recalculated (live production safety)."""
    if not job:
        return False
    if _normalize_status(job.status) != 'draft':
        return True
    stage = (job.planning_stage or '').strip()
    return bool(stage and stage not in {'', 'draft'})


def _prior_production_on_other_po(sku, exclude_po=None) -> bool:
    key = _sku_key(sku)
    if not key:
        return False
    qs = PlanningJob.objects.filter(sku__iexact=sku, is_active=True)
    if exclude_po:
        qs = qs.exclude(po_number__iexact=exclude_po)
    for job in qs.only('status', 'planning_stage'):
        status = _normalize_status(job.status)
        stage = (job.planning_stage or '').strip()
        if status in REPEAT_CLASSIFICATION_LOCKED_STATUSES:
            return True
        if stage in REPEAT_CLASSIFICATION_LOCKED_STAGES:
            return True
    return False


def _active_job_on_other_po(sku, exclude_po=None) -> bool:
    key = _sku_key(sku)
    if not key:
        return False
    qs = PlanningJob.objects.filter(sku__iexact=sku, is_active=True)
    if exclude_po:
        qs = qs.exclude(po_number__iexact=exclude_po)
    return qs.exists()


def _earlier_po_has_sku(sku, po_number, po_doc_created_at=None, po_doc_id=None) -> bool:
    """Another PO (uploaded earlier) already contains this SKU."""
    key = _sku_key(sku)
    po_norm = _normalize_po_number(po_number)
    if not key:
        return False

    docs = PoDocument.objects.exclude(extracted_payload__isnull=True).order_by('created_at', 'id')
    for doc in docs:
        if po_doc_id and doc.id == po_doc_id:
            break
        payload = doc.extracted_payload or {}
        other_po = _normalize_po_number(payload.get('po_number') or '')
        if other_po == po_norm:
            continue
        ignored = {_sku_key(s) for s in (payload.get('new_skus_ignored') or []) if s}
        for item in _po_payload_items(payload):
            item_key = _sku_key(item.get('sku'))
            if item_key == key and item_key not in ignored:
                if po_doc_created_at is None:
                    return True
                if doc.created_at < po_doc_created_at:
                    return True
                if doc.created_at == po_doc_created_at and po_doc_id and doc.id < po_doc_id:
                    return True
    return False


def classify_po_line(
    sku,
    po_number=None,
    *,
    po_doc_created_at=None,
    po_doc_id=None,
    recipe=None,
    explicit_repeat_flag=None,
):
    """Return ('Repeat'|'New', reason_code) for one PO line.

    Priority:
      1. Manual repeat_flag on PO line
      2. legacy_produced / sheet bulk master
      3. Prior production on another PO
      4. Concurrent PO (earlier PO has same SKU)
      5. Active planning job on another PO (same SKU entered twice)
    """
    explicit = (explicit_repeat_flag or '').strip().lower()
    if explicit in {'new', 'repeat'}:
        return ('New' if explicit == 'new' else 'Repeat', 'manual')

    if recipe and getattr(recipe, 'legacy_produced', False):
        return ('Repeat', 'legacy')

    if recipe_is_bulk_like(recipe):
        return ('Repeat', 'legacy_bulk')

    if _prior_production_on_other_po(sku, exclude_po=po_number):
        return ('Repeat', 'prior_production')

    if _earlier_po_has_sku(sku, po_number, po_doc_created_at, po_doc_id):
        return ('Repeat', 'concurrent_po')

    if _active_job_on_other_po(sku, exclude_po=po_number):
        return ('Repeat', 'concurrent_job')

    return ('New', 'new')


def forward_as_new_for_po_line(
    item,
    *,
    po_number,
    po_doc_created_at=None,
    po_doc_id=None,
    recipe=None,
    existing_job=None,
):
    """Whether this PO line should carry repeat_flag='New' when creating/updating a job."""
    if existing_job and is_job_repeat_classification_locked(existing_job):
        return (existing_job.repeat_flag or '').strip().lower() != 'repeat'

    existing_flag = (existing_job.repeat_flag or '').strip().lower() if existing_job else ''
    if existing_flag in {'new', 'repeat'}:
        return existing_flag == 'new'

    explicit = (item.get('repeat_flag') or item.get('repeat') or '').strip()
    label, _reason = classify_po_line(
        item.get('sku'),
        po_number,
        po_doc_created_at=po_doc_created_at,
        po_doc_id=po_doc_id,
        recipe=recipe,
        explicit_repeat_flag=explicit,
    )
    return label == 'New'


def repeat_flag_value_for_po_line(
    item,
    *,
    po_number,
    po_doc_created_at=None,
    po_doc_id=None,
    recipe=None,
    existing_job=None,
):
    """Resolve repeat_flag string for job defaults; respects locked / stored flags."""
    if existing_job and is_job_repeat_classification_locked(existing_job):
        flag = (existing_job.repeat_flag or '').strip()
        return flag if flag else 'Repeat'

    existing_flag = (existing_job.repeat_flag or '').strip() if existing_job else ''
    if existing_flag:
        return existing_flag

    return 'New' if forward_as_new_for_po_line(
        item,
        po_number=po_number,
        po_doc_created_at=po_doc_created_at,
        po_doc_id=po_doc_id,
        recipe=recipe,
        existing_job=existing_job,
    ) else 'Repeat'


PLATE_MAKING_STAGES = frozenset({'new_plate_making', 'repeat_plate_making'})


def plate_making_stage_for_repeat_flag(repeat_flag):
    """Map repeat_flag to the correct plate-making planning stage."""
    return 'new_plate_making' if (repeat_flag or '').strip() == 'New' else 'repeat_plate_making'


def sync_plate_making_stage_with_repeat_flag(job, *, save=False):
    """Align planning_stage with repeat_flag when job is in plate making."""
    stage = (job.planning_stage or '').strip()
    if stage not in PLATE_MAKING_STAGES:
        return False
    expected = plate_making_stage_for_repeat_flag(job.repeat_flag)
    if stage == expected:
        return False
    job.planning_stage = expected
    if save:
        job.save(update_fields=['planning_stage', 'updated_at'])
    return True


def repair_inconsistent_plate_making_stages(queryset=None):
    """Fix jobs where repeat_flag and planning_stage disagree."""
    from planning.models import PlanningJob

    qs = queryset or PlanningJob.objects.filter(planning_stage__in=PLATE_MAKING_STAGES)
    fixed = 0
    for job in qs.iterator():
        if sync_plate_making_stage_with_repeat_flag(job, save=True):
            fixed += 1
    return fixed


def annotate_items_repeat_new(items, recipe_map, *, po_number=None, po_doc_created_at=None, po_doc_id=None):
    """Annotate PO items with is_repeat / recipe_status and return counts + missing master SKUs."""
    annotated = []
    repeat_count = 0
    new_count = 0
    missing_skus = []

    for item in items:
        sku = (item.get('sku') or '').strip()
        key = _sku_key(sku)
        recipe = recipe_map.get(key) if key else None

        label, _reason = classify_po_line(
            sku,
            po_number,
            po_doc_created_at=po_doc_created_at,
            po_doc_id=po_doc_id,
            recipe=recipe,
            explicit_repeat_flag=(item.get('repeat_flag') or item.get('repeat')),
        )
        is_repeat = label == 'Repeat'

        item_copy = dict(item)
        item_copy['is_repeat'] = is_repeat
        item_copy['recipe_status'] = label
        annotated.append(item_copy)

        is_recipe_approved = bool(recipe and recipe.master_data_status == 'approved')
        if is_repeat:
            repeat_count += 1
        else:
            new_count += 1
            if not is_recipe_approved and sku:
                missing_skus.append(sku)

    return annotated, repeat_count, new_count, sorted(set(missing_skus))
