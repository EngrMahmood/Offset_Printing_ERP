import io
import json
import re
from datetime import datetime, date
from decimal import Decimal, ROUND_HALF_UP
from collections import defaultdict
from difflib import SequenceMatcher
from django.db import transaction
from django.db.models import Sum, Q
from django.db.models.functions import Upper
from django.utils import timezone
from core.models import Machine, Department, Material
from .models import PLANNING_STATUS_ALIASES, PlanningJob, PoDocument, SkuRecipe
from workflow.services import _append_unique_note_line, _parse_iso_date, _format_display_qty, _build_cost_mismatch_note, _normalize_status, _to_int, _to_decimal, SKU_MASTER_APPROVAL_REQUIRED_FIELDS
from core.jc_numbering import allocate_next_jc_number

NEW_SKU_REQUIREMENT_NOTE = 'NEW SKU: Shade matching and setup verification required before production run.'


def normalize_die_cutting(raw_value):
    """Die cutting is Yes/No only (boolean-style). Legacy die names map to YES."""
    if raw_value is None:
        return ''
    if isinstance(raw_value, bool):
        return 'YES' if raw_value else 'NO'
    text = str(raw_value).strip()
    if not text:
        return ''
    lowered = text.lower()
    if lowered in {'no', 'none', 'n/a', 'na', 'nil', 'false', '0', 'n', 'not applicable'}:
        return 'NO'
    if lowered in {'yes', 'y', 'true', '1'}:
        return 'YES'
    # Historical free-text die names mean die is used.
    return 'YES'


def normalize_awc_no(raw_value):
    """
    AWC # is always a free-text string (letters, digits, punctuation).

    Excel often delivers numeric codes as floats (1050.0); store as '1050', never '1050.0'.
    Values like 'AWC-777' or '12A' are kept as-is.
    """
    if raw_value is None:
        return ''
    if isinstance(raw_value, bool):
        return ''
    if isinstance(raw_value, int):
        return str(raw_value)
    if isinstance(raw_value, float):
        if raw_value != raw_value:  # NaN
            return ''
        if raw_value == int(raw_value):
            return str(int(raw_value))
        text = format(raw_value, 'f').rstrip('0').rstrip('.')
        return text
    if isinstance(raw_value, Decimal):
        if raw_value == raw_value.to_integral_value():
            return str(int(raw_value))
        return format(raw_value, 'f').rstrip('0').rstrip('.')

    text = str(raw_value).strip()
    if not text or text.lower() in {'none', 'nan'}:
        return ''
    # Excel/CSV artifact: "3483.0" / "3483.00"
    if re.fullmatch(r'\d+\.0+', text):
        return text.split('.', 1)[0]
    return text


def get_awc_conflict_message(awc_no, *, sku='', exclude_recipe_id=None, exclude_plate_request_id=None):
    """AWC is unique per design/SKU. Same SKU may reuse its code (repeat/remake)."""
    awc = normalize_awc_no(awc_no)
    if not awc:
        return ''

    sku_key = (sku or '').strip()

    recipe_qs = SkuRecipe.objects.filter(awc_no__iexact=awc).exclude(awc_no='')
    if exclude_recipe_id:
        recipe_qs = recipe_qs.exclude(pk=exclude_recipe_id)
    if sku_key:
        recipe_qs = recipe_qs.exclude(sku__iexact=sku_key)
    other_recipe = recipe_qs.order_by('id').first()
    if other_recipe:
        return (
            f'AWC # "{awc}" is already assigned to SKU {other_recipe.sku}. '
            'Artwork codes are unique per design/SKU.'
        )

    from printing_plates.models import PlateRequest

    plate_qs = (
        PlateRequest.objects.filter(awc_no__iexact=awc)
        .exclude(awc_no='')
        .select_related('planning_job', 'sku_recipe', 'job_card')
        .order_by('id')
    )
    if exclude_plate_request_id:
        plate_qs = plate_qs.exclude(pk=exclude_plate_request_id)

    for req in plate_qs[:100]:
        req_sku = ''
        if req.sku_recipe_id and req.sku_recipe:
            req_sku = (req.sku_recipe.sku or '').strip()
        if not req_sku and req.planning_job_id and req.planning_job:
            req_sku = (req.planning_job.sku or '').strip()
        if not req_sku and req.job_card_id and req.job_card:
            req_sku = (getattr(req.job_card, 'SKU', None) or '').strip()

        if sku_key and req_sku and req_sku.lower() == sku_key.lower():
            continue
        if req_sku and (not sku_key or req_sku.lower() != sku_key.lower()):
            return (
                f'AWC # "{awc}" is already assigned to SKU {req_sku}. '
                'Artwork codes are unique per design/SKU.'
            )
        if not req_sku and sku_key:
            return (
                f'AWC # "{awc}" is already used on another plate request. '
                'Artwork codes are unique per design/SKU.'
            )
    return ''


def _user_is_admin(user):
    profile = getattr(user, 'profile', None)
    return getattr(user, 'is_superuser', False) or (profile is not None and profile.normalized_role == 'admin')


def _user_is_graphics_designer(user):
    profile = getattr(user, 'profile', None)
    return profile is not None and profile.normalized_role == 'graphics_designer'


SKU_RECIPE_DESIGNER_FIELDS = [
    'color_spec',
    'size_w_mm',
    'size_h_mm',
    'ups',
    'print_sheet_size',
    'purchase_sheet_size',
    'purchase_sheet_ups',
    'awc_no',
    'die_cutting',
    'plate_set_no',
]

SKU_RECIPE_PLANNER_FIELDS = [
    'job_process_type',
    'material',
    'application',
    'machine_name',
    'product_type',
    'default_unit_cost',
    'daily_demand',
    'notes',
    'remarks',
]

SKU_RECIPE_SHARED_FIELDS = [
    'sku',
    'job_name',
]

SKU_RECIPE_FIELD_ROLE_LABELS = {
    'planner': 'Planner',
    'designer': 'Designer',
    'shared': 'Shared',
}


def get_sku_recipe_field_role(field_name):
    if field_name in SKU_RECIPE_PLANNER_FIELDS:
        return 'planner'
    if field_name in SKU_RECIPE_DESIGNER_FIELDS:
        return 'designer'
    return 'shared'


def sku_recipe_field_editable_by_user(field_name, user, *, is_readonly=False, recipe=None):
    if is_readonly:
        return False
    if _user_is_admin(user):
        return True

    # SKU and job name are identity fields — admin only.
    if field_name in SKU_RECIPE_SHARED_FIELDS:
        return False

    role = get_sku_recipe_field_role(field_name)
    if _user_is_graphics_designer(user):
        return role == 'designer'
    if role == 'designer':
        return planner_can_edit_designer_fields(recipe)
    return role == 'planner'


def get_sku_recipe_form_ui_context(user, *, is_readonly=False):
    if is_readonly:
        viewer_role = 'readonly'
    elif _user_is_admin(user):
        viewer_role = 'admin'
    elif _user_is_graphics_designer(user):
        viewer_role = 'designer'
    else:
        viewer_role = 'planner'

    return {
        'sku_recipe_viewer_role': viewer_role,
        'sku_recipe_planner_fields': SKU_RECIPE_PLANNER_FIELDS,
        'sku_recipe_designer_fields': SKU_RECIPE_DESIGNER_FIELDS,
    }


def planner_can_edit_designer_fields(recipe):
    """
    After change management (Reopen SKU / send back to Draft), planner may edit
    designer layout fields (machine-driven plate size, print sheet size, etc.).
    First-time drafts keep designer fields for graphics only.
    """
    if not recipe or recipe.master_data_status == 'approved':
        return False
    return bool(recipe.last_rejected_at or (recipe.rejection_comment or '').strip())


def enrich_sku_recipe_form_ui(form, user, *, is_readonly=False, recipe=None):
    """Attach role metadata used by SKU master templates for highlights and badges."""
    for field_name, field in form.fields.items():
        role = get_sku_recipe_field_role(field_name)
        is_mine = sku_recipe_field_editable_by_user(
            field_name, user, is_readonly=is_readonly, recipe=recipe,
        )
        field.sku_role = role
        field.sku_role_label = SKU_RECIPE_FIELD_ROLE_LABELS[role]
        field.sku_is_mine = is_mine
        css_class = field.widget.attrs.get('class', '')
        marker = f'sku-field--{role}'
        state = 'is-your-role' if is_mine else 'is-other-role'
        field.widget.attrs['class'] = f'{css_class} {marker} {state}'.strip()
    return form


def apply_sku_recipe_form_role_permissions(form, user, *, is_readonly=False, recipe=None):
    """Restrict SKU master fields by role: planners edit planner fields, designers edit layout fields."""
    if is_readonly:
        for field in form.fields.values():
            field.disabled = True
        return enrich_sku_recipe_form_ui(form, user, is_readonly=True, recipe=recipe)

    if _user_is_admin(user):
        return enrich_sku_recipe_form_ui(form, user, is_readonly=False, recipe=recipe)

    # SKU + job name locked for everyone except admin.
    for field_name in SKU_RECIPE_SHARED_FIELDS:
        if field_name in form.fields:
            form.fields[field_name].disabled = True

    if _user_is_graphics_designer(user):
        for field_name in SKU_RECIPE_PLANNER_FIELDS:
            if field_name in form.fields:
                form.fields[field_name].disabled = True
        return enrich_sku_recipe_form_ui(form, user, is_readonly=False, recipe=recipe)

    # Planner: unlock designer fields only during change management (reopened SKU).
    if not planner_can_edit_designer_fields(recipe):
        for field_name in SKU_RECIPE_DESIGNER_FIELDS:
            if field_name in form.fields:
                form.fields[field_name].disabled = True

    return enrich_sku_recipe_form_ui(form, user, is_readonly=is_readonly, recipe=recipe)


def merge_preserved_sku_recipe_fields(posted, recipe, user):
    """Keep locked field values when the current user cannot edit them."""
    if not recipe:
        return posted

    if not _user_is_admin(user):
        for field_name in SKU_RECIPE_SHARED_FIELDS:
            value = getattr(recipe, field_name, None)
            if value is not None and str(value).strip():
                posted[field_name] = str(value)

    if _user_is_admin(user):
        return posted

    # Designer posts must keep planner values that were not editable.
    if _user_is_graphics_designer(user):
        for field_name in SKU_RECIPE_PLANNER_FIELDS:
            if field_name not in posted or not str(posted.get(field_name) or '').strip():
                value = getattr(recipe, field_name, None)
                if value is not None and str(value).strip():
                    posted[field_name] = str(value)
        return posted

    # Planner posts must keep designer values unless change-management unlock.
    if planner_can_edit_designer_fields(recipe):
        return posted

    for field_name in SKU_RECIPE_DESIGNER_FIELDS:
        if field_name not in posted or not str(posted.get(field_name) or '').strip():
            value = getattr(recipe, field_name, None)
            if value is not None and str(value).strip():
                posted[field_name] = str(value)
    return posted


def _field_is_blank(value):
    return value is None or (isinstance(value, str) and not str(value).strip())


def _restore_fields_from_recipe(obj, recipe, field_names):
    for field_name in field_names:
        new_val = getattr(obj, field_name, None)
        old_val = getattr(recipe, field_name, None)
        if _field_is_blank(new_val) and not _field_is_blank(old_val):
            setattr(obj, field_name, old_val)
    return obj


def restore_locked_designer_fields_on_recipe(obj, recipe, user):
    """
    After form.save(commit=False), put back values the current user cannot edit.

    Disabled HTML inputs are not posted; Django may also ignore them. This keeps
    bulk-uploaded AWC / die cutting / layout / planner fields from being wiped.
    """
    if not obj or not recipe:
        return obj
    if _user_is_admin(user):
        return obj

    # Designer must not wipe planner fields.
    if _user_is_graphics_designer(user):
        return _restore_fields_from_recipe(obj, recipe, SKU_RECIPE_PLANNER_FIELDS)

    # Planner must not wipe designer fields (unless change-management unlock).
    if planner_can_edit_designer_fields(recipe):
        return obj
    return _restore_fields_from_recipe(obj, recipe, SKU_RECIPE_DESIGNER_FIELDS)


def _safe_color_spec(*candidates):
    """First non-blank print color that is not plate-ink chip text."""
    from printing_plates.constants import is_plate_ink_spec

    for candidate in candidates:
        value = str(candidate or '').strip()
        if value and not is_plate_ink_spec(value):
            return value
    return ''


def build_sku_recipe_initial_from_recipe(recipe=None, *, planning_job=None, po_defaults=None):
    """Full SKU form initial: recipe first, then planning jobs for the SKU, then PO defaults.

    Layout data often lives on planning jobs while the draft recipe row is still sparse.
    Fall back across all jobs for the SKU so master entry shows complete known data.
    """
    from planning.models import PlanningJob
    from core.print_colors import normalize_print_color_for_form

    po_defaults = po_defaults or {}
    jobs = []
    if planning_job is not None:
        jobs.append(planning_job)
    sku_hint = ''
    if recipe is not None and (recipe.sku or '').strip():
        sku_hint = recipe.sku.strip()
    elif planning_job is not None and (planning_job.sku or '').strip():
        sku_hint = planning_job.sku.strip()
    if sku_hint:
        for job in PlanningJob.objects.filter(sku__iexact=sku_hint).order_by('-updated_at', '-id'):
            if planning_job is not None and job.pk == planning_job.pk:
                continue
            jobs.append(job)

    def _r(name, default=''):
        if recipe is not None:
            value = getattr(recipe, name, None)
            if not _field_is_blank(value):
                return value
        for job in jobs:
            if hasattr(job, name):
                value = getattr(job, name, None)
                if not _field_is_blank(value):
                    return value
        if name in po_defaults and not _field_is_blank(po_defaults.get(name)):
            return po_defaults.get(name)
        return default

    sku = _r('sku', '') or sku_hint
    job_name = _r('job_name', '')
    if not job_name:
        job_name = sku

    color_candidates = []
    if recipe is not None:
        color_candidates.append(getattr(recipe, 'color_spec', None))
    for job in jobs:
        color_candidates.append(getattr(job, 'color_spec', None))
    color_candidates.append(po_defaults.get('color_spec'))
    color_spec = normalize_print_color_for_form(_safe_color_spec(*color_candidates))

    return {
        'sku': sku,
        'job_name': job_name,
        'job_process_type': _r('job_process_type', 'print_and_pack') or 'print_and_pack',
        'material': _r('material', ''),
        'color_spec': color_spec,
        'application': _r('application', ''),
        'product_type': _r('product_type', ''),
        'machine_name': _r('machine_name', ''),
        'plate_set_no': _r('plate_set_no', ''),
        'size_w_mm': _r('size_w_mm', None),
        'size_h_mm': _r('size_h_mm', None),
        'ups': _r('ups', None),
        'print_sheet_size': _r('print_sheet_size', ''),
        'purchase_sheet_size': _r('purchase_sheet_size', ''),
        'purchase_sheet_ups': _r('purchase_sheet_ups', None),
        'default_unit_cost': _r('default_unit_cost', None),
        'daily_demand': _r('daily_demand', None),
        'awc_no': _r('awc_no', ''),
        'die_cutting': _r('die_cutting', ''),
        'notes': _r('notes', ''),
        'remarks': _r('remarks', ''),
    }


def build_sku_recipe_initial_from_planning_job(planning_job, *, recipe=None, po_defaults=None):
    """Build form initial values, preferring recipe then planning job then PO defaults."""
    return build_sku_recipe_initial_from_recipe(
        recipe,
        planning_job=planning_job,
        po_defaults=po_defaults,
    )


def hydrate_sku_recipe_from_planning_jobs(recipe, *, planning_job=None):
    """
    Copy blank recipe fields from planning jobs for the same SKU.

    Designer layout often exists only on jobs. After reopen/send-back the recipe
    row can look empty even though jobs still hold the data — fill blanks only.
    """
    if not recipe:
        return False

    initial = build_sku_recipe_initial_from_recipe(recipe, planning_job=planning_job)
    update_fields = []
    for field_name in (
        SKU_RECIPE_DESIGNER_FIELDS
        + SKU_RECIPE_PLANNER_FIELDS
        + SKU_RECIPE_SHARED_FIELDS
    ):
        if field_name in {'sku', 'notes', 'remarks'}:
            continue
        current = getattr(recipe, field_name, None)
        if not _field_is_blank(current):
            continue
        value = initial.get(field_name)
        if _field_is_blank(value):
            continue
        setattr(recipe, field_name, value)
        update_fields.append(field_name)

    if not update_fields:
        return False
    update_fields.append('updated_at')
    recipe.save(update_fields=list(dict.fromkeys(update_fields)))
    return True


# Fields that affect plate size / layout — warn about plate remake when these change.
PLATE_REMAKE_IMPACT_FIELDS = {
    'machine_name': 'Machine',
    'print_sheet_size': 'Print Sheet Size',
    'size_w_mm': 'Size Width',
    'size_h_mm': 'Size Height',
    'ups': 'UPS',
}


def sku_has_jobs_with_plates(sku):
    """True when any job for this SKU has plates sent, received, issued, or archived."""
    sku_value = (sku or '').strip()
    if not sku_value:
        return False

    from printing_plates.models import PlateRequest

    plate_statuses = {
        PlateRequest.STATUS_SENT,
        PlateRequest.STATUS_RECEIVED,
        PlateRequest.STATUS_AVAILABLE,
        PlateRequest.STATUS_ARCHIVED,
    }
    if PlateRequest.objects.filter(
        Q(planning_job__sku__iexact=sku_value) | Q(job_card__SKU__iexact=sku_value) | Q(sku_recipe__sku__iexact=sku_value),
        status__in=plate_statuses,
    ).exists():
        return True

    return PlanningJob.objects.filter(
        sku__iexact=sku_value,
        planning_stage__in={'plate_received', 'repeat_plate_making', 'new_plate_making', 'planning_done'},
    ).exists()


def get_plate_remake_impact_changes(old_values, new_values):
    """Return labels of plate-impact fields that changed between old and new value maps."""
    changed = []
    for field_name, label in PLATE_REMAKE_IMPACT_FIELDS.items():
        old_value = old_values.get(field_name)
        new_value = new_values.get(field_name)
        if not _master_sync_field_values_equal(old_value, new_value, field_name=field_name):
            changed.append(label)
    return changed


def build_plate_remake_warning(changed_labels, *, context='save'):
    if not changed_labels:
        return ''
    labels = ', '.join(changed_labels)
    if context == 'sync':
        return (
            f'Plate remake may be required ({labels} changed). '
            'Machine change needs plates sized for that press even if print sheet size stays the same. '
            'If plates already exist, use Production → Released Jobs → Request plates.'
        )
    return (
        f'Plate remake may be required ({labels} changed). '
        'Machine change needs plates sized for that press even if print sheet size stays the same. '
        'After re-approval, sync the job and use Production → Released Jobs → Request plates if plates already exist.'
    )


def get_plate_remake_warning_for_recipe_save(recipe, previous_values):
    if not recipe or not sku_has_jobs_with_plates(recipe.sku):
        return ''
    new_values = {field: getattr(recipe, field, None) for field in PLATE_REMAKE_IMPACT_FIELDS}
    changed = get_plate_remake_impact_changes(previous_values or {}, new_values)
    return build_plate_remake_warning(changed, context='save')


def job_has_plates(job):
    from printing_plates.models import PlateRequest

    plate_statuses = {
        PlateRequest.STATUS_SENT,
        PlateRequest.STATUS_RECEIVED,
        PlateRequest.STATUS_AVAILABLE,
        PlateRequest.STATUS_ARCHIVED,
    }
    if PlateRequest.objects.filter(
        Q(planning_job=job) | Q(job_card__planning_job=job),
        status__in=plate_statuses,
    ).exists():
        return True
    return (job.planning_stage or '') in {
        'plate_received',
        'repeat_plate_making',
        'new_plate_making',
        'planning_done',
    }


def get_plate_remake_warning_for_job_sync(job, diffs):
    if not diffs or not job_has_plates(job):
        return ''
    changed = [
        PLATE_REMAKE_IMPACT_FIELDS[field_name]
        for field_name in PLATE_REMAKE_IMPACT_FIELDS
        if field_name in diffs
    ]
    if not changed:
        return ''
    return build_plate_remake_warning(changed, context='sync')


SKU_RECIPE_STATUS_ORDER = {'approved': 0, 'reviewed': 1, 'pending_review': 2, 'draft': 3}

PLANNING_TO_SKU_RECIPE_FIELDS = (
    'job_name',
    'material',
    'application',
    'machine_name',
    'color_spec',
    'size_w_mm',
    'size_h_mm',
    'ups',
    'print_sheet_size',
    'purchase_sheet_size',
    'purchase_sheet_ups',
    'plate_set_no',
    'remarks',
)


def get_best_sku_recipe_for_sku(sku):
    sku_value = (sku or '').strip()
    if not sku_value:
        return None
    candidates = list(SkuRecipe.objects.filter(sku__iexact=sku_value))
    if not candidates:
        return None
    return min(
        candidates,
        key=lambda recipe: SKU_RECIPE_STATUS_ORDER.get(recipe.master_data_status or '', 99),
    )


def ensure_sku_recipe_for_planning_job(planning_job, *, actor=None, create_if_missing=True):
    """Return the SKU master row for a planning job, creating a draft when missing."""
    sku = (getattr(planning_job, 'sku', None) or '').strip()
    if not sku:
        return None

    recipe = get_best_sku_recipe_for_sku(sku)
    if recipe or not create_if_missing:
        return recipe

    return SkuRecipe.objects.create(
        sku=sku,
        job_name=(planning_job.job_name or '').strip() or sku,
        master_data_status='draft',
        created_by=actor,
    )


def sync_planning_job_fields_to_sku_recipe(planning_job, recipe, *, submit_for_review=False):
    """Copy planning/designer fields onto the SKU master when recipe values are blank."""
    if not planning_job or not recipe:
        return False

    from core.print_colors import apply_print_color_to_planning_job, apply_print_color_to_sku_recipe
    from printing_plates.constants import is_plate_ink_spec

    # Clear ink-chip pollution from planning print-color before syncing.
    if is_plate_ink_spec(planning_job.color_spec):
        planning_job.color_spec = ''
        planning_job.save(update_fields=['color_spec', 'updated_at'])
    apply_print_color_to_planning_job(planning_job)

    update_fields = []
    for field_name in PLANNING_TO_SKU_RECIPE_FIELDS:
        source_value = getattr(planning_job, field_name, None)
        if source_value is None or (isinstance(source_value, str) and not source_value.strip()):
            continue
        if field_name == 'color_spec' and is_plate_ink_spec(source_value):
            continue

        current_value = getattr(recipe, field_name, None)
        if isinstance(current_value, str):
            if current_value.strip():
                continue
        elif current_value is not None:
            continue

        setattr(recipe, field_name, source_value)
        update_fields.append(field_name)

    if is_plate_ink_spec(recipe.color_spec):
        recipe.color_spec = ''
        update_fields.append('color_spec')

    if apply_print_color_to_sku_recipe(recipe) and 'color_spec' not in update_fields:
        update_fields.append('color_spec')

    if submit_for_review and recipe.master_data_status == 'draft' and update_fields:
        missing = _missing_required_master_fields(recipe, recipe.job_name or planning_job.job_name)
        if not missing:
            recipe.master_data_status = 'pending_review'
            update_fields.append('master_data_status')

    if update_fields:
        update_fields.append('updated_at')
        recipe.save(update_fields=list(dict.fromkeys(update_fields)))
        return True
    return False


def apply_designer_layout_to_sku_recipe(planning_job, recipe, posted_values, *, submit_for_review=True):
    """Persist designer layout specs from plate making onto the SKU master.

    Values like die cutting \"NO\" / \"No\" are valid and must be saved (never treated as blank).
    Plate ink chips must never be written into color_spec.
    """
    if not planning_job or not recipe:
        return False

    from printing_plates.constants import is_plate_ink_spec

    simple_fields = {
        'size_w_mm': 'size_w_mm',
        'size_h_mm': 'size_h_mm',
        'ups': 'ups',
        'print_sheet_size': 'print_sheet_size',
        'purchase_sheet_size': 'purchase_sheet_size',
        'purchase_sheet_ups': 'purchase_sheet_ups',
        'die_cutting': 'die_cutting',
        # plate_color chips are ink names only — never overwrite production print color (color_spec).
        'awc_no': 'awc_no',
    }

    posted_awc = normalize_awc_no(posted_values.get('awc_no'))
    if posted_awc:
        conflict = get_awc_conflict_message(
            posted_awc,
            sku=recipe.sku or getattr(planning_job, 'sku', ''),
            exclude_recipe_id=recipe.pk if recipe.pk else None,
        )
        if conflict:
            raise ValueError(conflict)

    update_fields = []
    for recipe_field, posted_field in simple_fields.items():
        if posted_field not in posted_values:
            continue
        raw_value = posted_values.get(posted_field)
        if raw_value is None:
            continue
        if recipe_field == 'awc_no':
            value = normalize_awc_no(raw_value)
        elif recipe_field == 'die_cutting':
            value = normalize_die_cutting(raw_value)
        else:
            value = str(raw_value).strip()
        # Only skip truly empty strings. \"NO\" / \"None\" / \"N/A\" are valid die-cutting answers.
        if value == '':
            continue

        if recipe_field in {'size_w_mm', 'size_h_mm', 'ups', 'purchase_sheet_ups'}:
            try:
                value = int(round(float(value)))
            except (TypeError, ValueError):
                continue

        setattr(recipe, recipe_field, value)
        update_fields.append(recipe_field)

    set_no = str(posted_values.get('set_no') or '').strip()
    new_set_no = str(posted_values.get('new_set_no') or '').strip()
    plate_set_no = set_no or new_set_no
    if plate_set_no:
        recipe.plate_set_no = plate_set_no
        update_fields.append('plate_set_no')

    # Never keep plate-ink text in production print-color field.
    if is_plate_ink_spec(recipe.color_spec):
        recipe.color_spec = ''
        update_fields.append('color_spec')

    # Only send to QC when all approval-required fields are present.
    # Incomplete recipes stay Draft so planner can finish material/application/etc.
    if submit_for_review and recipe.master_data_status in {'draft', ''}:
        missing = _missing_required_master_fields(recipe, recipe.job_name or planning_job.job_name)
        if not missing:
            recipe.master_data_status = 'pending_review'
            update_fields.append('master_data_status')

    if update_fields:
        update_fields.append('updated_at')
        recipe.save(update_fields=list(dict.fromkeys(update_fields)))
        return True
    return False


def designer_layout_missing_fields(posted_values):
    """Return labels of required designer layout fields that are blank."""
    required = [
        ('size_w_mm', 'Size Width (mm)'),
        ('size_h_mm', 'Size Height (mm)'),
        ('print_sheet_size', 'Print Sheet Size'),
        ('ups', 'Ups'),
        ('purchase_sheet_size', 'Purchase Sheet Size'),
        ('purchase_sheet_ups', 'Purchase Sheet Ups'),
        ('die_cutting', 'Die Cutting'),
        ('awc_no', 'AWC #'),
    ]
    missing = []
    for field_name, label in required:
        if str(posted_values.get(field_name) or '').strip() == '':
            missing.append(label)
    set_no = str(posted_values.get('set_no') or '').strip()
    new_set_no = str(posted_values.get('new_set_no') or '').strip()
    if not set_no and not new_set_no:
        missing.append('Set No / New Set No')
    return missing


def resolve_designer_layout_values(posted_values, *, recipe=None, planning_job=None, plate_request=None):
    """Merge posted designer values with existing recipe/job/request (POST wins)."""

    def _pick(*candidates):
        for candidate in candidates:
            if candidate is None:
                continue
            text = str(candidate).strip()
            if text != '':
                return text
        return ''

    recipe = recipe
    job = planning_job
    req = plate_request
    return {
        'size_w_mm': _pick(
            posted_values.get('size_w_mm'),
            getattr(recipe, 'size_w_mm', None),
            getattr(job, 'size_w_mm', None),
        ),
        'size_h_mm': _pick(
            posted_values.get('size_h_mm'),
            getattr(recipe, 'size_h_mm', None),
            getattr(job, 'size_h_mm', None),
        ),
        'print_sheet_size': _pick(
            posted_values.get('print_sheet_size'),
            getattr(recipe, 'print_sheet_size', None),
            getattr(job, 'print_sheet_size', None),
        ),
        'ups': _pick(
            posted_values.get('ups'),
            getattr(recipe, 'ups', None),
            getattr(job, 'ups', None),
        ),
        'purchase_sheet_size': _pick(
            posted_values.get('purchase_sheet_size'),
            getattr(recipe, 'purchase_sheet_size', None),
            getattr(job, 'purchase_sheet_size', None),
        ),
        'purchase_sheet_ups': _pick(
            posted_values.get('purchase_sheet_ups'),
            getattr(recipe, 'purchase_sheet_ups', None),
            getattr(job, 'purchase_sheet_ups', None),
        ),
        'die_cutting': normalize_die_cutting(_pick(
            posted_values.get('die_cutting'),
            getattr(req, 'die_cutting', None),
            getattr(recipe, 'die_cutting', None),
        )),
        'awc_no': normalize_awc_no(_pick(
            posted_values.get('awc_no'),
            getattr(req, 'awc_no', None),
            getattr(recipe, 'awc_no', None),
        )),
        'set_no': _pick(
            posted_values.get('set_no'),
            getattr(req, 'set_no', None),
            getattr(recipe, 'plate_set_no', None),
            getattr(job, 'plate_set_no', None),
        ),
        'new_set_no': _pick(
            posted_values.get('new_set_no'),
            getattr(req, 'new_set_no', None),
        ),
    }


def prepare_sku_recipe_form_for_master_entry(form, *, action=''):
    """Early plate-making saves only require planner fields; layout specs can follow later."""
    if action not in {'send_to_plate_making', 'save_draft'}:
        return

    for field_name in SKU_RECIPE_DESIGNER_FIELDS:
        if field_name in form.fields:
            form.fields[field_name].required = False
    if action == 'send_to_plate_making' and 'product_type' in form.fields:
        form.fields['product_type'].required = False


def trigger_plate_request_for_planning_job(planning_job, user):
    """Create or return an active plate request when planning stage requires it."""
    from printing_plates.services import create_or_get_plate_request_from_planning_job

    return create_or_get_plate_request_from_planning_job(planning_job, user)


def get_plate_request_block_for_master_entry(planning_job):
    """
    If planner must not create another plate request from master entry, return
    (plate_request, reason_code) where reason_code is 'open' or 'issued'.
    """
    from printing_plates.services import (
        get_issued_plate_request_for_planning_job,
        get_open_plate_request_for_planning_job,
    )

    open_req = get_open_plate_request_for_planning_job(planning_job)
    if open_req:
        return open_req, 'open'
    issued_req = get_issued_plate_request_for_planning_job(planning_job)
    if issued_req:
        return issued_req, 'issued'
    return None, ''


def _planning_status_filter_values(status):
    normalized_status = _normalize_status(status, default='')
    if not normalized_status:
        return []
    return sorted(PLANNING_STATUS_ALIASES.get(normalized_status, {normalized_status}))



def _parse_date_filter(raw_value):
    if not raw_value:
        return None
    try:
        return datetime.strptime(str(raw_value).strip(), '%Y-%m-%d').date()
    except (ValueError, TypeError):
        return None


def _normalize_purchase_material_origin(raw_value):
    value = (raw_value or '').strip().lower()
    if value in {'local'}:
        return 'local'
    if value in {'import', 'imported'}:
        return 'import'
    return ''


def _normalize_po_number(raw_value):
    value = (raw_value or '').strip().upper()
    if not value:
        return ''
    if value.startswith('PO'):
        match = re.search(r'(\d+)$', value)
        if match:
            return match.group(1)
    return re.sub(r'[^A-Z0-9]+', '', value)


def _build_job_card_pdf_bytes(job, scan_url):
    if not REPORTLAB_AVAILABLE:
        raise RuntimeError('reportlab is required to generate PDF job cards. Install reportlab and restart the server.')
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        leftMargin=12 * mm,
        rightMargin=12 * mm,
        topMargin=12 * mm,
        bottomMargin=12 * mm,
    )

    styles = getSampleStyleSheet()
    normal = styles['Normal']
    normal.fontName = 'Helvetica'
    normal.fontSize = 9
    normal.leading = 11

    title_style = ParagraphStyle('Title', parent=normal, fontName='Helvetica-Bold', fontSize=18, leading=20)
    subtitle_style = ParagraphStyle('Subtitle', parent=normal, fontName='Helvetica-Bold', fontSize=11, leading=13)
    section_title_style = ParagraphStyle('SectionTitle', parent=normal, fontName='Helvetica-Bold', fontSize=10.5, leading=12)
    label_style = ParagraphStyle('Label', parent=normal, fontName='Helvetica-Bold', fontSize=8.5, leading=10)

    story = [Paragraph('UTOPIA PRINTING & PACKAGING', title_style), Spacer(1, 4), Paragraph('PRODUCTION JOB CARD', subtitle_style), Spacer(1, 8)]

    header_data = [
        [Paragraph('JOB CARD #', label_style), _format_job_value(job.jc_number), Paragraph('PO #', label_style), _format_job_value(job.po_number)],
        [Paragraph('DATE', label_style), _format_job_value(job.plan_date), Paragraph('STATUS', label_style), _format_job_value(job.workflow_status_label)],
        [Paragraph('SKU', label_style), _format_job_value(job.sku), Paragraph('JOB NAME', label_style), _format_job_value(job.job_name)],
        [Paragraph('REPEAT FLAG', label_style), _format_job_value(job.repeat_flag), Paragraph('DEPARTMENT', label_style), _format_job_value(job.department)],
    ]
    header_table = Table(header_data, colWidths=[32 * mm, 65 * mm, 32 * mm, 65 * mm], hAlign='LEFT')
    header_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('GRID', (0, 0), (-1, -1), 0.35, colors.black),
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#dedede')),
        ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
    ]))
    story.extend([header_table, Spacer(1, 10)])

    material_data = [
        [Paragraph('ORDER QTY', label_style), _format_job_value(job.order_qty), Paragraph('PRINT PCS', label_style), _format_job_value(job.print_pcs)],
        [Paragraph('MATERIAL TYPE', label_style), _format_job_value(job.material), Paragraph('PRINT COLOR', label_style), _format_job_value(job.color_spec)],
        [Paragraph('APPLICATION', label_style), _format_job_value(job.application), Paragraph('PRINT SHEET SIZE', label_style), _format_job_value(job.print_sheet_size)],
        [Paragraph('UPS', label_style), _format_job_value(job.ups), Paragraph('PRINT SHEETS', label_style), _format_job_value(job.print_sheets)],
        [Paragraph('ACTUAL SHEETS', label_style), _format_job_value(job.calculated_sheets_required), Paragraph('WASTAGE', label_style), _format_job_value(job.wastage_sheets)],
        [Paragraph('PURCHASE ORIGIN', label_style), _format_job_value(job.purchase_material_origin), Paragraph('PURCHASE SHEET SIZE', label_style), _format_job_value(job.purchase_sheet_size)],
        [Paragraph('PURCHASE SHEET UPS', label_style), _format_job_value(job.purchase_sheet_ups), Paragraph('PURCHASE REQ', label_style), _format_job_value(job.purchase_sheet_required)],
        [Paragraph('MACHINE', label_style), _format_job_value(job.machine_name), Paragraph('TOTAL COLORS', label_style), _format_job_value(job.number_of_colors)],
        [Paragraph('PLATE SET NO.', label_style), _format_job_value(job.plate_set_no), Paragraph('AWC NO.', label_style), _format_job_value(job.awc_no_display)],
        [Paragraph('AGING DAYS', label_style), _format_job_value(job.aging_days), Paragraph('DIE CUTTING', label_style), _format_job_value(job.die_cutting_display)],
    ]
    material_table = Table(material_data, colWidths=[32 * mm, 65 * mm, 32 * mm, 65 * mm], hAlign='LEFT')
    material_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.grey),
        ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#eeeeee')),
    ]))
    story.extend([Paragraph('MATERIAL AND WORK PROCESS', section_title_style), Spacer(1, 4), material_table, Spacer(1, 10)])

    recipe = job.sku_recipe
    application_data = [
        [Paragraph('LAMINATION', label_style), _format_job_value(job.application), Paragraph('DIE CUTTING', label_style), _format_job_value(job.die_cutting_display)],
        [Paragraph('ART WORK NO.', label_style), '-', Paragraph('P SET NO.', label_style), _format_job_value(job.plate_set_no)],
        [Paragraph('SPECIAL INSTRUCTIONS', label_style), _paragraph_text(job.requirement or '-'), '', ''],
    ]
    application_table = Table(application_data, colWidths=[30 * mm, 67 * mm, 30 * mm, 65 * mm], hAlign='LEFT')
    application_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('INNERGRID', (0, 0), (-1, 1), 0.25, colors.grey),
        ('BOX', (0, 0), (-1, 1), 0.5, colors.black),
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#f3f3f3')),
    ]))
    story.extend([application_table, Spacer(1, 12)])

    signature_data = [
        [Paragraph('Prepared by', label_style), '', Paragraph('Checked By', label_style), '', Paragraph('Plate Check By', label_style), '', Paragraph('Approved By', label_style), ''],
    ]
    signature_table = Table(signature_data, colWidths=[28 * mm, 34 * mm, 28 * mm, 34 * mm, 28 * mm, 34 * mm, 28 * mm, 34 * mm], hAlign='LEFT')
    signature_table.setStyle(TableStyle([
        ('LINEABOVE', (1, 0), (1, 0), 0.25, colors.black),
        ('LINEABOVE', (3, 0), (3, 0), 0.25, colors.black),
        ('LINEABOVE', (5, 0), (5, 0), 0.25, colors.black),
        ('LINEABOVE', (7, 0), (7, 0), 0.25, colors.black),
    ]))
    story.extend([signature_table, Spacer(1, 10)])

    material_issue_data = [[Paragraph('MATERIAL ISSUANCE', section_title_style), '', '', '', '', '']]
    material_issue_data.append([Paragraph('Date', label_style), Paragraph('Machine', label_style), Paragraph('Operator', label_style), Paragraph('Shift A/B', label_style), Paragraph('Sheet Size', label_style), Paragraph('Full Sheet Qty', label_style)])
    for _ in range(3):
        material_issue_data.append(['-', '-', '-', '-', '-', '-'])
    material_issue_table = Table(material_issue_data, colWidths=[24 * mm, 30 * mm, 35 * mm, 28 * mm, 35 * mm, 30 * mm], hAlign='LEFT')
    material_issue_table.setStyle(TableStyle([
        ('GRID', (0, 1), (-1, -1), 0.25, colors.grey),
        ('BACKGROUND', (0, 1), (-1, 1), colors.HexColor('#d9d9d9')),
        ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
    ]))

    printing_data = [[Paragraph('PRINTING', section_title_style), '', '', '', '', '', '']]
    printing_data.append([Paragraph('Date', label_style), Paragraph('Machine', label_style), Paragraph('Operator', label_style), Paragraph('Shift A/B', label_style), Paragraph('Print Sheet Qty', label_style), Paragraph('Wastage Sheet', label_style), Paragraph('Half Good', label_style)])
    for _ in range(4):
        printing_data.append(['-', '-', '-', '-', '-', '-', '-'])
    printing_table = Table(printing_data, colWidths=[24 * mm, 30 * mm, 30 * mm, 28 * mm, 34 * mm, 34 * mm, 26 * mm], hAlign='LEFT')
    printing_table.setStyle(TableStyle([
        ('GRID', (0, 1), (-1, -1), 0.25, colors.grey),
        ('BACKGROUND', (0, 1), (-1, 1), colors.HexColor('#d9d9d9')),
        ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
    ]))

    story.extend([material_issue_table, Spacer(1, 10), printing_table, Spacer(1, 12)])

    dispatch_data = [[Paragraph('DISPATCH', section_title_style), '', '', '', '', '']]
    dispatch_data.append([Paragraph('Delivery Date', label_style), Paragraph('DC #', label_style), Paragraph('Qty', label_style), Paragraph('Packing', label_style), Paragraph('Delivered To', label_style), ''])
    for _ in range(6):
        dispatch_data.append(['-', '-', '-', '-', '-', '-'])
    dispatch_table = Table(dispatch_data, colWidths=[30 * mm, 24 * mm, 24 * mm, 30 * mm, 35 * mm, 40 * mm], hAlign='LEFT')
    dispatch_table.setStyle(TableStyle([
        ('GRID', (0, 1), (-1, -1), 0.25, colors.grey),
        ('BACKGROUND', (0, 1), (-1, 1), colors.HexColor('#d9d9d9')),
        ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
    ]))
    story.extend([dispatch_table, Spacer(1, 10)])

    cutting_data = [
        [Paragraph('CUTTING SLIP', section_title_style), '', '', '', '', ''],
        [Paragraph('Job Card #', label_style), _format_job_value(job.jc_number), Paragraph('Job Name', label_style), _format_job_value(job.job_name), Paragraph('Purch sheet size', label_style), _format_job_value(job.purchase_sheet_size)],
        [Paragraph('Purch sheet Ups', label_style), _format_job_value(job.purchase_sheet_ups), Paragraph('Print sheet size', label_style), _format_job_value(job.print_sheet_size), Paragraph('Type', label_style), _format_job_value(job.material)],
        [Paragraph('Purch sheet Qty', label_style), _format_job_value(job.purchase_sheet_required), Paragraph('Remarks', label_style), _paragraph_text(job.remarks or (recipe.notes if recipe else '') or job.requirement or '-'), '', ''],
    ]
    cutting_table = Table(cutting_data, colWidths=[30 * mm, 35 * mm, 30 * mm, 35 * mm, 30 * mm, 35 * mm], hAlign='LEFT')
    cutting_table.setStyle(TableStyle([
        ('SPAN', (0, 0), (-1, 0)),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('GRID', (0, 1), (-1, -1), 0.25, colors.grey),
        ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#f2f2f2')),
    ]))
    story.extend([cutting_table])

    doc.build(story)
    return buffer.getvalue()



def _sku_key(sku):
    return (sku or '').strip().upper()



def _missing_required_master_fields(recipe, fallback_job_name=''):
    missing = []
    cut_and_pack = bool(
        recipe and (getattr(recipe, 'job_process_type', '') or 'print_and_pack') == 'cut_and_pack'
    )
    skip_for_cut_and_pack = {'color_spec'}

    if not recipe:
        fallback = (fallback_job_name or '').strip()
        return [
            label
            for field, label in SKU_MASTER_APPROVAL_REQUIRED_FIELDS
            if not (field == 'job_name' and fallback)
        ]

    for field, label in SKU_MASTER_APPROVAL_REQUIRED_FIELDS:
        if cut_and_pack and field in skip_for_cut_and_pack:
            continue
        value = getattr(recipe, field, None)
        if isinstance(value, str):
            if not value.strip():
                missing.append(label)
        elif value is None:
            missing.append(label)
    return missing



def sync_recipe_operational_fields_to_job(job, recipe=None):
    """Copy machine/plate set from approved SKU master onto the job when job fields are blank."""
    recipe = recipe or job.approved_sku_recipe
    if not recipe:
        return False

    update_fields = []
    if not str(job.machine_name or '').strip() and str(recipe.machine_name or '').strip():
        job.machine_name = str(recipe.machine_name).strip()
        update_fields.append('machine_name')
    if not str(job.plate_set_no or '').strip() and str(recipe.plate_set_no or '').strip():
        job.plate_set_no = str(recipe.plate_set_no).strip()
        update_fields.append('plate_set_no')
    recipe_process = (getattr(recipe, 'job_process_type', None) or '').strip()
    if recipe_process and recipe_process != (job.job_process_type or 'print_and_pack'):
        job.job_process_type = recipe_process
        update_fields.append('job_process_type')

    if update_fields:
        update_fields.append('updated_at')
        job.save(update_fields=update_fields)
        return True
    return False


def get_job_qc_submission_blockers(job, *, apply_recipe_sync=True):
    """Return human-readable blockers before a draft job can move to pending_qc."""
    blockers = []

    approved_recipe = job.approved_sku_recipe
    active_recipe = job.sku_recipe
    if not active_recipe:
        blockers.append(f'SKU recipe for {job.sku or "this job"} is missing.')
        return blockers
    if not approved_recipe:
        blockers.append(
            f'SKU recipe for {job.sku or "this job"} exists but is not approved; QC submission is blocked until approval.'
        )
        return blockers

    missing_master = _missing_required_master_fields(approved_recipe, job.job_name)
    if missing_master:
        blockers.append(
            'Approved SKU master is incomplete: '
            f'{", ".join(missing_master)}. Reopen SKU, update the missing fields, and re-approve.'
        )

    if apply_recipe_sync:
        sync_recipe_operational_fields_to_job(job, approved_recipe)

    for field_name, error_message in job.pre_submit_qc_validation_errors().items():
        if error_message in blockers:
            continue
        if field_name in {'machine_name', 'plate_set_no'}:
            blockers.append(
                f'{error_message} Update these on the locked SKU master (Reopen SKU), then re-approve.'
            )
        else:
            blockers.append(error_message)

    return blockers


def preview_job_qc_submission_blockers(job):
    return get_job_qc_submission_blockers(job, apply_recipe_sync=False)


def _sync_new_sku_requirement(existing_requirement, is_new):
    """Ensure NEW SKU requirement note exists only for New jobs."""
    lines = [line.strip() for line in str(existing_requirement or '').splitlines() if line.strip()]
    filtered_lines = [line for line in lines if line != NEW_SKU_REQUIREMENT_NOTE]

    if is_new:
        return '\n'.join([NEW_SKU_REQUIREMENT_NOTE] + filtered_lines)
    return '\n'.join(filtered_lines)



def _build_recipe_map(items):
    """Return a map of SKU-upper -> SkuRecipe for any existing recipe (any status).

    Priority: approved > reviewed > pending_review > draft.
    This ensures that repeat POs are recognised even when the bulk-uploaded
    recipe has not yet been formally approved in the ERP.
    """
    sku_values = sorted({_sku_key(item.get('sku')) for item in items if item.get('sku')})
    if not sku_values:
        return {}

    STATUS_PRIORITY = {'approved': 0, 'reviewed': 1, 'pending_review': 2, 'draft': 3}

    recipes = (
        SkuRecipe.objects
        .annotate(sku_upper=Upper('sku'))
        .filter(sku_upper__in=sku_values)
        .order_by('sku_upper')
    )

    result = {}
    for recipe in recipes:
        key = recipe.sku.upper()
        if key not in result:
            result[key] = recipe
        else:
            # Keep the higher-priority (more approved) record
            existing_priority = STATUS_PRIORITY.get(result[key].master_data_status, 99)
            incoming_priority = STATUS_PRIORITY.get(recipe.master_data_status, 99)
            if incoming_priority < existing_priority:
                result[key] = recipe
    return result



def _to_optional_positive_int(raw_value):
    value = _to_int(raw_value)
    if value is None:
        return None
    return value if value >= 0 else None



def _to_optional_decimal(raw_value):
    value = _to_decimal(raw_value)
    if value is None:
        return None
    return value if value >= 0 else None



def _sanitize_po_payload_items(payload):
    """Normalize payload items for workflow screens.

    Applies SKU-level deduplication and respects expected line count when available
    to avoid noisy extra rows from fallback parsers.
    """
    items, _ = _deduplicate_po_items_by_sku((payload or {}).get('items', []))

    # Merge OCR-near-duplicate SKUs when qty/date match and text is almost identical.
    consolidated = []
    for item in items:
        sku = (item.get('sku') or '').strip()
        qty = _to_int(item.get('quantity'))
        ddate = (item.get('delivery_date') or '').strip()
        sku_norm = ''.join(ch for ch in sku.upper() if ch.isalnum())
        merged = False
        for existing in consolidated:
            ex_sku = (existing.get('sku') or '').strip()
            ex_qty = _to_int(existing.get('quantity'))
            ex_ddate = (existing.get('delivery_date') or '').strip()
            ex_norm = ''.join(ch for ch in ex_sku.upper() if ch.isalnum())
            similar = SequenceMatcher(a=sku_norm, b=ex_norm).ratio() >= 0.985
            if similar and qty == ex_qty and ddate == ex_ddate:
                merged = True
                break
        if not merged:
            consolidated.append(item)
    items = consolidated

    expected_line_count = _to_int((payload or {}).get('expected_line_count'))
    if expected_line_count and expected_line_count > 0 and len(items) > expected_line_count:
        items = items[:expected_line_count]
    return items



def _po_payload_items(payload, exclude_ignored=True):
    items = _sanitize_po_payload_items(payload)
    if not exclude_ignored:
        return items

    ignored_skus = {
        _sku_key(s)
        for s in (payload.get('new_skus_ignored') or [])
        if s
    }
    if not ignored_skus:
        return items

    return [
        item
        for item in items
        if _sku_key(item.get('sku')) not in ignored_skus
    ]



def _annotate_items_with_recipe(items, recipe_map):
    annotated = []
    repeat_count = 0
    new_count = 0
    missing_skus = []

    for item in items:
        sku = (item.get('sku') or '').strip()
        key = _sku_key(sku)
        has_recipe = bool(key and key in recipe_map)
        item_copy = dict(item)
        item_copy['is_repeat'] = has_recipe
        item_copy['recipe_status'] = 'Repeat' if has_recipe else 'New'
        annotated.append(item_copy)

        if has_recipe:
            repeat_count += 1
        else:
            new_count += 1
            if sku:
                missing_skus.append(sku)

    return annotated, repeat_count, new_count, sorted(set(missing_skus))



def _deduplicate_po_items_by_sku(items):
    """Ensure one row per SKU in a PO payload by merging duplicate SKU lines."""
    merged = {}
    order = []
    duplicate_skus = set()

    for item in items:
        item_copy = dict(item)
        sku = (item_copy.get('sku') or '').strip()
        sku_key = _sku_key(sku)
        if not sku_key:
            continue

        if sku_key not in merged:
            merged[sku_key] = item_copy
            order.append(sku_key)
            continue

        duplicate_skus.add(sku)
        existing = merged[sku_key]

        existing_qty = _to_int(existing.get('quantity'))
        current_qty = _to_int(item_copy.get('quantity'))
        if existing_qty is None:
            existing['quantity'] = current_qty
        elif current_qty is not None:
            existing['quantity'] = existing_qty + current_qty

        existing_net = _to_decimal(existing.get('net_total'))
        current_net = _to_decimal(item_copy.get('net_total'))
        if existing_net is None:
            existing['net_total'] = _format_decimal_string(current_net)
        elif current_net is not None:
            existing['net_total'] = _format_decimal_string(existing_net + current_net)

        existing_subtotal = _to_decimal(existing.get('subtotal'))
        current_subtotal = _to_decimal(item_copy.get('subtotal'))
        if existing_subtotal is None:
            existing['subtotal'] = _format_decimal_string(current_subtotal)
        elif current_subtotal is not None:
            existing['subtotal'] = _format_decimal_string(existing_subtotal + current_subtotal)

        for field in ['job_name', 'delivery_date', 'unit', 'unit_cost']:
            if not existing.get(field) and item_copy.get(field):
                existing[field] = item_copy.get(field)

    deduped = [merged[key] for key in order]
    for idx, item in enumerate(deduped, start=1):
        item['line_no'] = idx
    return deduped, sorted(duplicate_skus)



def _history_repeat_new_counts(items, recipe_map=None):
    """Classify Repeat/New from approved SKU recipes, not historical PlanningJobs."""
    if recipe_map is None:
        recipe_map = _build_recipe_map(items)

    repeat_count = 0
    new_count = 0
    for item in items:
        sku = item.get('sku')
        sku_key = _sku_key(sku)
        if not sku_key:
            continue
        if sku_key in recipe_map:
            repeat_count += 1
        else:
            new_count += 1

    return repeat_count, new_count



def _sync_repeat_jobs_from_po(po_doc, actor=None):
    """Create or update draft planning jobs for all PO lines from one PO document."""
    payload = po_doc.extracted_payload or {}
    items, _ = _deduplicate_po_items_by_sku(payload.get('items', []))
    po_number = (payload.get('po_number') or '').strip()
    pr_number = (payload.get('pr_number') or '').strip()
    po_date = _parse_iso_date(payload.get('po_date'))
    delivery_location = payload.get('delivery_location', '')
    department = payload.get('department', '')

    if not items:
        return {'created': 0, 'updated': 0, 'locked': 0, 'missing_recipe': 0}

    item_sku_keys = {_sku_key(item.get('sku')) for item in items if item.get('sku')}
    existing_any_jobs_skus = set()
    if item_sku_keys:
        sku_any_query = Q()
        for sku_key in item_sku_keys:
            sku_any_query |= Q(sku__iexact=sku_key)
        existing_any_jobs_skus = {
            _sku_key(sku)
            for sku in PlanningJob.objects.filter(sku_any_query).values_list('sku', flat=True)
            if sku
        }

    recipe_map = _build_recipe_map(items)
    existing_jobs_by_sku = {}
    if po_number and item_sku_keys:
        existing_jobs = PlanningJob.objects.filter(po_number=po_number).order_by('-updated_at', '-id')
        for job in existing_jobs:
            key = _sku_key(job.sku)
            if key in item_sku_keys and key not in existing_jobs_by_sku:
                existing_jobs_by_sku[key] = job

    created_count = 0
    updated_count = 0
    locked_count = 0
    missing_recipe_count = 0
    seen_skus_in_payload = set()

    for item in items:
        sku = (item.get('sku') or '').strip()
        sku_key = _sku_key(sku)
        if not sku_key:
            continue

        recipe = recipe_map.get(sku_key)
        if not recipe:
            missing_recipe_count += 1

        existing_job = existing_jobs_by_sku.get(sku_key)
        if existing_job and _normalize_status(existing_job.status) != 'draft':
            locked_count += 1
            continue

        delivery_date = _parse_iso_date(item.get('delivery_date'))
        plan_date = po_doc.created_at.date() if po_doc and getattr(po_doc, 'created_at', None) else (delivery_date or po_date)
        qty = item.get('quantity')
        order_qty = int(qty) if qty is not None else None
        unit_cost_val = item.get('unit_cost')
        unit_cost_dec = Decimal(str(unit_cost_val)) if unit_cost_val is not None else None
        jc_number = (
            (item.get('jc_number') or item.get('jc') or item.get('job_card_no') or item.get('jobcardno'))
            or (existing_job.jc_number if existing_job else None)
            or allocate_next_jc_number(plan_date)
        )
        is_first_production = bool(
            sku_key
            and sku_key not in existing_any_jobs_skus
            and sku_key not in seen_skus_in_payload
        )

        explicit_repeat_flag = (item.get('repeat_flag') or item.get('repeat') or '').strip()
        if explicit_repeat_flag.lower() in {'new', 'repeat'}:
            forward_as_new = explicit_repeat_flag.lower() == 'new'
        elif existing_job:
            existing_repeat_flag = (existing_job.repeat_flag or '').strip().lower()
            if existing_repeat_flag in {'new', 'repeat'}:
                forward_as_new = existing_repeat_flag == 'new'
            else:
                prior_jobs_exist = PlanningJob.objects.filter(sku__iexact=sku).exclude(id=existing_job.id).exists()
                forward_as_new = not prior_jobs_exist
        else:
            # Even if there is no prior planning job, a bulk-uploaded recipe
            # (any status) signals that this SKU has been produced before.
            any_recipe_exists = sku_key and SkuRecipe.objects.filter(sku__iexact=sku_key).exists()
            forward_as_new = is_first_production and not any_recipe_exists

        current_requirement = existing_job.requirement if existing_job else ''

        fallback_job_name = (item.get('job_name') or '').strip() or sku
        if item.get('job_name') and (item.get('job_name') or '').strip():
            job_name_value = item.get('job_name').strip()
        elif recipe and (recipe.job_name or '').strip():
            job_name_value = recipe.job_name
        elif existing_job and (existing_job.job_name or '').strip():
            job_name_value = existing_job.job_name
        else:
            job_name_value = fallback_job_name

        material_value = (item.get('material') or '').strip() or (recipe.material if recipe else (existing_job.material if existing_job else ''))
        # Job process is SKU-master owned; planning never overrides.
        job_process_type_value = (recipe.job_process_type if recipe else '') or 'print_and_pack'
        color_spec_value = (item.get('color_spec') or item.get('color') or '').strip() or (recipe.color_spec if recipe else (existing_job.color_spec if existing_job else ''))
        application_value = (item.get('application') or '').strip() or (recipe.application if recipe else (existing_job.application if existing_job else ''))
        size_w_mm_value = _to_decimal(item.get('size_w_mm') or '') or (recipe.size_w_mm if recipe else (existing_job.size_w_mm if existing_job else None))
        size_h_mm_value = _to_decimal(item.get('size_h_mm') or '') or (recipe.size_h_mm if recipe else (existing_job.size_h_mm if existing_job else None))
        ups_value = _to_decimal(item.get('ups') or item.get('no_of_ups')) or (recipe.ups if recipe else (existing_job.ups if existing_job else None))
        print_sheet_size_value = (item.get('print_sheet_size') or '').strip() or (recipe.print_sheet_size if recipe else (existing_job.print_sheet_size if existing_job else ''))
        purchase_sheet_size_value = (item.get('purchase_sheet_size') or '').strip() or (recipe.purchase_sheet_size if recipe else (existing_job.purchase_sheet_size if existing_job else ''))
        purchase_sheet_ups_value = _to_decimal(item.get('purchase_sheet_ups') or '') or (recipe.purchase_sheet_ups if recipe else (existing_job.purchase_sheet_ups if existing_job else None))
        daily_demand_value = _to_decimal(item.get('daily_demand') or '') or (recipe.daily_demand if recipe else (existing_job.daily_demand if existing_job else None))
        unit_cost_value = unit_cost_dec if unit_cost_dec is not None else (recipe.default_unit_cost if recipe else (existing_job.unit_cost if existing_job else None))
        actual_sheet_required_value = _to_int(item.get('actual_sheet_required') or item.get('actual_sheet_require') or item.get('sheet')) or (existing_job.actual_sheet_required if existing_job else None)
        wastage_sheets_value = _to_int(item.get('wastage') or item.get('wastage_sheets')) or (existing_job.wastage_sheets if existing_job else None)
        purchase_sheet_required_value = _to_int(item.get('purchase_sheet_required') or item.get('purchase_sheet_require')) or (existing_job.purchase_sheet_required if existing_job else None)
        pkt_value = _to_decimal(item.get('pkt') or item.get('pkt_value') or '') or (existing_job.pkt_value if existing_job else None)
        stock_qty_value = _to_decimal(item.get('stock_qty') or item.get('stock') or '') or (existing_job.stock_qty if existing_job else None)
        balance_qty_value = _to_int(item.get('balance_qty') or item.get('balance') or '') or (existing_job.balance_qty if existing_job else None)
        plate_set_no_value = (item.get('plate_set_no') or item.get('p_set_no') or '').strip() or (recipe.plate_set_no if recipe else (existing_job.plate_set_no if existing_job else ''))
        die_cutting_value = (item.get('die_cutting') or '').strip() or (recipe.die_cutting if recipe else (existing_job.die_cutting if hasattr(existing_job, 'die_cutting') else ''))
        purchase_material_origin_value = _normalize_purchase_material_origin(item.get('purchase_material_origin') or item.get('purchase_material') or '') or (existing_job.purchase_material_origin if existing_job else '')
        machine_name_value = (item.get('machine_name') or item.get('machine') or '').strip() or (recipe.machine_name if recipe else (existing_job.machine_name if existing_job else ''))
        status_value = _normalize_status(item.get('status') or '') or 'draft'
        requirement_value = (item.get('requirement') or '').strip() or current_requirement

        if existing_job and not requirement_value:
            requirement_value = existing_job.requirement or ''

        requirement_value = _sync_new_sku_requirement(requirement_value, forward_as_new)
        if recipe and not forward_as_new:
            requirement_value = _append_unique_note_line(
                requirement_value,
                _build_cost_mismatch_note(recipe.default_unit_cost, unit_cost_dec),
            )

        defaults = {
            'po_number': po_number,
            'pr_reference': pr_number,
            'sku': sku,
            'job_name': job_name_value,
            'order_qty': order_qty,
            'department': department,
            'destination': delivery_location,
            'delivery_date': delivery_date,
            'unit_cost': unit_cost_value,
            'status': status_value,
            'repeat_flag': 'New' if forward_as_new else 'Repeat',
            'requirement': requirement_value,
            'material': material_value,
            'job_process_type': job_process_type_value,
            'color_spec': color_spec_value,
            'application': application_value,
            'size_w_mm': size_w_mm_value,
            'size_h_mm': size_h_mm_value,
            'ups': ups_value,
            'print_sheet_size': print_sheet_size_value,
            'purchase_sheet_size': purchase_sheet_size_value,
            'purchase_sheet_ups': purchase_sheet_ups_value,
            'daily_demand': daily_demand_value,
            'plate_set_no': plate_set_no_value,
            'machine_name': machine_name_value,
            'actual_sheet_required': actual_sheet_required_value,
            'wastage_sheets': wastage_sheets_value,
            'purchase_sheet_required': purchase_sheet_required_value,
            'pkt_value': pkt_value,
            'remarks': (item.get('remarks') or '').strip() or (existing_job.remarks if existing_job else '') or (recipe.notes if recipe else ''),
            'purchase_material_origin': purchase_material_origin_value,
            'stock_qty': stock_qty_value,
            'balance_qty': balance_qty_value,
        }
        if plan_date:
            defaults['plan_date'] = plan_date
        if payload.get('plan_month'):
            defaults['plan_month'] = payload.get('plan_month')
        if actor and not existing_job:
            defaults['created_by'] = actor

        job_obj, created = PlanningJob.objects.update_or_create(
            jc_number=jc_number,
            defaults=defaults,
        )
        if created:
            created_count += 1
        else:
            updated_count += 1
        existing_jobs_by_sku[sku_key] = job_obj
        existing_any_jobs_skus.add(sku_key)
        seen_skus_in_payload.add(sku_key)

    payload['repeat_jobs_synced'] = True
    payload['repeat_jobs_created_count'] = created_count
    payload['repeat_jobs_updated_count'] = updated_count
    payload['repeat_jobs_locked_count'] = locked_count
    payload['repeat_jobs_missing_recipe_count'] = missing_recipe_count
    po_doc.extracted_payload = payload
    po_doc.save(update_fields=['extracted_payload'])

    return {
        'created': created_count,
        'updated': updated_count,
        'locked': locked_count,
        'missing_recipe': missing_recipe_count,
    }



def _sync_new_jobs_for_approved_sku(sku, actor=None):
    """After SKU master approval, refresh matching existing Planning Jobs only."""
    sku_key = _sku_key(sku)
    if not sku_key:
        return {'created': 0, 'updated': 0, 'locked': 0, 'sent': 0, 'missing_jobs': 0}

    recipe = SkuRecipe.objects.filter(sku__iexact=sku, master_data_status='approved').first()
    if not recipe:
        return {'created': 0, 'updated': 0, 'locked': 0, 'sent': 0, 'missing_jobs': 0}

    existing_any_jobs_skus = {
        _sku_key(value)
        for value in PlanningJob.objects.values_list('sku', flat=True)
        if value
    }

    created_count = 0
    updated_count = 0
    locked_count = 0
    sent_count = 0
    missing_job_count = 0

    po_docs = PoDocument.objects.exclude(extracted_payload__isnull=True).order_by('created_at', 'id')
    for po_doc in po_docs:
        payload = po_doc.extracted_payload or {}
        items, _ = _deduplicate_po_items_by_sku(payload.get('items', []))
        target_item = None
        for item in items:
            if _sku_key(item.get('sku')) == sku_key:
                target_item = item
                break

        if not target_item:
            continue

        po_number = (payload.get('po_number') or '').strip()
        if not po_number:
            continue

        existing_job = PlanningJob.objects.filter(po_number=po_number, sku__iexact=sku).order_by('-updated_at', '-id').first()
        if not existing_job:
            missing_job_count += 1
            continue

        if existing_job and _normalize_status(existing_job.status) != 'draft':
            locked_count += 1
            continue

        delivery_date = _parse_iso_date(target_item.get('delivery_date'))
        po_date = _parse_iso_date(payload.get('po_date'))
        plan_date = po_doc.created_at.date() if po_doc and getattr(po_doc, 'created_at', None) else (delivery_date or po_date)
        qty = target_item.get('quantity')
        order_qty = int(qty) if qty is not None else None
        unit_cost_val = target_item.get('unit_cost')
        unit_cost_dec = Decimal(str(unit_cost_val)) if unit_cost_val is not None else None
        jc_number = existing_job.jc_number
        current_requirement = existing_job.requirement

        existing_repeat_flag = (existing_job.repeat_flag or '').strip().lower()
        if existing_repeat_flag in {'new', 'repeat'}:
            forward_as_new = existing_repeat_flag == 'new'
        else:
            prior_jobs_exist = PlanningJob.objects.filter(sku__iexact=sku).exclude(id=existing_job.id).exists()
            forward_as_new = not prior_jobs_exist

        defaults = {
            'po_number': po_number,
            'sku': sku,
            'job_name': recipe.job_name or (target_item.get('job_name') or '').strip() or sku,
            'order_qty': order_qty,
            'department': payload.get('department') or '',
            'destination': payload.get('delivery_location') or '',
            'delivery_date': delivery_date,
            'unit_cost': unit_cost_dec if unit_cost_dec is not None else recipe.default_unit_cost,
            'status': 'draft',
            'repeat_flag': 'New' if forward_as_new else 'Repeat',
            'requirement': _sync_new_sku_requirement(current_requirement, forward_as_new),
            'material': recipe.material,
            'job_process_type': recipe.job_process_type or 'print_and_pack',
            'color_spec': recipe.color_spec,
            'application': recipe.application,
            'size_w_mm': recipe.size_w_mm,
            'size_h_mm': recipe.size_h_mm,
            'ups': recipe.ups,
            'print_sheet_size': recipe.print_sheet_size,
            'purchase_sheet_size': recipe.purchase_sheet_size,
            'purchase_sheet_ups': recipe.purchase_sheet_ups,
            'daily_demand': recipe.daily_demand,
            'plate_set_no': existing_job.plate_set_no,
            'remarks': existing_job.remarks or recipe.notes,
        }

        if not forward_as_new:
            defaults['requirement'] = _append_unique_note_line(
                defaults['requirement'],
                _build_cost_mismatch_note(recipe.default_unit_cost, unit_cost_dec),
            )
        if plan_date:
            defaults['plan_date'] = plan_date

        job_obj, created = PlanningJob.objects.update_or_create(
            jc_number=jc_number,
            defaults=defaults,
        )
        if created:
            created_count += 1
        else:
            updated_count += 1

        existing_any_jobs_skus.add(sku_key)
        sent_count += 1

        sent_to_planning = set(payload.get('new_skus_sent_to_planning') or [])
        sent_to_planning.add(sku)
        payload['new_skus_sent_to_planning'] = sorted(sent_to_planning)
        po_doc.extracted_payload = payload
        po_doc.save(update_fields=['extracted_payload'])

    return {
        'created': created_count,
        'updated': updated_count,
        'locked': locked_count,
        'sent': sent_count,
        'missing_jobs': missing_job_count,
    }



def _merge_po_items_for_existing_po(existing_items, incoming_items):
    """Merge incoming PO lines into existing PO lines without creating duplicates."""
    existing_by_sku = {}
    merged_items = []

    for item in existing_items:
        sku = (item.get('sku') or '').strip()
        sku_key = _sku_key(sku)
        if not sku_key or sku_key in existing_by_sku:
            continue
        item_copy = dict(item)
        existing_by_sku[sku_key] = item_copy
        merged_items.append(item_copy)

    added_skus = []
    updated_skus = []
    ignored_lines = []

    for item in incoming_items:
        sku = (item.get('sku') or '').strip()
        sku_key = _sku_key(sku)
        if not sku_key:
            continue

        incoming_qty = _to_int(item.get('quantity'))
        existing_item = existing_by_sku.get(sku_key)

        if existing_item is None:
            item_copy = dict(item)
            merged_items.append(item_copy)
            existing_by_sku[sku_key] = item_copy
            added_skus.append(sku)
            continue

        existing_qty = _to_int(existing_item.get('quantity'))
        if existing_qty == incoming_qty:
            ignored_lines.append({'sku': sku, 'qty': incoming_qty})
            continue

        # Same SKU but changed qty/fields: treat as correction, not duplicate row.
        for field, value in item.items():
            if value not in (None, ''):
                existing_item[field] = value
        updated_skus.append(sku)

    for idx, item in enumerate(merged_items, start=1):
        item['line_no'] = idx

    return merged_items, sorted(set(added_skus)), sorted(set(updated_skus)), ignored_lines



def _collect_pending_sku_rows(po_docs):
    """Build pending SKU rows from PO documents where SKU recipe is missing."""
    rows = []
    for po_doc in po_docs:
        payload = po_doc.extracted_payload or {}
        items = _po_payload_items(payload)
        if not items:
            continue

        recipe_map = _build_recipe_map(items)
        _, _, _, missing_skus = _annotate_items_with_recipe(items, recipe_map)
        if not missing_skus:
            continue

        item_map = {}
        for item in items:
            key = _sku_key(item.get('sku'))
            if key and key not in item_map:
                item_map[key] = item

        po_number = payload.get('po_number') or '-'
        ignored_skus = {
            _sku_key(s)
            for s in (payload.get('new_skus_ignored') or [])
            if s
        }
        for sku in missing_skus:
            if _sku_key(sku) in ignored_skus:
                continue
            item = item_map.get(_sku_key(sku), {})
            rows.append(
                {
                    'po_doc_id': po_doc.id,
                    'po_number': po_number,
                    'sku': sku,
                    'job_name': (item.get('job_name') or '').strip() or sku,
                    'qty': _format_display_qty(item.get('quantity')),
                    'delivery_date': item.get('delivery_date') or '-',
                    'uploaded_at': po_doc.created_at,
                }
            )

    return rows


MASTER_SYNC_FIELD_LABELS = {
    'job_name': 'Job Name',
    'material': 'Material',
    'job_process_type': 'Job Process',
    'color_spec': 'Print Color',
    'application': 'Application',
    'machine_name': 'Machine',
    'size_w_mm': 'Size W (mm)',
    'size_h_mm': 'Size H (mm)',
    'ups': 'UPS',
    'print_sheet_size': 'Print Sheet Size',
    'purchase_sheet_size': 'Purchase Sheet Size',
    'purchase_sheet_ups': 'Purchase Sheet UPS',
    'daily_demand': 'Daily Demand',
}


def _normalize_sheet_size(value):
    return str(value or '').strip().lower().replace('x', '*').replace(' ', '')


def _master_sync_field_values_equal(left, right, field_name=''):
    if field_name in {'print_sheet_size', 'purchase_sheet_size'}:
        return _normalize_sheet_size(left) == _normalize_sheet_size(right)
    if left is None and right is None:
        return True
    if isinstance(left, Decimal) or isinstance(right, Decimal):
        try:
            return Decimal(str(left or 0)) == Decimal(str(right or 0))
        except Exception:
            return str(left or '') == str(right or '')
    return str(left or '').strip() == str(right or '').strip()


def get_master_data_field_diffs(job):
    recipe = job.approved_sku_recipe
    if not recipe:
        return {}

    diffs = {}
    for field_name, label in MASTER_SYNC_FIELD_LABELS.items():
        job_value = getattr(job, field_name, None)
        recipe_value = getattr(recipe, field_name, None)
        if not _master_sync_field_values_equal(job_value, recipe_value, field_name=field_name):
            diffs[field_name] = {
                'label': label,
                'job': job_value,
                'recipe': recipe_value,
            }
    return diffs


def job_has_master_data_mismatch(job):
    return bool(get_master_data_field_diffs(job))


def can_request_master_data_sync(job):
    if not job.is_active or job.master_data_sync_blocked():
        return False
    if not job.approved_sku_recipe:
        return False
    return job_has_master_data_mismatch(job)


def request_master_data_sync(job, actor, reason=''):
    reason = (reason or '').strip()
    if not reason:
        raise ValueError('A reason is required to request master data sync.')
    if job.master_data_sync_blocked():
        raise ValueError('Completed jobs cannot be synced with revised SKU master data.')
    if not job.approved_sku_recipe:
        raise ValueError('No approved SKU master exists for this job.')
    if not job_has_master_data_mismatch(job):
        raise ValueError('This job already matches the approved SKU master data.')

    job.master_sync_requested = True
    job.master_sync_reason = reason
    job.master_sync_requested_by = actor
    job.master_sync_requested_at = timezone.now()
    job.save(update_fields=[
        'master_sync_requested',
        'master_sync_reason',
        'master_sync_requested_by',
        'master_sync_requested_at',
        'updated_at',
    ])
    return job


def dismiss_master_data_sync_request(job, actor=None):
    job.master_sync_requested = False
    job.master_sync_reason = ''
    job.master_sync_requested_by = None
    job.master_sync_requested_at = None
    job.save(update_fields=[
        'master_sync_requested',
        'master_sync_reason',
        'master_sync_requested_by',
        'master_sync_requested_at',
        'updated_at',
    ])
    return job


def apply_master_data_sync(job, actor):
    from core.jobcard_service import ensure_job_card_from_planning_job
    from core.models import ChangeLog, JOB_CARD_PLANNING_EDITABLE_STATUSES, JobCard

    if job.master_data_sync_blocked():
        raise ValueError('Completed jobs cannot be synced with revised SKU master data.')

    recipe = job.approved_sku_recipe
    if not recipe:
        raise ValueError('No approved SKU master exists for this job.')

    diffs = get_master_data_field_diffs(job)
    if not diffs:
        dismiss_master_data_sync_request(job, actor=actor)
        return job, {'updated_fields': [], 'job_card_refreshed': False}

    field_changes = {}
    update_fields = ['updated_at', 'job_card_version']
    for field_name in MASTER_SYNC_FIELD_LABELS:
        recipe_value = getattr(recipe, field_name, None)
        job_value = getattr(job, field_name, None)
        if not _master_sync_field_values_equal(job_value, recipe_value, field_name=field_name):
            field_changes[field_name] = {
                'label': MASTER_SYNC_FIELD_LABELS[field_name],
                'from': str(job_value if job_value is not None else '-'),
                'to': str(recipe_value if recipe_value is not None else '-'),
            }
            setattr(job, field_name, recipe_value)
            update_fields.append(field_name)

    job.job_card_version = (job.job_card_version or 1) + 1
    job.master_sync_requested = False
    job.master_sync_reason = ''
    job.master_sync_requested_by = None
    job.master_sync_requested_at = None
    job.master_sync_applied_by = actor
    job.master_sync_applied_at = timezone.now()
    update_fields.extend([
        'master_sync_requested',
        'master_sync_reason',
        'master_sync_requested_by',
        'master_sync_requested_at',
        'master_sync_applied_by',
        'master_sync_applied_at',
        'actual_sheet_required',
        'purchase_sheet_required',
        'pkt_value',
        'balance_qty',
        'total_colors',
    ])

    from core.print_colors import apply_print_color_to_planning_job

    apply_print_color_to_planning_job(job)

    with transaction.atomic():
        job.save()
        job_card_refreshed = False
        try:
            job_card = job.job_card
        except JobCard.DoesNotExist:
            job_card = None

        if job_card and job_card.workflow_status in JOB_CARD_PLANNING_EDITABLE_STATUSES:
            ensure_job_card_from_planning_job(job, actor=actor)
            job_card_refreshed = True

        ChangeLog.objects.create(
            entity_type='planning_job',
            record_id=job.pk,
            record_label=str(job),
            action='master_sync',
            changed_by=actor,
            change_reason='Applied approved SKU master data to planning job',
            field_changes=field_changes,
        )

    return job, {
        'updated_fields': list(field_changes.keys()),
        'job_card_refreshed': job_card_refreshed,
    }


def preview_master_sync_calculations(job):
    """Preview sheet math using approved SKU master values (before apply)."""
    import math

    recipe = job.approved_sku_recipe
    if not recipe:
        return None

    net_qty = job.net_print_qty
    if net_qty is None or not recipe.ups:
        return None

    total_sheets = math.ceil(net_qty / recipe.ups) + (job.wastage_sheets or 0)
    purchase_sheets = None
    if recipe.purchase_sheet_ups:
        purchase_sheets = math.ceil(total_sheets / recipe.purchase_sheet_ups)

    pkt_value = None
    if purchase_sheets is not None:
        pkt_value = (Decimal(purchase_sheets) / Decimal('100')).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)

    return {
        'ups': recipe.ups,
        'print_sheet_size': recipe.print_sheet_size,
        'purchase_sheet_size': recipe.purchase_sheet_size,
        'purchase_sheet_ups': recipe.purchase_sheet_ups,
        'total_sheets': total_sheets,
        'purchase_sheets': purchase_sheets,
        'pkt_value': pkt_value,
    }


def job_requires_reopen_for_master_sync(job):
    from core.models import JOB_CARD_PLANNING_EDITABLE_STATUSES, JobCard

    if job.master_data_sync_blocked():
        return False
    try:
        job_card = job.job_card
    except JobCard.DoesNotExist:
        job_card = None
    if job.workflow_status not in {'draft', 'pending_qc'}:
        return True
    if job_card and job_card.workflow_status not in JOB_CARD_PLANNING_EDITABLE_STATUSES:
        return True
    return False


def reopen_and_apply_master_data_sync(job, actor, reason=''):
    from core.jobcard_service import ensure_job_card_from_planning_job, reopen_job_card_for_master_sync
    from core.models import JOB_CARD_PLANNING_EDITABLE_STATUSES, JobCard

    if job.master_data_sync_blocked():
        raise ValueError('Completed jobs cannot be synced with revised SKU master data.')

    recipe = job.approved_sku_recipe
    if not recipe:
        raise ValueError('No approved SKU master exists for this job.')
    if not get_master_data_field_diffs(job):
        raise ValueError('This job already matches the approved SKU master data.')

    reopened_planning = False
    reopened_job_card = False

    with transaction.atomic():
        if job.workflow_status not in {'draft', 'pending_qc'}:
            job.status = 'draft'
            job.issued_to_production = False
            job.save(update_fields=['status', 'issued_to_production', 'updated_at'])
            reopened_planning = True

        try:
            job_card = job.job_card
        except JobCard.DoesNotExist:
            job_card = None

        if job_card and job_card.workflow_status not in JOB_CARD_PLANNING_EDITABLE_STATUSES:
            reopen_job_card_for_master_sync(
                job_card,
                actor=actor,
                reason=reason or 'Reopened for SKU master sync',
            )
            reopened_job_card = True

        job.master_sync_requested = True
        job.master_sync_reason = reason or job.master_sync_reason or 'Reopen and apply SKU master sync'
        job.master_sync_requested_by = actor
        job.master_sync_requested_at = timezone.now()
        job.save(update_fields=[
            'master_sync_requested',
            'master_sync_reason',
            'master_sync_requested_by',
            'master_sync_requested_at',
            'updated_at',
        ])

        job, result = apply_master_data_sync(job, actor=actor)
        ensure_job_card_from_planning_job(job, actor=actor)
        result['job_card_refreshed'] = True
        result['reopened_planning'] = reopened_planning
        result['reopened_job_card'] = reopened_job_card

    return job, result

