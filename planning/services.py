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
from .models import (
    JOB_CANCEL_REQUEST_TYPE,
    PLANNING_CANCEL_REASON_CHOICES,
    PLANNING_STATUS_ALIASES,
    PURCHASE_MATERIAL_ORIGIN_CHOICES,
    JobCardChangeRequest,
    PlanningJob,
    PoDocument,
    SkuRecipe,
)
from workflow.services import _append_unique_note_line, _parse_iso_date, _format_display_qty, _build_cost_mismatch_note, _normalize_status, _to_int, _to_decimal, SKU_MASTER_APPROVAL_REQUIRED_FIELDS, _warning_master_fields
from core.jc_numbering import allocate_next_jc_number

try:
    from reportlab.lib import colors as _rl_colors
    from reportlab.lib.pagesizes import A4 as _RL_A4
    from reportlab.lib.styles import ParagraphStyle as _RLParagraphStyle, getSampleStyleSheet as _rl_get_sample_stylesheet
    from reportlab.lib.units import mm as _rl_mm
    from reportlab.platypus import (
        Paragraph as _RLParagraph,
        SimpleDocTemplate as _RLSimpleDocTemplate,
        Spacer as _RLSpacer,
        Table as _RLTable,
        TableStyle as _RLTableStyle,
    )
    REPORT_PDF_AVAILABLE = True
except ImportError:
    REPORT_PDF_AVAILABLE = False

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
    'print_passes',
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

# Print / purchase sheet: width then * or x or X then height (e.g. 25*36, 28x40, 7.5 x 10.5).
SHEET_SIZE_FORMAT_RE = re.compile(
    r'^\s*(\d+(?:\.\d+)?)\s*([*xX])\s*(\d+(?:\.\d+)?)\s*$'
)


def is_valid_sheet_size(value):
    text = str(value or '').strip()
    if not text or text == '.':
        return False
    return bool(SHEET_SIZE_FORMAT_RE.match(text))


def normalize_sheet_size(value):
    """Normalize sheet size to W*H using * separator."""
    text = str(value or '').strip()
    match = SHEET_SIZE_FORMAT_RE.match(text)
    if not match:
        return text
    return f'{match.group(1)}*{match.group(3)}'


def _parse_layout_positive_int(raw_value):
    text = str(raw_value or '').strip()
    if text == '':
        return None
    try:
        return int(round(float(text)))
    except (TypeError, ValueError):
        return None


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
        'print_passes': _r('print_passes', None),
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
        elif recipe_field in {'print_sheet_size', 'purchase_sheet_size'}:
            value = normalize_sheet_size(raw_value)
            if value and not is_valid_sheet_size(value):
                raise ValueError(
                    f'{recipe_field.replace("_", " ").title()} must use format '
                    'width x height (e.g. 25*36 or 28x40).'
                )
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
            if recipe_field in {'ups', 'purchase_sheet_ups'} and value < 1:
                label = 'Ups' if recipe_field == 'ups' else 'Purchase Sheet Ups'
                raise ValueError(f'{label} must be at least 1.')

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

    # Always transition to pending_review when submit_for_review is enabled,
    # so that it shows up in the QC review queue immediately.
    if submit_for_review and recipe.master_data_status in {'draft', ''}:
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
        if field_name in {'ups', 'purchase_sheet_ups'}:
            value = _parse_layout_positive_int(posted_values.get(field_name))
            if value is None or value < 1:
                missing.append(label)
            continue
        if str(posted_values.get(field_name) or '').strip() == '':
            missing.append(label)
    set_no = str(posted_values.get('set_no') or '').strip()
    new_set_no = str(posted_values.get('new_set_no') or '').strip()
    if not set_no and not new_set_no:
        missing.append('Set No / New Set No')
    return missing


def designer_layout_validation_errors(posted_values):
    """Return human-readable validation errors for designer layout values."""
    errors = []
    for field_name, label in (
        ('print_sheet_size', 'Print Sheet Size'),
        ('purchase_sheet_size', 'Purchase Sheet Size'),
    ):
        value = str(posted_values.get(field_name) or '').strip()
        if not value:
            continue
        if not is_valid_sheet_size(value):
            errors.append(
                f'{label} must use format width x height (e.g. 25*36, 28x40, or 7.5 x 10.5).'
            )
    for field_name, label in (
        ('ups', 'Ups'),
        ('purchase_sheet_ups', 'Purchase Sheet Ups'),
    ):
        value = _parse_layout_positive_int(posted_values.get(field_name))
        if value is not None and value < 1:
            errors.append(f'{label} must be at least 1.')
    return errors


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
    if action == 'save_draft' and 'print_passes' in form.fields:
        form.fields['print_passes'].required = False
        form.fields['print_passes'].widget.attrs.pop('required', None)
    if action == 'send_to_plate_making' and 'product_type' in form.fields:
        form.fields['product_type'].required = False


def get_plate_making_prerequisite_errors(planning_job):
    """Return human-readable blockers before opening plate making."""
    errors = []
    if not (planning_job.material or '').strip():
        errors.append('Material Type is required before Plate Making.')
    if not (planning_job.application or '').strip():
        errors.append('Application is required before Plate Making.')
    if not planning_job.effective_machine_name:
        errors.append('Machine Name is required before Plate Making.')
    if not planning_job.is_cut_and_pack() and not planning_job.effective_print_passes:
        errors.append('No. of Passes is required on SKU master before Plate Making.')
    if planning_job.is_merge_member_follower:
        group = planning_job.active_merge_group
        lead_jc = group.lead_job.jc_number if group and group.lead_job else ''
        errors.append(
            f'This job is merged into layout {group.code if group else ""}. '
            f'Plates are made once on the lead job {lead_jc}, not per SKU.'
        )
    return errors


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


def ensure_draft_planning_job_for_po_sku(po_doc, sku, *, actor=None, recipe=None):
    """Return a draft PlanningJob for a PO line, creating one when PO sync skipped it."""
    payload = po_doc.extracted_payload or {}
    po_number = (payload.get('po_number') or '').strip()
    sku_key = _sku_key(sku)
    if not po_number or not sku_key:
        return None

    existing_job = (
        PlanningJob.objects.filter(po_number=po_number, sku__iexact=sku)
        .order_by('-updated_at', '-id')
        .first()
    )
    if existing_job:
        if _normalize_status(existing_job.status) != 'draft':
            return None
        return existing_job

    items, _ = _deduplicate_po_items_by_sku(payload.get('items', []))
    item = next((row for row in items if _sku_key(row.get('sku')) == sku_key), None)
    if not item:
        return None

    po_date = _parse_iso_date(payload.get('po_date'))
    approval_date = _parse_iso_date(payload.get('approval_date'))
    po_approval_date = approval_date or po_date
    delivery_date = _parse_iso_date(item.get('delivery_date'))
    intake_plan_date = (
        po_doc.created_at.date()
        if po_doc and getattr(po_doc, 'created_at', None)
        else (po_date or delivery_date)
    )
    recipe_map = _build_recipe_map(items)
    recipe = recipe or recipe_map.get(sku_key)

    from .sku_classification import repeat_flag_value_for_po_line

    repeat_flag_value = repeat_flag_value_for_po_line(
        item,
        po_number=po_number,
        po_doc_created_at=getattr(po_doc, 'created_at', None),
        po_doc_id=getattr(po_doc, 'id', None),
        recipe=recipe,
        existing_job=None,
    )

    qty = item.get('quantity')
    order_qty = int(qty) if qty is not None else None
    unit_cost_val = item.get('unit_cost')
    unit_cost_dec = Decimal(str(unit_cost_val)) if unit_cost_val is not None else None
    sku_value = (item.get('sku') or '').strip()
    fallback_job_name = (item.get('job_name') or '').strip() or sku_value
    if (item.get('job_name') or '').strip():
        job_name_value = item.get('job_name').strip()
    elif recipe and (recipe.job_name or '').strip():
        job_name_value = recipe.job_name
    else:
        job_name_value = fallback_job_name

    defaults = {
        'po_number': po_number,
        'pr_reference': (payload.get('pr_number') or '').strip(),
        'sku': sku_value,
        'job_name': job_name_value,
        'order_qty': order_qty,
        'department': payload.get('department') or '',
        'destination': payload.get('delivery_location') or '',
        'delivery_date': delivery_date,
        'unit_cost': unit_cost_dec if unit_cost_dec is not None else (recipe.default_unit_cost if recipe else None),
        'status': 'draft',
        'repeat_flag': repeat_flag_value,
        'material': (recipe.material if recipe else '') or (item.get('material') or '').strip(),
        'application': (recipe.application if recipe else '') or (item.get('application') or '').strip(),
        'machine_name': (recipe.machine_name if recipe else '') or (item.get('machine_name') or item.get('machine') or '').strip(),
        'plate_set_no': (recipe.plate_set_no if recipe else '') or (item.get('plate_set_no') or item.get('p_set_no') or '').strip(),
        'color_spec': (recipe.color_spec if recipe else '') or (item.get('color_spec') or item.get('color') or '').strip(),
        'plan_date': intake_plan_date,
        'plan_month': payload.get('plan_month') or _plan_month_label_from_date(intake_plan_date),
    }
    if po_approval_date:
        defaults['po_approval_date'] = po_approval_date
    if actor:
        defaults['created_by'] = actor

    job = PlanningJob.objects.create(
        jc_number=allocate_next_jc_number(intake_plan_date),
        **defaults,
    )
    if recipe:
        sync_planning_job_fields_to_sku_recipe(job, recipe)
    return job


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


def build_job_card_merge_context(job):
    """Merge banner data for a job card, or None when the job is not merged.

    Single source of truth for the printed HTML card, the PDF card and the
    layout-builder field values, so the three can never disagree.
    """
    item = job.active_merge_item
    if not item:
        return None
    group = item.merge_group
    members = [
        {
            'jc_number': member.planning_job.jc_number,
            'sku': member.planning_job.sku,
            'allocated_ups': member.allocated_ups,
            'planned_produced_qty': member.planned_produced_qty,
            'source_awc_no': member.source_awc_no,
            'is_lead': member.is_lead,
        }
        for member in group.items.select_related('planning_job').order_by('-allocated_ups', 'id')
    ]
    return {
        'code': group.code,
        'artwork_code': group.artwork_code,
        'is_lead': item.is_lead,
        'lead_jc': group.lead_job.jc_number if group.lead_job else '',
        'allocated_ups': item.allocated_ups,
        'planned_produced_qty': item.planned_produced_qty,
        'total_sheet_ups': group.total_sheet_ups,
        'run_sheets': group.run_sheets,
        'combined_impressions': group.combined_impressions(),
        'print_sheet_size': group.print_sheet_size,
        'member_count': len(members),
        'members': members,
        'role_label': 'LEAD — combined layout' if item.is_lead else f'Printed with {group.lead_job.jc_number if group.lead_job else ""}',
    }


def merge_layout_master_data_report(group):
    """Per-member missing-master-field report for a merge group's print gate.

    Returns a list of (planning_job, [missing labels]); an empty list means every
    member is ready for the combined run.
    """
    report = []
    for item in group.items.select_related('planning_job'):
        job = item.planning_job
        recipe = job.sku_recipe
        missing = _missing_required_master_fields(
            recipe, recipe.job_name if recipe else job.job_name, allow_missing_plate_set_no=True
        )
        # AWC is optional for a merge — new artwork is drawn into the layout.
        missing = [m for m in missing if m != 'AWC #']
        if missing:
            report.append((job, missing))
    return report


def divide_combined_wastage(items, combined_wastage):
    """Split the run's total wastage across members by ups share, sum preserved.

    A wasted sheet spoils each SKU's ups on it, so the run's wastage is shared;
    dividing it by ups share avoids counting the same waste N times across the
    member job cards. Returns {planning_job_id: wastage_sheets}.
    """
    combined_wastage = int(combined_wastage or 0)
    total_ups = sum(item.allocated_ups for item in items) or 1
    shares = {}
    running = 0
    ordered = sorted(items, key=lambda it: it.allocated_ups, reverse=True)
    for index, item in enumerate(ordered):
        if index == len(ordered) - 1:
            shares[item.planning_job_id] = combined_wastage - running  # remainder to last
        else:
            share = round(combined_wastage * item.allocated_ups / total_ups)
            shares[item.planning_job_id] = share
            running += share
    return shares


def approve_merge_layout(group, actor=None, combined_wastage=None, material_origin=None):
    """Group-level production approval for a combined layout.

    Stands in for each member SKU's individual QC/PM/release gate. Requires every
    member's master data to be complete AND one combined wastage + material origin
    for the shared run. Those are reflected onto every member (material origin the
    same; wastage divided by ups share) so each member card satisfies its own QC
    fields and the whole sheet can print as one job.
    """
    from django.core.exceptions import ValidationError

    from core.jobcard_service import approve_card_for_merged_run, ensure_job_card_from_planning_job

    if group.status == 'cancelled':
        raise ValidationError('This merge group is cancelled.')
    if not group.lead_job_id:
        raise ValidationError('This group has no lead job.')

    report = merge_layout_master_data_report(group)
    if report:
        lines = '; '.join(
            f'{job.jc_number} ({job.sku}): {", ".join(missing)}' for job, missing in report
        )
        raise ValidationError(
            'Complete master data for every SKU before approving the combined layout — '
            + lines
        )

    # Fall back to any value already captured on the group so re-approval works.
    if combined_wastage in (None, ''):
        combined_wastage = group.combined_wastage_sheets
    if material_origin in (None, ''):
        material_origin = group.purchase_material_origin
    if combined_wastage in (None, ''):
        raise ValidationError('Enter the combined run wastage (sheets) before approving the layout.')
    valid_origins = {code for code, _ in PURCHASE_MATERIAL_ORIGIN_CHOICES}
    if material_origin not in valid_origins:
        raise ValidationError('Select the purchase material origin for the combined run.')
    combined_wastage = int(combined_wastage)

    items = list(group.items.select_related('planning_job'))
    wastage_shares = divide_combined_wastage(items, combined_wastage)

    with transaction.atomic():
        group.combined_wastage_sheets = combined_wastage
        group.purchase_material_origin = material_origin
        group.layout_approved_by = actor
        group.layout_approved_at = timezone.now()

        for item in items:
            job = item.planning_job
            # Reflect the combined inputs onto the member so its own QC fields pass.
            job.wastage_sheets = wastage_shares.get(job.id, 0)
            job.purchase_material_origin = material_origin
            job.save(update_fields=['wastage_sheets', 'purchase_material_origin', 'updated_at'])

            job_card, _ = ensure_job_card_from_planning_job(job, actor=actor)
            approve_card_for_merged_run(job_card, group, actor=actor)

        # Hand off to the graphics designer: create the one combined plate request
        # into their queue and move the group to 'artwork_requested'. The designer
        # builds the combined artwork on that request and sends it to the vendor.
        from printing_plates.services import create_combined_plate_for_group

        create_combined_plate_for_group(group, actor=actor)
        group.status = 'artwork_requested'
        group.designer_requested_at = timezone.now()
        group.save(update_fields=[
            'combined_wastage_sheets', 'purchase_material_origin', 'status',
            'layout_approved_by', 'layout_approved_at', 'designer_requested_at',
        ])

    return group


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

    # Smart-merge marker — the merged card stays the normal single-SKU card, with
    # just a small flag line here and a diagonal "DO NOT PRINT SEPARATELY" watermark
    # drawn on every page (see _merge_watermarker below). The combined-run detail
    # lives on the separate Combined Layout Sheet, not on the card.
    merge = build_job_card_merge_context(job)
    if merge:
        flag_style = ParagraphStyle(
            'MergeFlag', parent=normal, fontName='Helvetica-Bold', fontSize=9,
            leading=11, textColor=colors.HexColor('#7a1c12'),
            backColor=colors.HexColor('#ffe5e0'), borderColor=colors.HexColor('#c0392b'),
            borderWidth=1, borderPadding=4,
        )
        recorded = '' if merge['is_lead'] else f" Printing is recorded on {merge['lead_jc']}."
        flag_text = (
            f"MERGED IN {merge['code']} — see the Combined Layout Sheet for the run. "
            f"Do not print or plate this SKU separately.{recorded} "
            f"This card is for cutting, packing, QC and dispatch."
        )
        story.extend([Paragraph(flag_text, flag_style), Spacer(1, 8)])

    header_data = [
        [Paragraph('JOB CARD #', label_style), _format_job_value(job.jc_number), Paragraph('PO #', label_style), _format_job_value(job.po_number)],
        [Paragraph('DATE', label_style), _format_job_value(job.plan_date), Paragraph('STATUS', label_style), _format_job_value(job.effective_status_label)],
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
        [Paragraph('UPS', label_style),
         _format_job_value(f"{merge['total_sheet_ups']} (this SKU: {merge['allocated_ups']})" if merge else job.ups),
         Paragraph('PRINT SHEETS', label_style), _format_job_value(job.print_sheets)],
        [Paragraph('ACTUAL SHEETS', label_style),
         _format_job_value(f"{merge['run_sheets']} (combined run)" if merge else job.calculated_sheets_required),
         Paragraph('WASTAGE', label_style), _format_job_value(job.wastage_sheets)],
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
        [Paragraph('Purch sheet Qty', label_style), _format_job_value(job.purchase_sheet_required), Paragraph('Remarks', label_style), _paragraph_text(job.remarks_display or job.requirement or '-'), '', ''],
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

    def _merge_watermarker(canvas, doc_):
        """Diagonal repeating 'DO NOT PRINT SEPARATELY' watermark on merged cards."""
        canvas.saveState()
        canvas.setFont('Helvetica-Bold', 26)
        canvas.setFillColor(colors.HexColor('#c0392b'))
        try:
            canvas.setFillAlpha(0.12)
        except Exception:
            pass
        text = f"PRINTS WITH MERGE {merge['code']} - DO NOT PRINT SEPARATELY"
        canvas.translate(105 * mm, 148 * mm)
        canvas.rotate(32)
        for offset in (-170, -85, 0, 85, 170):
            canvas.drawCentredString(0, offset, text)
        canvas.restoreState()

    if merge:
        doc.build(story, onFirstPage=_merge_watermarker, onLaterPages=_merge_watermarker)
    else:
        doc.build(story)
    return buffer.getvalue()


def build_job_history_report_pdf_bytes(job):
    """
    Full recorded history for one JC — planning/PO reference, SKU master data,
    and every plate request, printing entry, packing entry, and dispatch entry
    on file — as a single downloadable A4 PDF. Distinct from the blank Job
    Card traveler (job_card_print/pdf): this is a read-only report of what
    actually happened on the job, for review, audit, or sharing outside the
    system.
    """
    if not REPORT_PDF_AVAILABLE:
        raise RuntimeError('reportlab is required to generate the job history report PDF. Install reportlab and restart the server.')

    colors = _rl_colors
    mm = _rl_mm

    def _fmt(value):
        if value in (None, ''):
            return '-'
        return str(value)

    def _p(value, style):
        return _RLParagraph(_fmt(value).replace('\n', '<br/>'), style)

    buffer = io.BytesIO()
    doc = _RLSimpleDocTemplate(
        buffer,
        pagesize=_RL_A4,
        leftMargin=12 * mm,
        rightMargin=12 * mm,
        topMargin=12 * mm,
        bottomMargin=12 * mm,
    )

    styles = _rl_get_sample_stylesheet()
    normal = styles['Normal']
    normal.fontName = 'Helvetica'
    normal.fontSize = 8.5
    normal.leading = 10.5

    title_style = _RLParagraphStyle('HRTitle', parent=normal, fontName='Helvetica-Bold', fontSize=16, leading=18)
    subtitle_style = _RLParagraphStyle('HRSubtitle', parent=normal, fontName='Helvetica-Bold', fontSize=11, leading=13)
    section_style = _RLParagraphStyle('HRSection', parent=normal, fontName='Helvetica-Bold', fontSize=10.5, leading=12, spaceBefore=6, spaceAfter=2, textColor=colors.HexColor('#1a1a1a'))
    label_style = _RLParagraphStyle('HRLabel', parent=normal, fontName='Helvetica-Bold', fontSize=8, leading=9.5)
    cell_style = _RLParagraphStyle('HRCell', parent=normal, fontSize=8, leading=9.5)
    header_row_style = _RLParagraphStyle('HRHeaderRow', parent=normal, fontName='Helvetica-Bold', fontSize=8, leading=9.5, textColor=colors.white)

    story = [
        _RLParagraph('UTOPIA PRINTING & PACKAGING', title_style),
        _RLSpacer(1, 3),
        _RLParagraph('JOB HISTORY REPORT', subtitle_style),
        _RLSpacer(1, 2),
        _RLParagraph(f'Generated {timezone.now().strftime("%Y-%m-%d %H:%M")}', cell_style),
        _RLSpacer(1, 10),
    ]

    def _grid_table(rows, col_widths):
        table = _RLTable(rows, colWidths=col_widths, hAlign='LEFT')
        table.setStyle(_RLTableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('GRID', (0, 0), (-1, -1), 0.35, colors.HexColor('#888888')),
            ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#dedede')),
        ]))
        return table

    # --- Reference ---
    story.append(_RLParagraph('PO INTAKE / PLANNING JOB', section_style))
    ref_rows = [
        [_p('JC No', label_style), _p(job.jc_number, cell_style), _p('PO Number', label_style), _p(job.po_number, cell_style)],
        [_p('SKU', label_style), _p(job.sku, cell_style), _p('Job Name', label_style), _p(job.job_name, cell_style)],
        [_p('PR Reference', label_style), _p(job.pr_reference, cell_style), _p('Order Quantity', label_style), _p(job.order_qty, cell_style)],
        [_p('Status', label_style), _p(job.effective_status_label, cell_style), _p('Repeat Flag', label_style), _p(job.repeat_flag, cell_style)],
        [_p('Planning Stage', label_style), _p(job.get_planning_stage_display() if job.planning_stage else '-', cell_style), _p('PO Approval Date', label_style), _p(job.po_approval_date, cell_style)],
        [_p('Delivery Date', label_style), _p(job.delivery_date, cell_style), _p('Delivery Location', label_style), _p(job.destination, cell_style)],
        [_p('Department', label_style), _p(job.department, cell_style), _p('Stock Qty (pcs)', label_style), _p(job.stock_qty, cell_style)],
        [_p('Planned Wastage (Sheets)', label_style), _p(job.wastage_sheets, cell_style), '', ''],
    ]
    story.append(_grid_table(ref_rows, [30 * mm, 60 * mm, 30 * mm, 60 * mm]))
    story.append(_RLSpacer(1, 10))

    # --- SKU master data ---
    recipe = job.approved_sku_recipe or job.sku_recipe
    if recipe:
        story.append(_RLParagraph('SKU MASTER DATA', section_style))
        recipe_rows = [
            [_p('Job Process', label_style), _p(recipe.get_job_process_type_display() if recipe.job_process_type else '-', cell_style), _p('Material', label_style), _p(recipe.material, cell_style)],
            [_p('Print Color', label_style), _p(recipe.color_spec, cell_style), _p('Application', label_style), _p(recipe.application, cell_style)],
            [_p('Product Type', label_style), _p(recipe.product_type, cell_style), _p('Machine Name', label_style), _p(recipe.machine_name, cell_style)],
            [_p('No. of Passes', label_style), _p(recipe.print_passes, cell_style), _p('Plate Set No.', label_style), _p(recipe.plate_set_no, cell_style)],
            [_p('Size W/H (mm)', label_style), _p(f'{_fmt(recipe.size_w_mm)} x {_fmt(recipe.size_h_mm)}', cell_style), _p('Print Sheet Size', label_style), _p(recipe.print_sheet_size, cell_style)],
            [_p('Purchase Sheet Size', label_style), _p(recipe.purchase_sheet_size, cell_style), _p('Purchase Sheet Ups', label_style), _p(recipe.purchase_sheet_ups, cell_style)],
            [_p('UPS', label_style), _p(recipe.ups, cell_style), _p('Daily Demand', label_style), _p(recipe.daily_demand, cell_style)],
            [_p('Unit Cost', label_style), _p(recipe.default_unit_cost, cell_style), _p('AWC No.', label_style), _p(recipe.awc_no, cell_style)],
            [_p('Die Cutting', label_style), _p(normalize_die_cutting(recipe.die_cutting), cell_style), _p('Recipe Status', label_style), _p(recipe.get_master_data_status_display() if recipe.master_data_status else '-', cell_style)],
            [_p('SKU Remarks', label_style), _p(recipe.remarks, cell_style), _p('SKU Notes', label_style), _p(recipe.notes, cell_style)],
        ]
        story.append(_grid_table(recipe_rows, [30 * mm, 60 * mm, 30 * mm, 60 * mm]))
        story.append(_RLSpacer(1, 10))

    # --- Plate requests ---
    plate_requests = list(job.plate_requests.select_related('requested_by').order_by('requested_at', 'created_at'))
    story.append(_RLParagraph(f'PLATE REQUESTS ({len(plate_requests)})', section_style))
    if plate_requests:
        pr_rows = [[
            _p('#', header_row_style), _p('Type', header_row_style), _p('Status', header_row_style),
            _p('Set / AWC', header_row_style), _p('Requested', header_row_style), _p('By', header_row_style),
            _p('Remarks', header_row_style),
        ]]
        for req in plate_requests:
            pr_rows.append([
                _p(req.pk, cell_style),
                _p(req.plate_request_type, cell_style),
                _p(req.status_label_display, cell_style),
                _p(f'{_fmt(req.display_set_no)} / {_fmt(req.display_awc_no)}', cell_style),
                _p(req.requested_at.strftime('%Y-%m-%d') if req.requested_at else '-', cell_style),
                _p(req.requested_by.username if req.requested_by else '-', cell_style),
                _p(req.remarks, cell_style),
            ])
        table = _RLTable(pr_rows, colWidths=[10 * mm, 22 * mm, 26 * mm, 26 * mm, 22 * mm, 18 * mm, 56 * mm], hAlign='LEFT', repeatRows=1)
        table.setStyle(_RLTableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('GRID', (0, 0), (-1, -1), 0.3, colors.HexColor('#aaaaaa')),
            ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#555555')),
        ]))
        story.append(table)
    else:
        story.append(_RLParagraph('No plate requests on file.', cell_style))
    story.append(_RLSpacer(1, 10))

    job_card = getattr(job, 'job_card', None)
    printing_entries = []
    packing_entries = []
    dispatch_entries = []
    wastage_metrics = None
    if job_card:
        printing_entries = list(job_card.productions.filter(is_active=True, entry_type='printing').select_related('machine', 'operator').order_by('date', 'id'))
        packing_entries = list(job_card.productions.filter(is_active=True, entry_type='packing').select_related('sorter', 'created_by').order_by('date', 'id'))
        dispatch_entries = list(job_card.dispatch_set.filter(is_active=True).select_related('created_by').order_by('dispatch_date', 'id'))
        from core.services import compute_job_card_wastage_metrics
        wastage_metrics = compute_job_card_wastage_metrics(job_card)

    # --- Printing entries ---
    story.append(_RLParagraph(f'PRINTING ENTRIES ({len(printing_entries)})', section_style))
    total_good_sheets = 0
    total_print_waste_sheets = 0
    if printing_entries:
        pe_rows = [[
            _p('Date', header_row_style), _p('Pass', header_row_style), _p('Machine', header_row_style),
            _p('Operator', header_row_style), _p('Shift', header_row_style), _p('Impressions', header_row_style),
            _p('Good Sheets', header_row_style), _p('Waste', header_row_style), _p('Remarks', header_row_style),
        ]]
        for entry in printing_entries:
            total_good_sheets += entry.output_sheets or 0
            total_print_waste_sheets += entry.waste_sheets or 0
            pe_rows.append([
                _p(entry.date, cell_style),
                _p(entry.print_pass_number, cell_style),
                _p(entry.machine.name if entry.machine else '-', cell_style),
                _p(entry.operator.name if entry.operator else '-', cell_style),
                _p(entry.get_shift_display() if entry.shift else '-', cell_style),
                _p(entry.impressions, cell_style),
                _p(entry.output_sheets, cell_style),
                _p(entry.waste_sheets, cell_style),
                _p(entry.remark_notes, cell_style),
            ])
        table = _RLTable(pe_rows, colWidths=[22 * mm, 10 * mm, 18 * mm, 22 * mm, 11 * mm, 17 * mm, 17 * mm, 13 * mm, 46 * mm], hAlign='LEFT', repeatRows=1)
        table.setStyle(_RLTableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('GRID', (0, 0), (-1, -1), 0.3, colors.HexColor('#aaaaaa')),
            ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#555555')),
        ]))
        story.append(table)
    else:
        story.append(_RLParagraph('No printing entries on file.', cell_style))
    story.append(_RLSpacer(1, 10))

    # --- Packing entries ---
    story.append(_RLParagraph(f'PACKING ENTRIES ({len(packing_entries)})', section_style))
    total_packed = 0
    total_sorting_waste = 0
    if packing_entries:
        pk_rows = [[
            _p('Date', header_row_style), _p('Sorter', header_row_style), _p('Shift', header_row_style),
            _p('Packed Qty', header_row_style), _p('Sorting Waste', header_row_style), _p('Remarks', header_row_style),
        ]]
        for entry in packing_entries:
            total_packed += entry.packing_qty or 0
            total_sorting_waste += entry.sorting_waste_qty or 0
            pk_rows.append([
                _p(entry.date, cell_style),
                _p(entry.sorter.name if entry.sorter else '-', cell_style),
                _p(entry.get_shift_display() if entry.shift else '-', cell_style),
                _p(entry.packing_qty, cell_style),
                _p(entry.sorting_waste_qty, cell_style),
                _p(entry.remark_notes, cell_style),
            ])
        table = _RLTable(pk_rows, colWidths=[22 * mm, 26 * mm, 16 * mm, 24 * mm, 26 * mm, 66 * mm], hAlign='LEFT', repeatRows=1)
        table.setStyle(_RLTableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('GRID', (0, 0), (-1, -1), 0.3, colors.HexColor('#aaaaaa')),
            ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#555555')),
        ]))
        story.append(table)
    else:
        story.append(_RLParagraph('No packing entries on file.', cell_style))
    story.append(_RLSpacer(1, 10))

    # --- Dispatch entries ---
    story.append(_RLParagraph(f'DISPATCH ENTRIES ({len(dispatch_entries)})', section_style))
    total_dispatched = 0
    if dispatch_entries:
        dp_rows = [[
            _p('Date', header_row_style), _p('DC No.', header_row_style), _p('Qty', header_row_style),
            _p('Added By', header_row_style),
        ]]
        for entry in dispatch_entries:
            total_dispatched += entry.dispatch_qty or 0
            dp_rows.append([
                _p(entry.dispatch_date, cell_style),
                _p(entry.dc_no, cell_style),
                _p(entry.dispatch_qty, cell_style),
                _p(entry.created_by.username if entry.created_by else '-', cell_style),
            ])
        table = _RLTable(dp_rows, colWidths=[28 * mm, 40 * mm, 40 * mm, 40 * mm], hAlign='LEFT', repeatRows=1)
        table.setStyle(_RLTableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('GRID', (0, 0), (-1, -1), 0.3, colors.HexColor('#aaaaaa')),
            ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#555555')),
        ]))
        story.append(table)
    else:
        story.append(_RLParagraph('No dispatch entries on file.', cell_style))
    story.append(_RLSpacer(1, 10))

    # --- Summary ---
    story.append(_RLParagraph('SUMMARY', section_style))
    order_qty = job.order_qty or 0
    balance = order_qty - total_dispatched
    summary_rows = [
        [_p('Order Qty (pcs)', label_style), _p(order_qty, cell_style), _p('Total Printed (sheets)', label_style), _p(total_good_sheets, cell_style)],
        [_p('Total Packed (pcs)', label_style), _p(total_packed, cell_style), _p('Total Dispatched (pcs)', label_style), _p(total_dispatched, cell_style)],
        [_p('Total Print Waste (sheets)', label_style), _p(total_print_waste_sheets, cell_style), _p('Total Sorting Waste (pcs)', label_style), _p(total_sorting_waste, cell_style)],
        [_p('Stock Qty (pcs)', label_style), _p(job.stock_qty, cell_style), _p('Balance Remaining (pcs)', label_style), _p(balance, cell_style)],
    ]
    if wastage_metrics:
        summary_rows.append([
            _p('Wastage Status', label_style), _p(wastage_metrics['wastage_status'], cell_style),
            _p('Total Wastage (pcs)', label_style), _p(wastage_metrics['total_wastage_pcs'], cell_style),
        ])
        summary_rows.append([
            _p('Total Wastage %', label_style), _p(f"{wastage_metrics['total_wastage_pct']}%", cell_style), '', '',
        ])
    story.append(_grid_table(summary_rows, [40 * mm, 40 * mm, 40 * mm, 40 * mm]))

    doc.build(story)
    return buffer.getvalue()


def _sku_key(sku):
    return (sku or '').strip().upper()


def document_type_label(po_number):
    """Return 'WO' or 'PO' based on the document number's prefix, for display/messages."""
    return 'WO' if (po_number or '').strip().upper().startswith('WO-') else 'PO'



def _missing_required_master_fields(recipe, fallback_job_name='', *, allow_missing_plate_set_no=False):
    missing = []
    cut_and_pack = bool(
        recipe and (getattr(recipe, 'job_process_type', '') or 'print_and_pack') == 'cut_and_pack'
    )
    skip_for_cut_and_pack = {'color_spec', 'print_passes'}

    if not recipe:
        fallback = (fallback_job_name or '').strip()
        return [
            label
            for field, label in SKU_MASTER_APPROVAL_REQUIRED_FIELDS
            if not (field == 'job_name' and fallback)
            and not (allow_missing_plate_set_no and field == 'plate_set_no')
        ]

    for field, label in SKU_MASTER_APPROVAL_REQUIRED_FIELDS:
        if cut_and_pack and field in skip_for_cut_and_pack:
            continue
        if allow_missing_plate_set_no and field == 'plate_set_no':
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

    frozen_statuses = {'qc_approved', 'released', 'in_production', 'completed', 'closed'}
    job_status = (job.workflow_status or '').strip().lower()
    if job_status not in frozen_statuses:
        if (recipe.job_process_type or 'print_and_pack') == 'cut_and_pack':
            if job.print_passes is not None:
                job.print_passes = None
                update_fields.append('print_passes')
        elif recipe.print_passes:
            master_passes = int(recipe.print_passes)
            if job.print_passes != master_passes:
                job.print_passes = master_passes
                update_fields.append('print_passes')

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

    # Plate Set No. is warning-only (optional on SKU master); do not block QC submission.
    missing_master = _missing_required_master_fields(
        approved_recipe,
        job.job_name,
        allow_missing_plate_set_no=True,
    )
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
        if field_name == 'plate_set_no':
            continue
        if field_name in {'machine_name', 'print_passes'}:
            blockers.append(
                f'{error_message} Update these on the locked SKU master (Reopen SKU), then re-approve.'
            )
        else:
            blockers.append(error_message)

    return blockers


def get_job_qc_submission_warnings(job):
    """Non-blocking warnings before Send to QC (e.g. optional Plate Set No.)."""
    if job.is_cut_and_pack():
        return []

    warnings = []
    approved_recipe = job.approved_sku_recipe
    if approved_recipe:
        warning_fields = _warning_master_fields(approved_recipe, job.job_name)
        if warning_fields:
            warnings.append(
                'SKU master is missing optional field(s): '
                f'{", ".join(warning_fields)}. You can send to QC now; update when available.'
            )
    elif not str(job.effective_plate_set_no or '').strip():
        warnings.append(
            'Plate Set No. is not set. You can send to QC now; update when available.'
        )
    return warnings


def preview_job_qc_submission_blockers(job):
    return get_job_qc_submission_blockers(job, apply_recipe_sync=False)


def preview_job_qc_submission_warnings(job):
    return get_job_qc_submission_warnings(job)


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



def _annotate_items_with_recipe(items, recipe_map, current_po_number=None, po_doc_created_at=None, po_doc_id=None, sku_doc_index=None):
    from .sku_classification import annotate_items_repeat_new

    return annotate_items_repeat_new(
        items,
        recipe_map,
        po_number=current_po_number,
        po_doc_created_at=po_doc_created_at,
        po_doc_id=po_doc_id,
        sku_doc_index=sku_doc_index,
    )



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



def _history_repeat_new_counts(items, recipe_map=None, current_po_number=None):
    """Classify Repeat/New from approved SKU recipes, not historical PlanningJobs."""
    if recipe_map is None:
        recipe_map = _build_recipe_map(items)
    _, repeat_count, new_count, _ = _annotate_items_with_recipe(items, recipe_map, current_po_number)
    return repeat_count, new_count


def _po_dates_from_payload(payload):
    """Return (approval_date, po_date) parsed from a PO document payload."""
    if not isinstance(payload, dict):
        return None, None
    approval_date = _parse_iso_date(payload.get('approval_date'))
    po_date = _parse_iso_date(payload.get('po_date'))
    return approval_date, po_date


def get_po_approval_date_for_job(job):
    """Resolve PO approval date for display/sync; falls back to PO dated when approval is missing."""
    if getattr(job, 'po_approval_date', None):
        return job.po_approval_date

    if hasattr(job, 'po_documents'):
        for po_document in job.po_documents.order_by('-created_at'):
            approval_date, po_date = _po_dates_from_payload(po_document.extracted_payload or {})
            if approval_date:
                return approval_date
            if po_date:
                return po_date

    if job.po_number:
        po_document = (
            PoDocument.objects.filter(
                extracted_payload__po_number__iexact=job.po_number,
                extraction_status='processed',
            )
            .order_by('-created_at')
            .first()
        )
        if po_document:
            approval_date, po_date = _po_dates_from_payload(po_document.extracted_payload or {})
            if approval_date:
                return approval_date
            if po_date:
                return po_date

    from core.models import ChangeLog

    job_card = getattr(job, 'job_card', None)
    if not job_card:
        return None

    logs = ChangeLog.objects.filter(
        entity_type='job_card',
        record_id=job_card.pk,
    ).order_by('-created_at')
    for log in logs:
        field_changes = log.field_changes if isinstance(log.field_changes, dict) else {}
        status_change = field_changes.get('status') if isinstance(field_changes, dict) else None
        if not isinstance(status_change, dict):
            continue
        to_status = str(status_change.get('to') or '').strip().lower()
        if to_status in {'production_approved', 'qc_approved'} and log.created_at:
            return log.created_at.date()
    return None


def _plan_month_label_from_date(value):
    if not value:
        return ''
    if isinstance(value, datetime):
        value = value.date()
    if isinstance(value, date):
        return value.strftime('%B')
    return ''


def get_planning_intake_date_for_job(job):
    """When the job entered planning (PO upload / create), not customer delivery."""
    if job.po_number:
        doc = (
            PoDocument.objects.filter(
                extracted_payload__po_number__iexact=job.po_number,
                extraction_status='processed',
            )
            .order_by('created_at')
            .first()
        )
        if doc and doc.created_at:
            return doc.created_at.date()

    if hasattr(job, 'po_documents'):
        doc = job.po_documents.order_by('created_at').first()
        if doc and doc.created_at:
            return doc.created_at.date()

    if job.created_at:
        return job.created_at.date()

    return None


def get_planning_month_label_for_job(job):
    """Month for manual working: PO planning intake month, not delivery month."""
    month_text = (job.plan_month or '').strip()
    if month_text:
        return month_text

    intake_date = get_planning_intake_date_for_job(job)
    if intake_date:
        return _plan_month_label_from_date(intake_date)

    if job.plan_date and (not job.delivery_date or job.plan_date != job.delivery_date):
        return _plan_month_label_from_date(job.plan_date)

    return _plan_month_label_from_date(job.plan_date)


def _sync_repeat_jobs_from_po(po_doc, actor=None, bypass_recipe_check=False):
    """Create or update draft planning jobs for all PO lines from one PO document."""
    payload = po_doc.extracted_payload or {}
    items, _ = _deduplicate_po_items_by_sku(payload.get('items', []))
    po_number = (payload.get('po_number') or '').strip()
    pr_number = (payload.get('pr_number') or '').strip()
    po_date = _parse_iso_date(payload.get('po_date'))
    approval_date = _parse_iso_date(payload.get('approval_date'))
    po_approval_date = approval_date or po_date
    delivery_location = payload.get('delivery_location') or ''
    department = payload.get('department') or ''

    if not items:
        return {'created': 0, 'updated': 0, 'locked': 0, 'missing_recipe': 0, 'pr_matched': []}

    item_sku_keys = {_sku_key(item.get('sku')) for item in items if item.get('sku')}
    existing_any_jobs_skus = set()
    if item_sku_keys:
        sku_any_query = Q()
        for sku_key in item_sku_keys:
            sku_any_query |= Q(sku__iexact=sku_key)
        query_qs = PlanningJob.objects.filter(sku_any_query)
        if po_number:
            query_qs = query_qs.exclude(po_number=po_number)
        existing_any_jobs_skus = {
            _sku_key(sku)
            for sku in query_qs.values_list('sku', flat=True)
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

    # Reconcile PR-only jobs (no po_number yet) with the arriving real PO:
    # match by SKU + order qty so an urgent PR-based job gets this PO linked
    # onto it instead of spawning a duplicate job card. Only auto-match when
    # exactly one candidate exists — ambiguous matches fall through to the
    # normal create-new path below.
    pr_matched = []
    if po_number:
        for item in items:
            sku_key = _sku_key(item.get('sku'))
            if not sku_key or sku_key in existing_jobs_by_sku:
                continue
            qty = item.get('quantity')
            item_order_qty = int(qty) if qty is not None else None
            if item_order_qty is None:
                continue
            candidates = list(
                PlanningJob.objects.filter(
                    Q(po_number='') | Q(po_number='-'),
                    sku__iexact=sku_key,
                    order_qty=item_order_qty,
                )
            )
            if len(candidates) == 1:
                matched_job = candidates[0]
                existing_jobs_by_sku[sku_key] = matched_job
                pr_matched.append((matched_job.jc_number, matched_job.pr_reference, po_number))

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

        existing_job = existing_jobs_by_sku.get(sku_key)
        recipe = recipe_map.get(sku_key)
        is_approved = bool(recipe and recipe.master_data_status == 'approved')
        if not bypass_recipe_check and not is_approved:
            if not existing_job:
                missing_recipe_count += 1
            elif _normalize_status(existing_job.status) != 'draft':
                locked_count += 1
                continue

        delivery_date = _parse_iso_date(item.get('delivery_date'))
        intake_plan_date = (
            po_doc.created_at.date()
            if po_doc and getattr(po_doc, 'created_at', None)
            else (po_date or delivery_date)
        )
        jc_plan_date = intake_plan_date if not existing_job else existing_job.plan_date
        qty = item.get('quantity')
        order_qty = int(qty) if qty is not None else None
        unit_cost_val = item.get('unit_cost')
        unit_cost_dec = Decimal(str(unit_cost_val)) if unit_cost_val is not None else None
        jc_number = (
            (item.get('jc_number') or item.get('jc') or item.get('job_card_no') or item.get('jobcardno'))
            or (existing_job.jc_number if existing_job else None)
            or allocate_next_jc_number(jc_plan_date or intake_plan_date)
        )
        is_first_production = bool(
            sku_key
            and sku_key not in existing_any_jobs_skus
            and sku_key not in seen_skus_in_payload
        )

        from .sku_classification import repeat_flag_value_for_po_line

        repeat_flag_value = repeat_flag_value_for_po_line(
            item,
            po_number=po_number,
            po_doc_created_at=getattr(po_doc, 'created_at', None),
            po_doc_id=getattr(po_doc, 'id', None),
            recipe=recipe,
            existing_job=existing_job,
        )
        forward_as_new = repeat_flag_value == 'New'

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

        item_remarks = (item.get('remarks') or '').strip()
        if item_remarks:
            remarks_value = item_remarks
        else:
            remarks_value = (existing_job.remarks if existing_job else '') or (recipe.notes if recipe else '')

        defaults = {
            'po_number': po_number,
            'pr_reference': pr_number or (existing_job.pr_reference if existing_job else ''),
            'sku': sku,
            'job_name': job_name_value,
            'order_qty': order_qty,
            'department': department,
            'destination': delivery_location,
            'delivery_date': delivery_date,
            'unit_cost': unit_cost_value,
            'status': status_value,
            'repeat_flag': repeat_flag_value,
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
            'remarks': remarks_value,
            'purchase_material_origin': purchase_material_origin_value,
            'stock_qty': stock_qty_value,
            'balance_qty': balance_qty_value,
        }
        if po_approval_date:
            defaults['po_approval_date'] = po_approval_date
        if not existing_job:
            if intake_plan_date:
                defaults['plan_date'] = intake_plan_date
            if payload.get('plan_month'):
                defaults['plan_month'] = payload.get('plan_month')
            elif intake_plan_date:
                defaults['plan_month'] = _plan_month_label_from_date(intake_plan_date)
        elif existing_job and not existing_job.plan_date and intake_plan_date:
            defaults['plan_date'] = intake_plan_date
            if payload.get('plan_month'):
                defaults['plan_month'] = payload.get('plan_month')
            elif intake_plan_date:
                defaults['plan_month'] = _plan_month_label_from_date(intake_plan_date)
        if actor and not existing_job:
            defaults['created_by'] = actor

        job_obj, created = PlanningJob.objects.update_or_create(
            jc_number=jc_number,
            defaults=defaults,
        )
        from .sku_classification import sync_plate_making_stage_with_repeat_flag

        sync_plate_making_stage_with_repeat_flag(job_obj, save=True)
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
        'pr_matched': pr_matched,
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

        if existing_job:
            jc_number = existing_job.jc_number
            current_requirement = existing_job.requirement
        else:
            jc_number = allocate_next_jc_number(plan_date)
            current_requirement = ''

        from .sku_classification import repeat_flag_value_for_po_line

        repeat_flag_value = repeat_flag_value_for_po_line(
            target_item,
            po_number=po_number,
            po_doc_created_at=getattr(po_doc, 'created_at', None),
            po_doc_id=getattr(po_doc, 'id', None),
            recipe=recipe,
            existing_job=existing_job,
        )
        forward_as_new = repeat_flag_value == 'New'

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
            'repeat_flag': repeat_flag_value,
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
            'plate_set_no': existing_job.plate_set_no if existing_job else (recipe.plate_set_no or ''),
            'remarks': (existing_job.remarks if existing_job else '') or recipe.notes or '',
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
        from .sku_classification import sync_plate_making_stage_with_repeat_flag

        sync_plate_making_stage_with_repeat_flag(job_obj, save=True)
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

        # Check if any incoming field differs from existing
        has_changes = False
        for field, incoming_value in item.items():
            if field in {'line_no'}:
                continue
            existing_value = existing_item.get(field)
            
            # Normalize to string to compare
            existing_str = str(existing_value or '').strip()
            incoming_str = str(incoming_value or '').strip()
            
            if field in {'quantity'}:
                if _to_int(existing_value) != _to_int(incoming_value):
                    has_changes = True
                    break
            elif field in {'unit_cost', 'net_total', 'subtotal'}:
                if _to_decimal(existing_value) != _to_decimal(incoming_value):
                    has_changes = True
                    break
            elif existing_str != incoming_str:
                has_changes = True
                break

        if not has_changes:
            ignored_lines.append({'sku': sku, 'qty': incoming_qty})
            continue

        # Same SKU but changed qty/fields: treat as correction, not duplicate row.
        for field, value in item.items():
            if field in {'line_no'}:
                continue
            existing_item[field] = value
        updated_skus.append(sku)

    for idx, item in enumerate(merged_items, start=1):
        item['line_no'] = idx

    return merged_items, sorted(set(added_skus)), sorted(set(updated_skus)), ignored_lines



def _collect_pending_sku_rows(po_docs):
    """Build pending SKU rows from PO documents where SKU master is not yet approved."""
    rows = []
    for po_doc in po_docs:
        payload = po_doc.extracted_payload or {}
        items = _po_payload_items(payload)
        if not items:
            continue

        recipe_map = _build_recipe_map(items)
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
        seen_skus = set()
        for item in items:
            sku = (item.get('sku') or '').strip()
            key = _sku_key(sku)
            if not key or key in ignored_skus or key in seen_skus:
                continue
            seen_skus.add(key)

            recipe = recipe_map.get(key)
            is_approved = bool(recipe and recipe.master_data_status == 'approved')
            if is_approved:
                continue

            item = item_map.get(key, {})
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
    if field_name in {'machine_name', 'plate_set_no'} and not str(right or '').strip():
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



def cancel_planning_job(job, actor=None, reason='', reason_code=''):
    """
    Cancel a planning job the customer no longer needs.

    A cancelled job is also archived (is_active=False) so every existing
    active-jobs filter, queue and report drops it without further changes; the
    cancel_* fields are what distinguish a customer cancellation from routine
    housekeeping.
    """
    from django.core.exceptions import ValidationError

    reason = (reason or '').strip()
    if not reason:
        raise ValidationError({'reason': 'A cancellation reason is required.'})

    valid_codes = {code for code, _ in PLANNING_CANCEL_REASON_CHOICES}
    if reason_code not in valid_codes:
        raise ValidationError({'reason_code': 'Select a valid cancellation reason code.'})

    blockers = job.cancellation_blockers()
    if blockers:
        raise ValidationError({'__all__': blockers})

    now = timezone.now()
    with transaction.atomic():
        job.is_cancelled = True
        job.cancel_reason = reason
        job.cancel_reason_code = reason_code
        job.cancelled_by = actor
        job.cancelled_at = now
        job.status = 'cancelled'
        # Mirror into the archive fields so the archived list keeps working.
        job.is_active = False
        job.archive_reason = f'Cancelled: {reason}'
        job.archived_by = actor
        job.archived_at = now
        job.save(update_fields=[
            'is_cancelled',
            'cancel_reason',
            'cancel_reason_code',
            'cancelled_by',
            'cancelled_at',
            'status',
            'is_active',
            'archive_reason',
            'archived_by',
            'archived_at',
            'updated_at',
        ])

        job_card = getattr(job, 'job_card', None)
        if job_card and job_card.is_active:
            job_card.is_active = False
            job_card.save(update_fields=['is_active'])

        # Cancelling a member invalidates the ganged layout — dissolve the whole
        # open merge group so the planner can re-merge the remaining SKUs cleanly.
        merge_item = job.active_merge_item
        if merge_item:
            group = merge_item.merge_group
            group.status = 'cancelled'
            group.cancelled_at = now
            group.save(update_fields=['status', 'cancelled_at'])

    return job


def request_job_cancellation(job, actor=None, reason='', reason_code=''):
    """Raise a PM approval request to cancel a job that is already released."""
    from django.core.exceptions import ValidationError

    reason = (reason or '').strip()
    if not reason:
        raise ValidationError({'reason': 'A cancellation reason is required.'})

    valid_codes = {code for code, _ in PLANNING_CANCEL_REASON_CHOICES}
    if reason_code not in valid_codes:
        raise ValidationError({'reason_code': 'Select a valid cancellation reason code.'})

    if job.is_cancelled:
        raise ValidationError({'__all__': ['This job is already cancelled.']})

    existing = JobCardChangeRequest.objects.filter(
        planning_job=job,
        request_type=JOB_CANCEL_REQUEST_TYPE,
        status='pending',
    ).first()
    if existing:
        raise ValidationError({'__all__': ['A cancellation request is already pending for this job.']})

    return JobCardChangeRequest.objects.create(
        planning_job=job,
        request_type=JOB_CANCEL_REQUEST_TYPE,
        # The code is stored inline so approval can replay the exact cancellation.
        # JobCardChangeRequest.cancel_reason_code/cancel_reason_text parse it back.
        reason=f'[{reason_code}] {reason}',
        requested_by=actor,
    )


def approve_job_cancellation(change_request, actor=None):
    """Apply a PM-approved cancellation request to its planning job."""
    return cancel_planning_job(
        change_request.planning_job,
        actor=actor,
        reason=change_request.cancel_reason_text,
        reason_code=change_request.cancel_reason_code,
    )
