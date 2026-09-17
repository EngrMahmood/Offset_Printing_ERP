"""Resolve a planning.SkuRecipe into a normalized 'spec signature' dict.

The signature is used for three things: (1) matching a SKU against
BomTemplateCondition rows, (2) the namespace formulas evaluate against, and
(3) the snapshot stored on SkuBom.spec_snapshot to detect drift later.

Extending the signature (a new match key) is meant to be a *data* operation —
add a SpecAttribute row picking one of ``available_sku_fields()`` — not a code
change. The only things that need code are genuinely computed values, and
those go in DERIVED_RESOLVERS.
"""
from __future__ import annotations

import re
from decimal import Decimal, InvalidOperation

# Fields/properties on planning.SkuRecipe that may be used as a SpecAttribute
# source_field. Kept as an explicit allow-list (rather than raw getattr) so a
# SpecAttribute can never be pointed at an unrelated or sensitive attribute.
_ALLOWED_SKU_FIELDS = (
    'sku', 'job_name', 'material', 'color_spec', 'application', 'product_type',
    'machine_name', 'job_process_type', 'print_passes', 'plate_set_no',
    'size_w_mm', 'size_h_mm', 'ups', 'print_sheet_size', 'purchase_sheet_size',
    'purchase_sheet_ups', 'default_unit_cost', 'daily_demand', 'awc_no',
    'die_cutting', 'lamination_front_and_back', 'legacy_produced', 'is_active',
)


def available_sku_fields():
    """Concrete SkuRecipe fields/properties usable as a SpecAttribute.source_field.

    Cross-checked against the live model so a stale allow-list entry (a field
    renamed or removed upstream) never silently passes validation.
    """
    from planning.models import SkuRecipe

    model_field_names = {f.name for f in SkuRecipe._meta.get_fields()}
    return tuple(name for name in _ALLOWED_SKU_FIELDS if name in model_field_names)


_GSM_RE = re.compile(r'(\d{2,4})\s*gsm', re.IGNORECASE)
_COLOR_COUNT_RE = re.compile(r'\d+')


def _parse_gsm(sku_recipe):
    """Grammage is embedded in the free-text material name (no gsm field exists)."""
    match = _GSM_RE.search(sku_recipe.material or '')
    if not match:
        return None
    try:
        return Decimal(match.group(1))
    except InvalidOperation:
        return None


def _parse_total_colors(sku_recipe):
    """color_spec holds forms like '4' or '1+1' — sum the numbers."""
    text = (sku_recipe.color_spec or '').strip()
    if not text:
        return None
    numbers = _COLOR_COUNT_RE.findall(text)
    if not numbers:
        return None
    return Decimal(sum(int(n) for n in numbers))


def _mm_to_m(value):
    if value is None:
        return None
    return Decimal(value) / Decimal(1000)


def _piece_area_sqm(sku_recipe):
    if sku_recipe.size_w_mm is None or sku_recipe.size_h_mm is None:
        return None
    return _mm_to_m(sku_recipe.size_w_mm) * _mm_to_m(sku_recipe.size_h_mm)


def _sheet_area_sqm(size_str):
    """Parse a 'W*H' or 'WxH' sheet-size string (in inches) into sqm."""
    if not size_str:
        return None
    match = re.match(r'^\s*(\d+(?:\.\d+)?)\s*[xX×*]\s*(\d+(?:\.\d+)?)\s*$', str(size_str).strip())
    if not match:
        return None
    w_in, h_in = Decimal(match.group(1)), Decimal(match.group(2))
    sqm_per_sqin = Decimal('0.00064516')
    return w_in * h_in * sqm_per_sqin


def _print_sheet_area_sqm(sku_recipe):
    return _sheet_area_sqm(sku_recipe.print_sheet_size)


def _purchase_sheet_area_sqm(sku_recipe):
    return _sheet_area_sqm(sku_recipe.purchase_sheet_size)


def _sheets_per_1000_pcs(sku_recipe):
    ups = sku_recipe.ups or 0
    if not ups:
        return None
    return (Decimal(1000) / Decimal(ups)).quantize(Decimal('0.0001'))


# Registry of computed values not present as a plain SkuRecipe field.
# Adding a *new* one is the one legitimate case that needs code (nothing here
# can be expressed purely as data), but it stays a single small function.
DERIVED_RESOLVERS = {
    'gsm': _parse_gsm,
    'total_colors': _parse_total_colors,
    'piece_area_sqm': _piece_area_sqm,
    'print_sheet_area_sqm': _print_sheet_area_sqm,
    'purchase_sheet_area_sqm': _purchase_sheet_area_sqm,
    'sheets_per_1000_pcs': _sheets_per_1000_pcs,
}


def _apply_normalizer(normalizer, value):
    if value is None:
        return None
    if normalizer == 'none':
        return value
    if normalizer == 'lower_trim':
        return ' '.join(str(value).strip().split()).lower()
    if normalizer == 'material_name':
        from supply_chain.models import normalize_material_name
        return normalize_material_name(value)
    if normalizer == 'sheet_size':
        from supply_chain.models import normalize_purchase_sheet_size
        return normalize_purchase_sheet_size(value)
    if normalizer == 'yes_no':
        from planning.services import normalize_die_cutting
        return normalize_die_cutting(value)
    return value


def _json_safe(value):
    if isinstance(value, Decimal):
        return str(value)
    return value


def resolve_attribute_value(attribute, sku_recipe):
    """Resolve one SpecAttribute's value for a given SkuRecipe (raw, pre-JSON-safe)."""
    if attribute.source_kind == 'sku_field':
        if attribute.source_field not in available_sku_fields():
            return None
        value = getattr(sku_recipe, attribute.source_field, None)
    elif attribute.source_kind == 'regex_capture':
        if not attribute.pattern:
            return None
        match = re.search(attribute.pattern, sku_recipe.sku or '')
        value = match.group(attribute.capture_group) if match else None
    elif attribute.source_kind == 'derived':
        resolver = DERIVED_RESOLVERS.get(attribute.derived_key)
        value = resolver(sku_recipe) if resolver else None
    else:
        value = None

    value = _apply_normalizer(attribute.normalizer, value)

    if value is not None and attribute.value_type == 'number' and not isinstance(value, Decimal):
        try:
            value = Decimal(str(value))
        except InvalidOperation:
            value = None
    return value


def resolve_sku_spec(sku_recipe):
    """Full normalized spec signature for a SKU: {attribute_code: value}.

    Includes every active SpecAttribute plus the raw derived resolvers keyed
    by their own name (so formulas can reference e.g. `piece_area_sqm` even if
    no SpecAttribute row happens to expose it as a match key).
    """
    from .models import SpecAttribute

    result = {}
    for attribute in SpecAttribute.objects.filter(is_active=True):
        result[attribute.code] = _json_safe(resolve_attribute_value(attribute, sku_recipe))

    for key, resolver in DERIVED_RESOLVERS.items():
        result.setdefault(key, _json_safe(resolver(sku_recipe)))

    for field_name in available_sku_fields():
        result.setdefault(f'sku.{field_name}', _json_safe(getattr(sku_recipe, field_name, None)))

    return result


def spec_namespace_for_formula(spec):
    """Convert a resolved spec dict (JSON-safe strings) into a Decimal-friendly
    namespace for bom.formula.evaluate(), dropping non-numeric values that a
    formula could never legitimately reference by name anyway."""
    from decimal import Decimal as D

    namespace = {}
    for key, value in spec.items():
        if value is None:
            continue
        if isinstance(value, bool):
            namespace[key] = D(1) if value else D(0)
            continue
        if isinstance(value, (int, float)):
            namespace[key] = D(str(value))
            continue
        if isinstance(value, str):
            try:
                namespace[key] = D(value)
            except InvalidOperation:
                continue
    return namespace
