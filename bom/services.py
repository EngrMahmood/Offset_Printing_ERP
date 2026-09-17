"""BOM generation and workflow services.

generate_bom_for_sku() is the core operation: resolve a SKU's spec, find the
best matching approved BomTemplate, evaluate every line's item selection and
quantity formula, and save the result as a new draft SkuBom version. Nothing
here ever mutates an approved BOM in place — regeneration always creates a
new version and supersedes is_current.
"""
from __future__ import annotations

from decimal import Decimal, ROUND_HALF_UP

from django.db import transaction
from django.utils import timezone

from .formula import FormulaError, evaluate
from .matching import best_template
from .models import BomTemplateLine, RawItem, SkuBom, SkuBomLine
from .spec import resolve_sku_spec, spec_namespace_for_formula

QUANT = Decimal('0.00000001')
COST_QUANT = Decimal('0.0001')


class BomGenerationError(Exception):
    pass


def _resolve_fixed_item(line, warnings):
    if line.raw_item_id and line.raw_item.is_active:
        return line.raw_item
    warnings.append(f'Line "{line.line_code}": fixed raw item is missing or inactive.')
    return None


def _resolve_by_selector_item(line, namespace, warnings):
    candidates = RawItem.objects.filter(is_active=True, item_role=line.item_role)
    for expr_key, expr in (line.selector or {}).items():
        try:
            target_value = evaluate(expr, namespace)
        except FormulaError as exc:
            warnings.append(f'Line "{line.line_code}": selector formula for "{expr_key}" failed: {exc}')
            return None
        matched = []
        for candidate in candidates:
            attr_value = candidate.attributes.get(expr_key)
            if attr_value is None:
                continue
            try:
                if Decimal(str(attr_value)) == target_value:
                    matched.append(candidate)
            except Exception:
                if str(attr_value) == str(target_value):
                    matched.append(candidate)
        candidates = candidates.filter(pk__in=[c.pk for c in matched]) if matched else RawItem.objects.none()
        if not matched:
            break
    result = candidates.first() if hasattr(candidates, 'first') else (candidates[0] if candidates else None)
    if result is None:
        warnings.append(f'Line "{line.line_code}": no raw item matched role "{line.item_role}" with selector {line.selector}.')
    return result


def _resolve_from_sku_substrate_item(line, sku_recipe, warnings):
    from supply_chain.raw_material_sku import resolve_raw_material_sku
    from core.jobcard_service import _resolve_by_name
    from core.models import Material

    material = _resolve_by_name(Material, sku_recipe.material)
    if material is None:
        warnings.append(f'Line "{line.line_code}": SKU material "{sku_recipe.material}" not found in master data.')
        return None
    rm_sku = resolve_raw_material_sku(material, sku_recipe.purchase_sheet_size)
    if rm_sku is None:
        warnings.append(
            f'Line "{line.line_code}": no RawMaterialSku for material "{sku_recipe.material}" '
            f'/ sheet size "{sku_recipe.purchase_sheet_size}".'
        )
        return None
    raw_item = RawItem.objects.filter(raw_material_sku=rm_sku, is_active=True).first()
    if raw_item is None:
        warnings.append(f'Line "{line.line_code}": RawMaterialSku "{rm_sku.sku}" has no bridged RawItem yet.')
        return None
    return raw_item


def _resolve_line_item(line, sku_recipe, namespace, warnings):
    if line.resolution_mode == 'fixed':
        return _resolve_fixed_item(line, warnings)
    if line.resolution_mode == 'by_selector':
        return _resolve_by_selector_item(line, namespace, warnings)
    if line.resolution_mode == 'from_sku_substrate':
        return _resolve_from_sku_substrate_item(line, sku_recipe, warnings)
    warnings.append(f'Line "{line.line_code}": unknown resolution mode "{line.resolution_mode}".')
    return None


def _evaluate_simple_line(line, raw_item):
    """entry_mode='simple': a plain per-unit quantity and wastage, typed directly
    against this exact spec combination — no formula, no driver dependency,
    matching the carton department's manual Excel workflow."""
    qty_per_unit = (line.fixed_qty or Decimal(0)).quantize(QUANT, rounding=ROUND_HALF_UP)
    wastage = line.fixed_wastage_percent or Decimal(0)
    gross_qty = (qty_per_unit * (Decimal(1) + wastage)).quantize(QUANT, rounding=ROUND_HALF_UP)
    resolved_formula = f'fixed: {qty_per_unit} (+{wastage * 100}% wastage)'
    return raw_item, qty_per_unit, wastage, gross_qty, resolved_formula


def _evaluate_line(line, sku_recipe, namespace, line_qty_by_code, warnings):
    """Returns (raw_item, qty_per_unit, wastage_percent, gross_qty, resolved_formula) or None on failure."""
    raw_item = _resolve_line_item(line, sku_recipe, namespace, warnings)
    if raw_item is None:
        return None

    if line.entry_mode == 'simple':
        return _evaluate_simple_line(line, raw_item)

    eval_namespace = dict(namespace)
    if line.driver_line_id and line.driver_line.line_code in line_qty_by_code:
        eval_namespace['driver_qty'] = line_qty_by_code[line.driver_line.line_code]

    scale = Decimal(1)
    if line.scale_by_colors:
        colors = namespace.get('total_colors') or namespace.get('sku.total_colors') or Decimal(1)
        scale *= colors
    if line.scale_by_passes:
        passes = namespace.get('print_passes') or namespace.get('sku.print_passes') or Decimal(1)
        scale *= passes

    try:
        qty_per_unit = evaluate(line.qty_expression, eval_namespace) * scale
    except FormulaError as exc:
        warnings.append(f'Line "{line.line_code}": quantity formula failed: {exc}')
        return None

    try:
        wastage = evaluate(line.wastage_expression or '0', eval_namespace)
    except FormulaError as exc:
        warnings.append(f'Line "{line.line_code}": wastage formula failed: {exc}')
        wastage = Decimal(0)

    qty_per_unit = qty_per_unit.quantize(QUANT, rounding=ROUND_HALF_UP)
    gross_qty = (qty_per_unit * (Decimal(1) + wastage)).quantize(QUANT, rounding=ROUND_HALF_UP)
    resolved_formula = f'{line.qty_expression} [x scale={scale}]' if scale != 1 else line.qty_expression
    return raw_item, qty_per_unit, wastage, gross_qty, resolved_formula


@transaction.atomic
def generate_bom_for_sku(sku_recipe, user=None, mode='auto'):
    """Generate (and save as draft) a new SkuBom version for sku_recipe.

    Returns the new SkuBom, or None if no approved template matches (recorded
    nowhere persistent in that case — callers/signals decide whether that's
    worth surfacing).
    """
    match = best_template(sku_recipe, production_line='OFFSET')
    if match.template is None:
        return None

    spec = resolve_sku_spec(sku_recipe)
    namespace = spec_namespace_for_formula(spec)

    warnings = []
    if match.is_ambiguous:
        warnings.append(
            f'Multiple templates tied for best match on priority/specificity/recency; '
            f'"{match.template.code}" was picked arbitrarily among: '
            f'{", ".join(t.code for t in match.candidates)}.'
        )

    previous = SkuBom.objects.filter(sku_recipe=sku_recipe).order_by('-version').first()
    next_version = (previous.version + 1) if previous else 1

    bom = SkuBom.objects.create(
        sku_recipe=sku_recipe,
        version=next_version,
        is_current=True,
        source_template=match.template,
        generation_mode=mode,
        spec_snapshot=spec,
        status='draft',
        created_by=user,
    )

    line_qty_by_code = {}
    total_cost = Decimal(0)
    for sequence, template_line in enumerate(match.template.ordered_lines()):
        result = _evaluate_line(template_line, sku_recipe, namespace, line_qty_by_code, warnings)
        if result is None:
            if not template_line.is_optional:
                warnings.append(f'Line "{template_line.line_code}" could not be generated and is not optional.')
            continue
        raw_item, qty_per_unit, wastage, gross_qty, resolved_formula = result
        line_qty_by_code[template_line.line_code] = gross_qty
        line_cost = (gross_qty * raw_item.unit_cost).quantize(COST_QUANT, rounding=ROUND_HALF_UP)
        total_cost += line_cost

        SkuBomLine.objects.create(
            bom=bom,
            sequence=sequence,
            line_code=template_line.line_code,
            process_step=template_line.process_step,
            raw_item=raw_item,
            uom_code=raw_item.uom.code,
            qty_per_unit=qty_per_unit,
            wastage_percent=wastage,
            gross_qty=gross_qty,
            unit_cost_snapshot=raw_item.unit_cost,
            line_cost=line_cost,
            source_template_line=template_line,
            resolved_formula=resolved_formula,
        )

    bom.total_material_cost = total_cost.quantize(COST_QUANT, rounding=ROUND_HALF_UP)
    bom.cost_computed_at = timezone.now()
    bom.warnings = warnings
    bom.save(update_fields=['total_material_cost', 'cost_computed_at', 'warnings'])

    if previous:
        SkuBom.objects.filter(pk=previous.pk).update(is_current=False)

    return bom


def recompute_bom_cost(bom):
    total = Decimal(0)
    lines = list(bom.lines.select_related('raw_item'))
    for line in lines:
        line.unit_cost_snapshot = line.raw_item.unit_cost
        line.line_cost = (line.gross_qty * line.raw_item.unit_cost).quantize(COST_QUANT, rounding=ROUND_HALF_UP)
        total += line.line_cost
    SkuBomLine.objects.bulk_update(lines, ['unit_cost_snapshot', 'line_cost'])
    bom.total_material_cost = total.quantize(COST_QUANT, rounding=ROUND_HALF_UP)
    bom.cost_computed_at = timezone.now()
    bom.save(update_fields=['total_material_cost', 'cost_computed_at'])
    return bom


def _notify(bom, roles, event_type, title, message):
    from core.notifications import notify_roles
    from django.urls import reverse
    try:
        link = reverse('bom:bom_detail', args=[bom.pk])
    except Exception:
        link = ''
    notify_roles(roles, event_type=event_type, title=title, message=message, link=link,
                 entity_type='skubom', entity_id=bom.pk)


def submit_bom(bom, user):
    bom.status = 'pending_review'
    bom.save(update_fields=['status'])
    _notify(bom, ['qc'], 'bom.submitted', f'BOM submitted: {bom.sku_recipe.sku}', 'A BOM was submitted for review.')


def review_bom(bom, user):
    bom.status = 'reviewed'
    bom.reviewed_by = user
    bom.reviewed_at = timezone.now()
    bom.save(update_fields=['status', 'reviewed_by', 'reviewed_at'])
    _notify(bom, ['admin', 'manager'], 'bom.reviewed', f'BOM ready for approval: {bom.sku_recipe.sku}', '')


def approve_bom(bom, user):
    bom.status = 'approved'
    bom.approved_by = user
    bom.approved_at = timezone.now()
    bom.save(update_fields=['status', 'approved_by', 'approved_at'])


def reject_bom(bom, user, comment=''):
    bom.status = 'draft'
    bom.rejection_comment = comment
    bom.save(update_fields=['status', 'rejection_comment'])


def soft_delete_bom(bom, user):
    bom.is_active = False
    bom.save(update_fields=['is_active'])
