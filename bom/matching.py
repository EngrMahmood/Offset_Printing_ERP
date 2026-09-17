"""Match a resolved SKU spec against BomTemplateCondition rows.

All matching logic lives in data (BomTemplateCondition rows) — this module
only implements the fixed set of operators. Adding a new match key is adding
a SpecAttribute + condition rows, never touching this file.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation

from .models import BomTemplate
from .spec import resolve_attribute_value


def _as_decimal(value):
    try:
        return Decimal(str(value))
    except (InvalidOperation, TypeError, ValueError):
        return None


def _as_text(value):
    if value is None:
        return ''
    return str(value).strip().lower()


def evaluate_condition(condition, sku_recipe_or_spec) -> bool:
    """True if the SKU's resolved value for condition.attribute satisfies the condition.

    Accepts either a ``planning.SkuRecipe`` instance (resolved live via
    ``resolve_attribute_value``, exactly as before) or an already-resolved spec
    dict keyed by attribute code (e.g. built from wizard selections that have
    no backing SkuRecipe yet).
    """
    if condition.operator == 'any':
        return True

    if isinstance(sku_recipe_or_spec, dict):
        value = sku_recipe_or_spec.get(condition.attribute.code)
    else:
        value = resolve_attribute_value(condition.attribute, sku_recipe_or_spec)

    if condition.operator == 'eq':
        return _as_text(value) == _as_text(condition.value_text)
    if condition.operator == 'neq':
        return _as_text(value) != _as_text(condition.value_text)
    if condition.operator == 'contains':
        return _as_text(condition.value_text) in _as_text(value)
    if condition.operator == 'in':
        return _as_text(value) in {_as_text(v) for v in (condition.value_list or [])}
    if condition.operator == 'not_in':
        return _as_text(value) not in {_as_text(v) for v in (condition.value_list or [])}

    if condition.operator in ('gt', 'gte', 'lt', 'lte', 'between'):
        left = _as_decimal(value)
        if left is None:
            return False
        if condition.operator == 'gt':
            right = _as_decimal(condition.value_text)
            return right is not None and left > right
        if condition.operator == 'gte':
            right = _as_decimal(condition.value_text)
            return right is not None and left >= right
        if condition.operator == 'lt':
            right = _as_decimal(condition.value_text)
            return right is not None and left < right
        if condition.operator == 'lte':
            right = _as_decimal(condition.value_text)
            return right is not None and left <= right
        if condition.operator == 'between':
            lo, hi = _as_decimal(condition.value_min), _as_decimal(condition.value_max)
            return lo is not None and hi is not None and lo <= left <= hi

    return False


def template_matches(template, sku_recipe_or_spec) -> bool:
    conditions = list(template.conditions.select_related('attribute').all())
    return all(evaluate_condition(c, sku_recipe_or_spec) for c in conditions)


def find_matching_templates(sku_recipe_or_spec, production_line='OFFSET'):
    """``sku_recipe_or_spec`` may be a SkuRecipe instance or a resolved spec dict
    (e.g. from the Recipe Wizard, where no SkuRecipe exists yet)."""
    templates = (
        BomTemplate.objects.filter(is_active=True, status='approved', production_line=production_line)
        .prefetch_related('conditions__attribute')
    )
    return [t for t in templates if template_matches(t, sku_recipe_or_spec)]


@dataclass
class MatchResult:
    template: BomTemplate | None
    candidates: list = field(default_factory=list)
    is_ambiguous: bool = False


def best_template(sku_recipe_or_spec, production_line='OFFSET') -> MatchResult:
    """Highest priority wins; ties broken by most conditions (specificity), then
    most recently updated. A true tie across all three is reported as ambiguous
    rather than silently resolved by, e.g., pk order.

    ``sku_recipe_or_spec`` may be a SkuRecipe instance or a resolved spec dict.
    """
    candidates = find_matching_templates(sku_recipe_or_spec, production_line=production_line)
    if not candidates:
        return MatchResult(template=None, candidates=[])

    def sort_key(t):
        return (-t.priority, -t.conditions.count(), -(t.updated_at.timestamp() if t.updated_at else 0))

    ranked = sorted(candidates, key=sort_key)
    top = ranked[0]
    top_key = sort_key(top)
    tied = [t for t in ranked if sort_key(t) == top_key]
    return MatchResult(template=top, candidates=candidates, is_ambiguous=len(tied) > 1)
