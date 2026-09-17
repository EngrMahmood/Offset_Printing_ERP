from django.contrib import messages
from django.shortcuts import get_object_or_404, redirect, render
from django.urls import reverse

from planning.models import SkuRecipe

from .decorators import bom_approve_required, bom_manage_required, bom_required
from .excel_io import (
    commit_import_raw_items, download_raw_items_template, download_sku_bom_template,
    export_raw_items, export_sku_bom,
)
from .forms import (
    BomTemplateConditionFormSet, BomTemplateForm, BomTemplateLineFormSet, RawItemForm,
    RawItemImportForm, RaiseItemRequestForm, SpecAttributeForm, TestMatchForm,
)
from .matching import best_template, find_matching_templates
from .models import BomTemplate, RawItem, SkuBom, SpecAttribute
from .requirements import explode_requirements, raise_item_requests_from_shortfalls
from .services import approve_bom, generate_bom_for_sku, reject_bom, review_bom, submit_bom
from .spec import resolve_attribute_value, resolve_sku_spec


@bom_required
def dashboard(request):
    context = {
        'raw_item_count': RawItem.objects.filter(is_active=True).count(),
        'template_count': BomTemplate.objects.filter(is_active=True).count(),
        'draft_bom_count': SkuBom.objects.filter(is_current=True, status='draft').count(),
        'approved_bom_count': SkuBom.objects.filter(is_current=True, status='approved').count(),
    }
    return render(request, 'bom/dashboard.html', context)


# ---------------------------------------------------------------------------
# Raw items
# ---------------------------------------------------------------------------

@bom_required
def raw_item_list(request):
    items = RawItem.objects.filter(is_active=True).select_related('category', 'uom', 'default_vendor')
    category = request.GET.get('category')
    if category:
        items = items.filter(category_id=category)
    q = request.GET.get('q')
    if q:
        items = items.filter(name__icontains=q) | items.filter(item_code__icontains=q)
    if request.GET.get('export') == 'xlsx':
        return export_raw_items(items)
    from .models import ItemCategory
    return render(request, 'bom/raw_item_list.html', {
        'items': items, 'categories': ItemCategory.objects.filter(is_active=True),
    })


@bom_manage_required
def raw_item_form(request, pk=None):
    instance = get_object_or_404(RawItem, pk=pk) if pk else None
    if request.method == 'POST':
        form = RawItemForm(request.POST, instance=instance)
        if form.is_valid():
            item = form.save(commit=False)
            if not instance:
                item.created_by = request.user
            item.save()
            messages.success(request, f'Raw item {item.item_code} saved.')
            return redirect('bom:raw_item_list')
    else:
        form = RawItemForm(instance=instance)
    return render(request, 'bom/raw_item_form.html', {'form': form, 'instance': instance})


@bom_manage_required
def raw_item_import(request):
    if request.method == 'POST':
        form = RawItemImportForm(request.POST, request.FILES)
        if form.is_valid():
            result, errors = commit_import_raw_items(form.cleaned_data['upload'], user=request.user)
            if errors:
                for e in errors:
                    messages.error(request, e)
            else:
                messages.success(request, f"Imported: {result['created']} created, {result['updated']} updated.")
                return redirect('bom:raw_item_list')
    else:
        form = RawItemImportForm()
    return render(request, 'bom/raw_item_import.html', {'form': form})


@bom_required
def raw_item_template_download(request):
    return download_raw_items_template()


@bom_required
def sku_bom_template_download(request):
    return download_sku_bom_template()


# ---------------------------------------------------------------------------
# Spec attributes
# ---------------------------------------------------------------------------

@bom_required
def spec_attribute_list(request):
    attributes = SpecAttribute.objects.all()
    return render(request, 'bom/spec_attribute_list.html', {'attributes': attributes})


@bom_manage_required
def spec_attribute_form(request, pk=None):
    from django.utils.http import url_has_allowed_host_and_scheme

    instance = get_object_or_404(SpecAttribute, pk=pk) if pk else None
    next_url = request.POST.get('next') or request.GET.get('next') or ''
    if next_url and not url_has_allowed_host_and_scheme(next_url, allowed_hosts={request.get_host()}, require_https=request.is_secure()):
        next_url = ''
    if request.method == 'POST':
        form = SpecAttributeForm(request.POST, instance=instance)
        if form.is_valid():
            attribute = form.save(commit=False)
            attribute.full_clean()
            attribute.save()
            messages.success(request, f'Spec attribute "{attribute.label}" saved.')
            return redirect(next_url or 'bom:spec_attribute_list')
    else:
        form = SpecAttributeForm(instance=instance)
    return render(request, 'bom/spec_attribute_form.html', {'form': form, 'instance': instance, 'next_url': next_url})


# ---------------------------------------------------------------------------
# Templates
# ---------------------------------------------------------------------------

@bom_required
def template_list(request):
    templates = BomTemplate.objects.filter(is_active=True)
    return render(request, 'bom/template_list.html', {'templates': templates})


@bom_required
def template_detail(request, pk):
    template = get_object_or_404(BomTemplate, pk=pk)
    return render(request, 'bom/template_detail.html', {
        'template': template,
        'conditions': template.conditions.select_related('attribute'),
        'lines': template.ordered_lines(),
    })


@bom_manage_required
def template_form(request, pk=None):
    instance = get_object_or_404(BomTemplate, pk=pk) if pk else None
    if request.method == 'POST':
        form = BomTemplateForm(request.POST, instance=instance)
        condition_formset = BomTemplateConditionFormSet(request.POST, instance=instance)
        line_formset = BomTemplateLineFormSet(request.POST, instance=instance)
        if form.is_valid() and condition_formset.is_valid() and line_formset.is_valid():
            template = form.save(commit=False)
            if not instance:
                template.created_by = request.user
            template.save()
            condition_formset.instance = template
            condition_formset.save()
            line_formset.instance = template
            line_formset.save()
            try:
                template.full_clean()
            except Exception as exc:
                messages.warning(request, f'Saved, but validation warning: {exc}')
            messages.success(request, f'Template {template.code} saved.')
            return redirect('bom:template_detail', pk=template.pk)
    else:
        form = BomTemplateForm(instance=instance)
        condition_initial = []
        if not instance:
            for attr_pk in request.GET.getlist('prefill_attr'):
                value = request.GET.get(f'prefill_val_{attr_pk}', '').strip()
                if value:
                    condition_initial.append({'attribute': attr_pk, 'operator': 'eq', 'value_text': value})
        condition_formset = BomTemplateConditionFormSet(instance=instance, initial=condition_initial)
        if condition_initial:
            condition_formset.extra = max(condition_formset.extra, len(condition_initial))
        line_formset = BomTemplateLineFormSet(instance=instance)
    return render(request, 'bom/template_form.html', {
        'form': form, 'condition_formset': condition_formset, 'line_formset': line_formset, 'instance': instance,
    })


@bom_approve_required
def template_submit(request, pk):
    template = get_object_or_404(BomTemplate, pk=pk)
    template.status = 'pending_review'
    template.save(update_fields=['status'])
    messages.success(request, 'Template submitted for review.')
    return redirect('bom:template_detail', pk=pk)


@bom_approve_required
def template_approve(request, pk):
    template = get_object_or_404(BomTemplate, pk=pk)
    template.status = 'approved'
    template.approved_by = request.user
    from django.utils import timezone
    template.approved_at = timezone.now()
    template.save(update_fields=['status', 'approved_by', 'approved_at'])
    messages.success(request, 'Template approved.')
    return redirect('bom:template_detail', pk=pk)


def _distinct_attribute_values(attribute, prior_selections, steps_by_code, sample_size=1000):
    """Distinct real values seen for `attribute` across recent active SkuRecipes,
    restricted to those whose already-selected prior-step attributes match.
    Cheap per-attribute resolution (no full resolve_sku_spec) so a wizard step
    with several prior selections stays a single query."""
    recipes = SkuRecipe.objects.filter(is_active=True).order_by('-updated_at')[:sample_size]
    values = set()
    for recipe in recipes:
        matches_prior = True
        for code, selected_value in prior_selections.items():
            prior_attr = steps_by_code.get(code)
            if prior_attr is None:
                continue
            resolved = resolve_attribute_value(prior_attr, recipe)
            if str(resolved if resolved is not None else '').strip().lower() != str(selected_value).strip().lower():
                matches_prior = False
                break
        if not matches_prior:
            continue
        resolved = resolve_attribute_value(attribute, recipe)
        if resolved not in (None, ''):
            values.add(str(resolved))
    return sorted(values)


@bom_required
def recipe_wizard(request):
    """Cascading attribute picker: product type -> material -> size -> color ->
    application -> die (or whatever SpecAttribute rows are configured, in their
    sort_order) -> either the recipe already built for that combination, or a
    one-click path to build one. Steps come entirely from SpecAttribute data;
    adding/reordering a step is a data change, never a code change."""
    steps = list(SpecAttribute.objects.filter(is_active=True, is_wizard_step=True).order_by('sort_order'))
    steps_by_code = {a.code: a for a in steps}

    selections = {}
    for attr in steps:
        value = request.GET.get(attr.code, '').strip()
        if not value:
            break
        selections[attr.code] = value

    current_step = next((a for a in steps if a.code not in selections), None)

    from urllib.parse import urlencode

    def qs(extra=None, without=None):
        params = {code: value for code, value in selections.items() if code != without}
        if extra:
            params.update(extra)
        return urlencode(params)

    selected_pairs = [
        {'attribute': attr, 'value': selections[attr.code], 'remove_qs': qs(without=attr.code)}
        for attr in steps if attr.code in selections
    ]

    option_links = []
    if current_step is not None:
        for value in _distinct_attribute_values(current_step, selections, steps_by_code):
            option_links.append({'value': value, 'qs': qs(extra={current_step.code: value})})

    spec = matches = result = None
    condition_prefill_url = ''
    if steps and current_step is None:
        spec = dict(selections)
        matches = find_matching_templates(spec)
        result = best_template(spec)
        if result.template is None:
            params = {'prefill_attr': [a.pk for a in steps]}
            query = urlencode(params, doseq=True) + '&' + urlencode({f'prefill_val_{a.pk}': selections[a.code] for a in steps})
            condition_prefill_url = f"{reverse('bom:template_create')}?{query}"

    return render(request, 'bom/recipe_wizard.html', {
        'steps': steps,
        'selections': selections,
        'selected_pairs': selected_pairs,
        'current_step': current_step,
        'current_step_qs': qs() if current_step is not None else '',
        'option_links': option_links,
        'spec': spec,
        'matches': matches,
        'result': result,
        'condition_prefill_url': condition_prefill_url,
    })


@bom_required
def test_match(request):
    result = None
    matches = []
    spec = None
    if request.method == 'POST':
        form = TestMatchForm(request.POST)
        if form.is_valid():
            sku_code = form.cleaned_data['sku'].strip()
            recipe = SkuRecipe.objects.filter(sku__iexact=sku_code).order_by('-updated_at').first()
            if recipe is None:
                messages.error(request, f'No SKU found matching "{sku_code}".')
            else:
                spec = resolve_sku_spec(recipe)
                matches = find_matching_templates(recipe)
                result = best_template(recipe)
    else:
        form = TestMatchForm()
    return render(request, 'bom/test_match.html', {
        'form': form, 'result': result, 'matches': matches, 'spec': spec,
    })


# ---------------------------------------------------------------------------
# Generated BOMs
# ---------------------------------------------------------------------------

@bom_required
def bom_list(request):
    boms = SkuBom.objects.filter(is_current=True, is_active=True).select_related('sku_recipe', 'source_template')
    status = request.GET.get('status')
    if status:
        boms = boms.filter(status=status)
    if request.GET.get('export') == 'xlsx':
        return export_sku_bom(boms)
    return render(request, 'bom/bom_list.html', {'boms': boms})


@bom_required
def bom_detail(request, pk):
    bom = get_object_or_404(SkuBom, pk=pk)
    return render(request, 'bom/bom_detail.html', {
        'bom': bom, 'lines': bom.lines.select_related('raw_item', 'process_step'),
    })


@bom_manage_required
def bom_regenerate(request, pk):
    bom = get_object_or_404(SkuBom, pk=pk)
    new_bom = generate_bom_for_sku(bom.sku_recipe, user=request.user, mode='manual')
    if new_bom is None:
        messages.error(request, 'No approved template matches this SKU.')
        return redirect('bom:bom_detail', pk=pk)
    messages.success(request, f'Regenerated as version {new_bom.version}.')
    return redirect('bom:bom_detail', pk=new_bom.pk)


@bom_required
def bom_submit(request, pk):
    bom = get_object_or_404(SkuBom, pk=pk)
    submit_bom(bom, request.user)
    messages.success(request, 'BOM submitted for review.')
    return redirect('bom:bom_detail', pk=pk)


@bom_approve_required
def bom_review(request, pk):
    bom = get_object_or_404(SkuBom, pk=pk)
    review_bom(bom, request.user)
    messages.success(request, 'BOM marked reviewed.')
    return redirect('bom:bom_detail', pk=pk)


@bom_approve_required
def bom_approve(request, pk):
    bom = get_object_or_404(SkuBom, pk=pk)
    approve_bom(bom, request.user)
    messages.success(request, 'BOM approved.')
    return redirect('bom:bom_detail', pk=pk)


@bom_approve_required
def bom_reject(request, pk):
    bom = get_object_or_404(SkuBom, pk=pk)
    comment = request.POST.get('comment', '')
    reject_bom(bom, request.user, comment=comment)
    messages.warning(request, 'BOM sent back to draft.')
    return redirect('bom:bom_detail', pk=pk)


# ---------------------------------------------------------------------------
# Requirements / purchasing
# ---------------------------------------------------------------------------

@bom_required
def requirements_report(request):
    from planning.models import PlanningJob

    jobs = PlanningJob.objects.exclude(status__in=['cancelled', 'archived'])
    department = request.GET.get('department')
    if department:
        jobs = jobs.filter(department=department)

    rows = explode_requirements(jobs)
    return render(request, 'bom/requirements_report.html', {'rows': rows})


@bom_manage_required
def raise_item_requests(request):
    from planning.models import PlanningJob

    jobs = PlanningJob.objects.exclude(status__in=['cancelled', 'archived'])
    rows = explode_requirements(jobs)
    selected_ids = set(request.POST.getlist('raw_item_id')) if request.method == 'POST' else set()

    if request.method == 'POST':
        form = RaiseItemRequestForm(request.POST)
        if form.is_valid() and selected_ids:
            selected_rows = [r for r in rows if str(r['raw_item'].pk) in selected_ids]
            created = raise_item_requests_from_shortfalls(
                selected_rows, request.user,
                form.cleaned_data['request_type'], form.cleaned_data['department'],
            )
            messages.success(request, f'Raised {len(created)} item request(s).')
            return redirect('bom:requirements_report')
    else:
        form = RaiseItemRequestForm()

    shortfall_rows = [r for r in rows if r['shortfall']]
    return render(request, 'bom/raise_item_requests.html', {'form': form, 'rows': shortfall_rows})
