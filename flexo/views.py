from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.core.exceptions import ValidationError
from django.core.paginator import Paginator
from django.db import transaction
from django.db.models import Count, Q
from django.http import JsonResponse
from django.shortcuts import get_object_or_404, redirect, render
from django.utils import timezone

from core.views import permission_required

from . import po_intake, services
from .forms import (
    CancelForm, FlexoJobCardForm, FlexoPlanningDispatchRunFormSet, FlexoPlanningJobEditForm,
    FlexoPrintingLogEntryForm, FlexoSkuRecipeForm, FlexoSlittingCuttingLogEntryForm,
    ManualPoHeaderForm, ManualPoLineFormSet,
)
from .models import (
    FLEXO_STATUS_CHOICES, LOCKED_STATUSES, MASTER_STATUS_CHOICES, FlexoPoDocument,
    PRIORITY_CHOICES, FlexoJobCard, FlexoJobChangeRequest, FlexoPlanningJob, FlexoSkuRecipe,
)

STATUS_LABELS = dict(FLEXO_STATUS_CHOICES)


def _profile_can(user, method):
    profile = getattr(user, 'profile', None)
    return bool(profile and getattr(profile, method)())


def _flash_errors(request, exc):
    for message in services.error_messages(exc):
        messages.error(request, message)


# ---------------------------------------------------------------------------
# Home
# ---------------------------------------------------------------------------

@login_required
def flexo_home(request):
    jobs = FlexoPlanningJob.objects.filter(is_active=True)
    by_status = dict(jobs.values_list('status').annotate(n=Count('id')))
    return render(request, 'flexo/flexo_home.html', {
        'status_counts': [(label, by_status.get(code, 0), code) for code, label in FLEXO_STATUS_CHOICES
                          if code not in ('cancelled',)],
        'active_count': jobs.count(),
        'pending_requests': FlexoJobChangeRequest.objects.filter(status='pending').count(),
        'sku_pending': FlexoSkuRecipe.objects.filter(
            is_active=True, master_data_status__in=['pending_review', 'reviewed']).count(),
    })


# ---------------------------------------------------------------------------
# Jobs queue
# ---------------------------------------------------------------------------

@login_required
@permission_required('can_view_jobcard')
def jobs(request):
    queryset = FlexoPlanningJob.objects.filter(is_active=True).select_related(
        'machine_name', 'department', 'job_card')

    if request.method == 'POST':
        if not _profile_can(request.user, 'can_plan'):
            messages.error(request, 'You do not have permission to modify planning jobs.')
            return redirect('flexo:jobs')
        if request.POST.get('action') == 'bulk_archive':
            ids = [int(i) for i in request.POST.getlist('selected_ids') if i.isdigit()]
            reason = (request.POST.get('archive_reason') or '').strip()
            if not ids:
                messages.error(request, 'Select at least one job to archive.')
            else:
                count = 0
                for job in FlexoPlanningJob.objects.filter(id__in=ids, is_active=True):
                    services.archive_job(job, actor=request.user, reason=reason)
                    count += 1
                messages.success(request, f'Archived {count} job(s).')
        return redirect('flexo:jobs')

    q = (request.GET.get('q') or '').strip()
    statuses = [s for s in request.GET.getlist('status') if s in STATUS_LABELS]
    hold = request.GET.get('hold') == '1'

    if q:
        queryset = queryset.filter(
            Q(jc_number__icontains=q) | Q(po_number__icontains=q) | Q(sku__icontains=q) | Q(job_name__icontains=q))
    if statuses:
        queryset = queryset.filter(status__in=statuses)
    elif not q:
        queryset = queryset.exclude(status='completed')
    if hold:
        queryset = queryset.filter(is_on_hold=True)

    queryset = queryset.order_by('-priority', '-created_at')
    page = Paginator(queryset, 50).get_page(request.GET.get('page'))
    params = request.GET.copy()
    params.pop('page', None)
    return render(request, 'flexo/jobs.html', {
        'page': page,
        'q': q, 'statuses': statuses, 'hold': hold,
        'status_choices': FLEXO_STATUS_CHOICES,
        'querystring': params.urlencode(),
        'can_plan': _profile_can(request.user, 'can_plan'),
    })


@login_required
@permission_required('can_edit_jobcard')
def jobs_archived(request):
    queryset = FlexoPlanningJob.objects.filter(is_active=False).select_related('archived_by')
    q = (request.GET.get('q') or '').strip()
    if q:
        queryset = queryset.filter(Q(jc_number__icontains=q) | Q(sku__icontains=q) | Q(po_number__icontains=q))
    page = Paginator(queryset.order_by('-archived_at', '-id'), 50).get_page(request.GET.get('page'))
    return render(request, 'flexo/jobs_archived.html', {
        'page': page, 'q': q, 'can_plan': _profile_can(request.user, 'can_plan'),
    })


# ---------------------------------------------------------------------------
# Job detail / edit / actions
# ---------------------------------------------------------------------------

@login_required
@permission_required('can_view_jobcard')
def job_detail(request, job_id):
    job = get_object_or_404(
        FlexoPlanningJob.objects.select_related('machine_name', 'department', 'sku_recipe'), id=job_id)
    pending_request = job.change_requests.filter(status='pending').first()
    return render(request, 'flexo/job_detail.html', {
        'job': job,
        'status_label': STATUS_LABELS.get(job.effective_status, job.effective_status),
        'dispatch_runs': job.dispatch_runs.all(),
        'events': job.events.select_related('actor')[:30],
        'pending_request': pending_request,
        'blockers': services.get_qc_submission_blockers(job) if job.status == 'draft' else [],
        'recipe_warnings': job.sku_recipe.warning_fields() if job.sku_recipe else [],
        'cancel_form': CancelForm(),
        'can_plan': _profile_can(request.user, 'can_plan'),
        'can_cancel': _profile_can(request.user, 'can_cancel_planning_job') or services.user_is_admin(request.user),
        'is_admin': services.user_is_admin(request.user),
    })


@login_required
@permission_required('can_plan')
def job_edit(request, job_id):
    job = get_object_or_404(FlexoPlanningJob, id=job_id)
    has_pending = job.change_requests.filter(status='pending').exists()
    if job.is_cancelled or not job.is_active:
        messages.error(request, 'Archived or cancelled jobs cannot be edited.')
        return redirect('flexo:job_detail', job_id=job.id)
    if job.status in LOCKED_STATUSES and not has_pending:
        messages.error(request, 'QC approved and released jobs are locked. Request a reopen from the job page.')
        return redirect('flexo:job_detail', job_id=job.id)

    if request.method == 'POST':
        form = FlexoPlanningJobEditForm(request.POST, instance=job)
        formset = FlexoPlanningDispatchRunFormSet(request.POST, instance=job)
        if form.is_valid() and formset.is_valid():
            with transaction.atomic():
                edited = form.save(commit=False)
                edited.job_card_version = (job.job_card_version or 1) + 1
                edited.last_edited_by = request.user
                edited.last_edited_at = timezone.now()
                edited.save()
                formset.save()
            if request.POST.get('submit_to_qc'):
                try:
                    services.submit_to_qc(edited, actor=request.user)
                    messages.success(request, f'{edited.jc_number} sent to QC.')
                except ValidationError as exc:
                    _flash_errors(request, exc)
                    return redirect('flexo:job_edit', job_id=edited.id)
            else:
                messages.success(request, f'Planning job {edited.jc_number} updated.')
            return redirect('flexo:job_detail', job_id=edited.id)
    else:
        form = FlexoPlanningJobEditForm(instance=job)
        formset = FlexoPlanningDispatchRunFormSet(instance=job)

    return render(request, 'flexo/job_edit.html', {
        'job': job, 'form': form, 'formset': formset,
        'blockers': services.get_qc_submission_blockers(job) if job.status == 'draft' else [],
    })


@login_required
@permission_required('can_plan')
def job_status_update(request, job_id):
    job = get_object_or_404(FlexoPlanningJob, id=job_id)
    if request.method == 'POST':
        transition = (request.POST.get('transition') or '').strip()
        try:
            if transition in {'submit_qc', 'submit_review'}:
                services.submit_to_qc(job, actor=request.user)
                messages.success(request, f'{job.jc_number}: draft -> pending_qc.')
            else:
                messages.error(request, 'Unknown status transition request.')
        except ValidationError as exc:
            _flash_errors(request, exc)
    return redirect('flexo:job_detail', job_id=job.id)


@login_required
@permission_required('can_plan')
def job_priority_update(request, job_id):
    if request.method != 'POST':
        return JsonResponse({'ok': False, 'error': 'POST request required'}, status=405)
    job = get_object_or_404(FlexoPlanningJob, id=job_id)
    try:
        value = int(request.POST.get('priority', 1))
    except (TypeError, ValueError):
        return JsonResponse({'ok': False, 'error': 'Invalid priority value'}, status=400)
    if value not in dict(PRIORITY_CHOICES):
        return JsonResponse({'ok': False, 'error': 'Invalid priority value'}, status=400)
    job.priority = value
    job.save(update_fields=['priority', 'updated_at'])
    return JsonResponse({'ok': True, 'priority': value, 'priority_display': job.get_priority_display()})


@login_required
def job_action(request, job_id):
    """Hold / release hold / archive / restore / cancel / request reopen / delete."""
    job = get_object_or_404(FlexoPlanningJob, id=job_id)
    if request.method != 'POST':
        return redirect('flexo:job_detail', job_id=job.id)

    action = (request.POST.get('action') or '').strip()
    reason = (request.POST.get('reason') or '').strip()
    user = request.user

    if not _profile_can(user, 'can_plan'):
        messages.error(request, 'You do not have permission to modify planning jobs.')
        return redirect('flexo:job_detail', job_id=job.id)

    try:
        if action == 'hold':
            services.hold_job(job, actor=user, reason=reason)
            messages.success(request, f'{job.jc_number} placed on hold.')
        elif action == 'release_hold':
            services.release_hold(job, actor=user)
            messages.success(request, f'{job.jc_number} hold released.')
        elif action == 'archive':
            services.archive_job(job, actor=user, reason=reason)
            messages.success(request, f'{job.jc_number} archived.')
            return redirect('flexo:jobs')
        elif action == 'restore':
            services.restore_job(job, actor=user, reason=reason)
            messages.success(request, f'{job.jc_number} restored.')
        elif action == 'cancel':
            if not (_profile_can(user, 'can_cancel_planning_job') or services.user_is_admin(user)):
                messages.error(request, 'You do not have permission to cancel planning jobs.')
            else:
                code = (request.POST.get('cancel_reason_code') or '').strip()
                if job.can_cancel_directly:
                    services.cancel_job(job, actor=user, reason=reason, reason_code=code)
                    messages.success(request, f'{job.jc_number} cancelled.')
                    return redirect('flexo:jobs')
                services.request_cancellation(job, actor=user, reason=reason, reason_code=code)
                messages.success(request, f'Cancellation request for {job.jc_number} sent for approval.')
        elif action == 'request_reopen':
            services.request_reopen(job, actor=user, reason=reason)
            messages.success(request, f'Reopen request for {job.jc_number} sent for Production Manager approval.')
        elif action == 'delete':
            if not services.user_is_admin(user):
                messages.error(request, 'Only administrators can delete planning jobs.')
            elif job.status not in {'draft', 'cancelled'} or hasattr(job, 'job_card'):
                messages.error(request, 'Only draft jobs without a job card can be deleted; archive or cancel instead.')
            else:
                label = job.jc_number
                job.delete()
                messages.success(request, f'{label} permanently deleted.')
                return redirect('flexo:jobs')
        else:
            messages.error(request, 'Unknown action.')
    except ValidationError as exc:
        _flash_errors(request, exc)
    return redirect('flexo:job_detail', job_id=job.id)


# ---------------------------------------------------------------------------
# Approval queue (QC -> release) and change requests
# ---------------------------------------------------------------------------

@login_required
@permission_required('can_view_approval_queue')
def approval_queue(request):
    user = request.user
    can_qc = _profile_can(user, 'can_approve_qc')
    can_release = _profile_can(user, 'can_approve_pm') or _profile_can(user, 'can_plan')
    can_pm = _profile_can(user, 'can_approve_pm')

    if request.method == 'POST':
        action = (request.POST.get('action') or '').strip()
        reason = (request.POST.get('reason') or '').strip()
        job = get_object_or_404(FlexoPlanningJob, pk=request.POST.get('planning_job_id'), is_active=True)
        try:
            if action in {'approve_qc', 'reject_qc'}:
                if not can_qc:
                    messages.error(request, 'You do not have permission to approve QC jobs.')
                elif action == 'approve_qc':
                    services.approve_qc(job, actor=user, reason=reason)
                    messages.success(request, f'{job.jc_number} QC approved.')
                else:
                    services.reject_qc(job, actor=user, reason=reason)
                    messages.success(request, f'{job.jc_number} sent back to Draft.')
            elif action == 'release_for_production':
                if not can_release:
                    messages.error(request, 'You do not have permission to release jobs for production.')
                else:
                    services.release_for_production(job, actor=user, reason=reason)
                    messages.success(request, f'{job.jc_number} released for production.')
            else:
                messages.error(request, 'Unknown action.')
        except ValidationError as exc:
            _flash_errors(request, exc)
        return redirect('flexo:approval_queue')

    q = (request.GET.get('q') or '').strip()
    base = FlexoPlanningJob.objects.filter(is_active=True).select_related('machine_name')
    if q:
        base = base.filter(Q(jc_number__icontains=q) | Q(po_number__icontains=q) | Q(sku__icontains=q))
    return render(request, 'flexo/approval_queue.html', {
        'q': q,
        'draft_jobs': base.filter(status='draft')[:200],
        'qc_jobs': base.filter(status='pending_qc').order_by('-updated_at')[:300],
        'release_jobs': base.filter(status='qc_approved').order_by('-updated_at')[:300],
        'change_requests': FlexoJobChangeRequest.objects.filter(status='pending')
            .select_related('planning_job', 'requested_by'),
        'can_qc': can_qc, 'can_release': can_release, 'can_pm': can_pm,
    })


@login_required
def change_request_approve(request, request_id):
    change_request = get_object_or_404(FlexoJobChangeRequest, id=request_id)
    if request.method == 'POST':
        if not _profile_can(request.user, 'can_approve_pm'):
            messages.error(request, 'You do not have permission to approve change requests.')
        else:
            try:
                services.approve_change_request(change_request, actor=request.user)
                messages.success(request, f'Request for {change_request.planning_job.jc_number} approved.')
            except ValidationError as exc:
                _flash_errors(request, exc)
    return redirect('flexo:approval_queue')


@login_required
def change_request_reject(request, request_id):
    change_request = get_object_or_404(FlexoJobChangeRequest, id=request_id)
    if request.method == 'POST':
        if not _profile_can(request.user, 'can_approve_pm'):
            messages.error(request, 'You do not have permission to reject change requests.')
        else:
            try:
                services.reject_change_request(
                    change_request, actor=request.user, reason=(request.POST.get('rejection_reason') or '').strip())
                messages.warning(request, f'Request for {change_request.planning_job.jc_number} rejected.')
            except ValidationError as exc:
                _flash_errors(request, exc)
    return redirect('flexo:approval_queue')


# ---------------------------------------------------------------------------
# SKU master
# ---------------------------------------------------------------------------

@login_required
@permission_required('can_edit_jobcard')
def sku_recipes(request, status=None):
    queryset = FlexoSkuRecipe.objects.filter(is_active=True).select_related('machine')
    q = (request.GET.get('q') or '').strip()
    if status:
        queryset = queryset.filter(master_data_status=status)
    if q:
        queryset = queryset.filter(Q(sku__icontains=q) | Q(job_name__icontains=q))
    counts = dict(FlexoSkuRecipe.objects.filter(is_active=True).values_list('master_data_status').annotate(n=Count('id')))
    page = Paginator(queryset.order_by('sku'), 50).get_page(request.GET.get('page'))
    return render(request, 'flexo/sku_recipes.html', {
        'page': page, 'q': q, 'status': status or '',
        'tabs': [('', 'All', sum(counts.values()))] + [(c, l, counts.get(c, 0)) for c, l in MASTER_STATUS_CHOICES],
    })


@login_required
@permission_required('can_edit_jobcard')
def sku_recipes_archived(request):
    queryset = FlexoSkuRecipe.objects.filter(is_active=False).order_by('-archived_at', 'sku')
    return render(request, 'flexo/sku_recipes_archived.html', {
        'page': Paginator(queryset, 50).get_page(request.GET.get('page')),
    })


def _sync_draft_jobs_for_recipe(recipe):
    updated = 0
    for job in FlexoPlanningJob.objects.filter(sku_recipe=recipe, status='draft', is_active=True):
        services.apply_recipe_to_job(job, recipe)
        job.save()
        updated += 1
    return updated


@login_required
@permission_required('can_edit_jobcard')
def sku_recipe_edit(request, recipe_id=None):
    recipe = get_object_or_404(FlexoSkuRecipe, id=recipe_id, is_active=True) if recipe_id else None
    user = request.user
    is_admin = services.user_is_admin(user)
    can_review = is_admin or _profile_can(user, 'can_approve_sku_master_review')
    locked = bool(recipe and recipe.is_approved)

    def render_form(form, obj=None):
        current = obj if obj is not None else recipe
        return render(request, 'flexo/sku_recipe_edit.html', {
            'form': form, 'recipe': current, 'locked': bool(current and current.is_approved),
            'missing': current.missing_required_fields() if current else [],
            'warnings': current.warning_fields() if current else [],
            'can_review': can_review, 'is_admin': is_admin,
        })

    if request.method == 'POST':
        action = (request.POST.get('action') or '').strip()

        if recipe and action == 'archive':
            if recipe.is_approved and not is_admin:
                messages.error(request, 'Approved records can only be archived by admin users.')
            else:
                recipe.is_active = False
                recipe.archived_by, recipe.archived_at = user, timezone.now()
                recipe.archive_reason = (request.POST.get('archive_reason') or '').strip()
                recipe.save(update_fields=['is_active', 'archived_by', 'archived_at', 'archive_reason', 'updated_at'])
                messages.success(request, f'SKU master "{recipe.sku}" archived.')
            return redirect('flexo:sku_recipes')

        if recipe and action == 'reopen_sku' and recipe.is_approved:
            comment = (request.POST.get('rejection_comment') or '').strip()
            if not comment:
                messages.error(request, 'Enter a reopen reason, then click Reopen SKU.')
                return render_form(FlexoSkuRecipeForm(instance=recipe))
            recipe.master_data_status = 'draft'
            recipe.reviewed_by = recipe.approved_by = None
            recipe.reviewed_at = recipe.approved_at = None
            recipe.rejection_comment, recipe.last_rejected_by, recipe.last_rejected_at = comment, user, timezone.now()
            recipe.save()
            messages.success(request, f'SKU master "{recipe.sku}" reopened to Draft.')
            return redirect('flexo:sku_recipe_edit', recipe_id=recipe.id)

        if locked:
            messages.error(request, 'Approved SKU is locked. Use Reopen SKU before making edits.')
            return render_form(FlexoSkuRecipeForm(instance=recipe))

        form = FlexoSkuRecipeForm(request.POST, instance=recipe)
        if not form.is_valid():
            messages.error(request, 'There are errors in the form. Please correct them and try again.')
            return render_form(form)

        obj = form.save(commit=False)
        if not recipe:
            obj.created_by = user
        status = recipe.master_data_status if recipe else 'draft'
        sync_jobs = False

        if recipe and action:
            missing = obj.missing_required_fields()
            if action == 'submit_review' and status == 'draft':
                if missing:
                    messages.error(request, f'Cannot submit for review. Missing: {", ".join(missing)}.')
                    return render_form(form, obj)
                obj.master_data_status = 'pending_review'
                messages.success(request, f'SKU master "{obj.sku}" submitted for review.')
            elif action == 'review' and status == 'pending_review':
                if not can_review:
                    messages.error(request, 'You do not have permission to review SKU masters.')
                    return render_form(form, obj)
                if missing:
                    messages.error(request, f'Cannot review. Missing: {", ".join(missing)}.')
                    return render_form(form, obj)
                obj.master_data_status = 'reviewed'
                obj.reviewed_by, obj.reviewed_at = user, timezone.now()
                messages.success(request, f'SKU master "{obj.sku}" reviewed.')
            elif action == 'approve' and status == 'reviewed':
                if not can_review:
                    messages.error(request, 'You do not have permission to approve SKU masters.')
                    return render_form(form, obj)
                if missing:
                    messages.error(request, f'Cannot approve. Missing: {", ".join(missing)}.')
                    return render_form(form, obj)
                obj.master_data_status = 'approved'
                obj.approved_by, obj.approved_at = user, timezone.now()
                sync_jobs = True
                messages.success(request, f'SKU master "{obj.sku}" approved.')
            elif action == 'back_to_draft' and status in ('pending_review', 'reviewed'):
                comment = (request.POST.get('rejection_comment') or '').strip()
                if not comment:
                    messages.error(request, 'Please provide a reason when sending a record back to Draft.')
                    return render_form(form, obj)
                obj.master_data_status = 'draft'
                obj.reviewed_by = obj.approved_by = None
                obj.reviewed_at = obj.approved_at = None
                obj.rejection_comment, obj.last_rejected_by, obj.last_rejected_at = comment, user, timezone.now()
                messages.success(request, f'SKU master "{obj.sku}" moved back to Draft.')
            else:
                messages.info(request, f'SKU master "{obj.sku}" saved without changing workflow status.')
        elif recipe:
            messages.success(request, f'SKU master "{obj.sku}" saved.')
        else:
            obj.master_data_status = 'draft'
            messages.success(request, f'SKU master "{obj.sku}" saved as Draft. Submit it for review from its page.')

        obj.save()
        if sync_jobs:
            refreshed = _sync_draft_jobs_for_recipe(obj)
            if refreshed:
                messages.success(request, f'{refreshed} draft planning job(s) refreshed from the approved SKU master.')
        return redirect('flexo:sku_recipes')

    return render_form(FlexoSkuRecipeForm(instance=recipe))


# ---------------------------------------------------------------------------
# Manual PO entry
# ---------------------------------------------------------------------------

@login_required
@permission_required('can_edit_jobcard')
def manual_po_entry(request):
    if request.method == 'POST':
        header = ManualPoHeaderForm(request.POST, prefix='h')
        formset = ManualPoLineFormSet(request.POST, prefix='l')
        if header.is_valid() and formset.is_valid():
            lines = [f.cleaned_data for f in formset if f.cleaned_data.get('sku')]
            if not lines:
                messages.error(request, 'Add at least one SKU line.')
            else:
                created = []
                try:
                    with transaction.atomic():
                        for line in lines:
                            created.append(services.create_job_from_recipe(
                                line['sku'], actor=request.user,
                                po_number=header.cleaned_data['po_number'],
                                po_received_date=header.cleaned_data['po_received_date'],
                                delivery_date=line.get('delivery_date'),
                                order_qty=line['order_qty'], stock_qty=line.get('stock_qty') or 0,
                                destination=header.cleaned_data.get('destination', ''),
                                priority=int(header.cleaned_data['priority']),
                            ))
                except ValidationError as exc:
                    _flash_errors(request, exc)
                else:
                    messages.success(request, f'Created {len(created)} planning job(s): '
                                              f'{", ".join(j.jc_number for j in created)}.')
                    return redirect('flexo:jobs')
    else:
        header = ManualPoHeaderForm(prefix='h')
        formset = ManualPoLineFormSet(prefix='l')
    return render(request, 'flexo/manual_po_entry.html', {
        'header': header, 'formset': formset,
        'approved_count': FlexoSkuRecipe.objects.filter(is_active=True, master_data_status='approved').count(),
    })


# ---------------------------------------------------------------------------
# Job card
# ---------------------------------------------------------------------------

@login_required
@permission_required('can_view_jobcard')
def job_card_detail(request, pk):
    card = get_object_or_404(FlexoJobCard, pk=pk)
    can_edit = _profile_can(request.user, 'can_edit_jobcard')
    if request.method == 'POST':
        if not can_edit:
            messages.error(request, 'You do not have permission to edit job cards.')
            return redirect('flexo:job_card_detail', pk=card.pk)
        form = FlexoJobCardForm(request.POST, instance=card)
        if form.is_valid():
            form.save()
            messages.success(request, f'Job Card {card.job_card_no} updated.')
            return redirect('flexo:job_card_detail', pk=card.pk)
    else:
        form = FlexoJobCardForm(instance=card)
    return render(request, 'flexo/job_card_detail.html', {
        'card': card, 'form': form, 'can_edit': can_edit,
        'printing_entries': card.printing_log_entries.all(),
        'slitting_entries': card.slitting_cutting_log_entries.all(),
        'dispatch_runs': card.planning_job.dispatch_runs.all(),
        'printing_form': FlexoPrintingLogEntryForm(),
        'slitting_form': FlexoSlittingCuttingLogEntryForm(),
    })


@login_required
@permission_required('can_edit_jobcard')
def job_card_add_printing_entry(request, pk):
    card = get_object_or_404(FlexoJobCard, pk=pk)
    if request.method == 'POST':
        form = FlexoPrintingLogEntryForm(request.POST)
        if form.is_valid():
            entry = form.save(commit=False)
            entry.job_card = card
            entry.save()
            messages.success(request, 'Printing log entry added.')
    return redirect('flexo:job_card_detail', pk=card.pk)


@login_required
@permission_required('can_edit_jobcard')
def job_card_add_slitting_entry(request, pk):
    card = get_object_or_404(FlexoJobCard, pk=pk)
    if request.method == 'POST':
        form = FlexoSlittingCuttingLogEntryForm(request.POST)
        if form.is_valid():
            entry = form.save(commit=False)
            entry.job_card = card
            entry.save()
            messages.success(request, 'Slitting/Cutting log entry added.')
    return redirect('flexo:job_card_detail', pk=card.pk)


@login_required
@permission_required('can_view_jobcard')
def job_card_print(request, pk):
    card = get_object_or_404(FlexoJobCard, pk=pk)
    return render(request, 'flexo/job_card_print.html', {
        'card': card,
        'job': card.planning_job,
        'issue_roll_qty': round(card.roll_meter_qty / 1000, 4) if card.roll_meter_qty else None,
        'printing_entries': card.printing_log_entries.all(),
        'slitting_entries': card.slitting_cutting_log_entries.all(),
        'dispatch_runs': card.planning_job.dispatch_runs.all(),
    })


# ---------------------------------------------------------------------------
# PO / WO intake
# ---------------------------------------------------------------------------

@login_required
@permission_required('can_edit_jobcard')
def po_upload(request):
    if request.method == 'POST':
        upload = request.FILES.get('po_file')
        if not upload:
            messages.error(request, 'Choose a PO/WO PDF to upload.')
        elif not upload.name.lower().endswith('.pdf'):
            messages.error(request, 'Only PDF files are supported.')
        else:
            try:
                payload = po_intake.extract_pdf(upload)
                doc = po_intake.save_extracted_document(payload, actor=request.user)
            except ValidationError as exc:
                _flash_errors(request, exc)
            except Exception as exc:  # noqa: BLE001 - extractor failures must not 500 the page
                messages.error(request, f'Could not read this PDF: {exc}')
            else:
                messages.success(request, f'{doc.po_number} read: {len(doc.items)} line(s). Review and create jobs below.')
                return redirect('flexo:po_review', doc_id=doc.pk)
    return render(request, 'flexo/po_upload.html')


@login_required
@permission_required('can_edit_jobcard')
def po_inbox(request):
    status = (request.GET.get('status') or '').strip()
    docs = FlexoPoDocument.objects.select_related('uploaded_by')
    if status in {'pending', 'processed'}:
        docs = docs.filter(status=status)
    q = (request.GET.get('q') or '').strip()
    if q:
        docs = docs.filter(po_number__icontains=q)
    page = Paginator(docs, 50).get_page(request.GET.get('page'))
    return render(request, 'flexo/po_inbox.html', {'page': page, 'status': status, 'q': q})


@login_required
@permission_required('can_edit_jobcard')
def po_review(request, doc_id):
    doc = get_object_or_404(FlexoPoDocument, pk=doc_id)

    if request.method == 'POST':
        action = (request.POST.get('action') or '').strip()
        try:
            if action == 'create_jobs':
                indexes = [int(i) for i in request.POST.getlist('line') if i.isdigit()]
                if not indexes:
                    messages.error(request, 'Select at least one ready line.')
                else:
                    created, skipped = po_intake.create_jobs_from_document(doc, indexes, actor=request.user)
                    if created:
                        messages.success(request, f'Created {len(created)} planning job(s): '
                                                  f'{", ".join(j.jc_number for j in created)}.')
                    if skipped:
                        messages.warning(request, f'{len(skipped)} selected line(s) were not ready and were skipped.')
            elif action == 'create_draft_skus':
                created = po_intake.create_draft_recipes(doc, actor=request.user)
                if created:
                    messages.success(request, f'Created {len(created)} draft SKU master(s). Complete and approve '
                                              'them, then come back to create the jobs.')
                else:
                    messages.info(request, 'No missing SKUs to create.')
            else:
                messages.error(request, 'Unknown action.')
        except ValidationError as exc:
            _flash_errors(request, exc)
        return redirect('flexo:po_review', doc_id=doc.pk)

    rows = po_intake.analyze_lines(doc)
    payload = doc.payload or {}
    return render(request, 'flexo/po_review.html', {
        'doc': doc, 'payload': payload, 'rows': rows,
        'ready_count': sum(1 for r in rows if r['state'] == 'ready'),
        'missing_count': sum(1 for r in rows if r['state'] == 'missing'),
    })
