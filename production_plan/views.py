import datetime
import io

from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.http import HttpResponse, HttpResponseForbidden, JsonResponse
from django.shortcuts import get_object_or_404, redirect, render
from django.urls import reverse
from django.utils import timezone
from django.views.decorators.http import require_POST

from core.models import Machine
from planning.models import PlanningJob
from production_plan import engine, services
from production_plan.models import PRIORITY_LABELS, PlanBlock, PlanItem, PlanSettings, ProductionPlan


def _back(plan_lane=''):
    url = reverse('production_plan:board')
    return redirect(f'{url}?lane={plan_lane}' if plan_lane else url)


def _slot_options(plan, settings_obj):
    out = []
    for d in plan.horizon_dates:
        day = datetime.date.fromisoformat(d)
        for shift in engine.SHIFTS:
            out.append({'value': f'{d}|{shift}', 'label': services.slot_label(day, shift)})
    return out


@login_required
def board(request):
    settings_obj = PlanSettings.get()
    plan = None
    plan_id = request.GET.get('plan')
    if plan_id and plan_id.isdigit():
        plan = ProductionPlan.objects.filter(pk=int(plan_id)).first()
    plan = plan or services.current_plan()
    can_edit = services.can_edit_plan(request.user)
    lane = request.GET.get('lane') if request.GET.get('lane') in engine.LANES else 'black'
    tab = request.GET.get('tab') if request.GET.get('tab') in ('plan', 'actual', 'history') else 'plan'
    context = {
        'plan': plan, 'can_edit': can_edit, 'is_admin': request.user.is_superuser or request.user.is_staff,
        'lane': lane, 'tab': tab, 'lanes': [(k, engine.LANE_LABELS[k]) for k in engine.LANES],
        'priority_choices': sorted(PRIORITY_LABELS.items()), 'settings': settings_obj,
        'gated': services.pending_draft() is not None,
    }
    if plan is None:
        return render(request, 'production_plan/board.html', context)
    grid = services.plan_grid(plan, settings_obj)
    context.update({
        'grid': grid[lane], 'lane_stats': (plan.summary or {}).get('lanes', {}), 'overflow_new': (plan.summary or {}).get('overflow_new', []),
        'diff': (plan.summary or {}).get('diff', {}), 'slot_options': _slot_options(plan, settings_obj),
        'lane_tabs': [{'key': k, 'label': engine.LANE_LABELS[k], 'planned': g['planned'], 'capacity': g['capacity']} for k, g in grid.items()],
        'stats': (plan.summary or {}).get('lanes', {}).get(lane),
        'machines_all': [m for g in grid.values() for m in g['machines']],
        'blocks': PlanBlock.objects.filter(date__gte=timezone.localdate()).select_related('machine'),
    })
    cells = []
    for m in grid[lane]['machines']:
        row = {'machine': m, 'cells': []}
        for d, sh in grid[lane]['slots']:
            cell = grid[lane]['cells'][(m.id, d, sh)]
            row['cells'].append({'date': d, 'shift': sh, 'slot_value': f'{d.isoformat()}|{sh}', **cell})
        cells.append(row)
    context['rows'] = cells
    context['slot_headers'] = [{'date': d, 'shift': sh, 'label': services.slot_label(d, sh)} for d, sh in grid[lane]['slots']]
    if tab == 'actual':
        context['actual_rows'] = [r for r in services.plan_vs_actual(plan, settings_obj) if r['lane'] == lane]
    if tab == 'history':
        context['events'] = plan.events.select_related('user')[:80]
    night_date, missing = services.night_entries_missing(plan.morning_date, settings_obj)
    context['night_missing'] = {'date': night_date, 'machines': missing}
    return render(request, 'production_plan/board.html', context)


def _require_editor(request):
    return None if services.can_edit_plan(request.user) else HttpResponseForbidden('Only planners and admins can change the plan.')


def _editable_plan(plan_id):
    plan = get_object_or_404(ProductionPlan, pk=plan_id)
    if plan.status == 'superseded':
        raise ValueError('This plan has been replaced by a newer one; open the current plan.')
    return plan


@login_required
@require_POST
def approve(request, plan_id):
    denied = _require_editor(request)
    if denied:
        return denied
    plan = get_object_or_404(ProductionPlan, pk=plan_id)
    try:
        services.approve(plan, request.user)
        messages.success(request, 'Plan approved. It goes out by email at the usual time, and planning is open again.')
    except ValueError as exc:
        messages.error(request, str(exc))
    return _back()


@login_required
@require_POST
def recalculate(request, plan_id):
    denied = _require_editor(request)
    if denied:
        return denied
    plan = get_object_or_404(ProductionPlan, pk=plan_id)
    services.rebalance(plan)
    services.log_event(plan, request.user, 'recalculated', 'planner asked for a recalculation')
    messages.success(request, 'Plan recalculated: your changes were kept, printed jobs removed, free capacity refilled.')
    return _back(request.POST.get('lane', ''))


def _job_remaining(job_id):
    info = engine.load_jobs(PlanSettings.get(), timezone.localdate()).get(job_id)
    return info


@login_required
@require_POST
def item_add(request, plan_id):
    denied = _require_editor(request)
    if denied:
        return denied
    lane = request.POST.get('lane', '')
    try:
        plan = _editable_plan(plan_id)
        settings_obj = PlanSettings.get()
        job_id = int(request.POST.get('job_id') or 0)
        info = _job_remaining(job_id)
        if info is None or not info['lane']:
            raise ValueError('That job cannot be planned: it must be released, have its plates, and not be on hold or merged under another job.')
        date_text, shift = (request.POST.get('slot') or '').split('|')
        day = datetime.date.fromisoformat(date_text)
        machine = Machine.objects.get(pk=int(request.POST.get('machine_id')))
        qty = int(request.POST.get('qty') or 0)
        if qty <= 0:
            raise ValueError('Enter how many impressions to plan.')
        already = sum(i.qty for i in plan.items.filter(planning_job_id=job_id))
        if qty > max(info['remaining'] - already, 0):
            raise ValueError(f"Only {max(info['remaining'] - already, 0):,} impressions of {info['job'].jc_number} are left to plan.")
        if shift == 'N' and info['is_new']:
            raise ValueError('A new job runs in a day shift. Move it to a day shift.')
        reason = (request.POST.get('fifo_reason') or '').strip()
        window_jobs = _fifo_window_jobs(plan, info['lane'], settings_obj)
        if job_id not in window_jobs and not reason:
            raise ValueError('This job is outside the FIFO order for its lane. Write the reason in the FIFO reason box to plan it anyway.')
        last = plan.items.filter(machine=machine, plan_date=day, shift=shift).order_by('-sequence').first()
        PlanItem.objects.create(
            plan=plan, lane=info['lane'], machine=machine, plan_date=day, shift=shift, sequence=(last.sequence + 1) if last else 1,
            planning_job=info['job'], jc_number=info['job'].jc_number, sku=info['job'].sku or '', order_date=info['order_date'], qty=qty,
            remaining_at_plan=info['remaining'], priority=info['priority'], is_new=info['is_new'], pinned=True, manual=True,
            fifo_reason=reason if job_id not in window_jobs else '',
        )
        services.log_event(plan, request.user, 'item_added', f"{info['job'].jc_number} {qty:,} on {machine.name} {date_text} {shift}" + (f' (FIFO jump: {reason})' if reason else ''))
        messages.success(request, f"Added {info['job'].jc_number}: {qty:,} impressions on {machine.name}.")
    except (ValueError, Machine.DoesNotExist) as exc:
        messages.error(request, str(exc) or 'Could not add the job.')
    return _back(lane)


def _fifo_window_jobs(plan, lane, settings_obj):
    """Job ids that FIFO order would already put in this lane's plan (what is planned plus the oldest waiting)."""
    info = engine.load_jobs(settings_obj, timezone.localdate())
    planned = {i.planning_job_id: i.qty for i in plan.items.filter(lane=lane)}
    waiting = sorted(
        (j for j, v in info.items() if v['lane'] == lane and v['remaining'] > planned.get(j, 0)),
        key=lambda j: (-info[j]['priority'], info[j]['order_date'], info[j]['job'].jc_number),
    )
    window, acc = set(planned), 0
    cap = sum(c['cap'] for c in services.plan_grid(plan, settings_obj)[lane]['cells'].values()) - sum(planned.values())
    for j in waiting:
        if acc >= max(cap, 0):
            break
        window.add(j)
        acc += info[j]['remaining'] - planned.get(j, 0)
    return window


@login_required
@require_POST
def item_remove(request, plan_id, item_id):
    denied = _require_editor(request)
    if denied:
        return denied
    plan = get_object_or_404(ProductionPlan, pk=plan_id)
    item = get_object_or_404(PlanItem, pk=item_id, plan=plan)
    services.log_event(plan, request.user, 'item_removed', f'{item.jc_number} {item.qty:,} from {item.machine.name} {item.plan_date} {item.shift}')
    lane = item.lane
    item.delete()
    messages.success(request, f'Removed {item.jc_number} from the plan. It stays eligible and may return when the plan is recalculated; pin other jobs or put the job on hold to keep it out.')
    return _back(lane)


@login_required
@require_POST
def item_move(request, plan_id, item_id):
    denied = _require_editor(request)
    if denied:
        return denied
    plan = get_object_or_404(ProductionPlan, pk=plan_id)
    item = get_object_or_404(PlanItem, pk=item_id, plan=plan)
    try:
        date_text, shift = (request.POST.get('slot') or '').split('|')
        day = datetime.date.fromisoformat(date_text)
        machine = Machine.objects.get(pk=int(request.POST.get('machine_id')))
        if shift == 'N' and item.is_new:
            raise ValueError('A new job runs in a day shift.')
        qty = int(request.POST.get('qty') or item.qty)
        if qty <= 0 or qty > item.qty:
            raise ValueError(f'Move between 1 and {item.qty:,} impressions.')
        last = plan.items.filter(machine=machine, plan_date=day, shift=shift).order_by('-sequence').first()
        target = dict(plan=plan, lane=item.lane, machine=machine, plan_date=day, shift=shift, sequence=(last.sequence + 1) if last else 1,
                      planning_job=item.planning_job, jc_number=item.jc_number, sku=item.sku, order_date=item.order_date, qty=qty,
                      remaining_at_plan=item.remaining_at_plan, priority=item.priority, is_new=item.is_new, is_chunk=item.is_chunk,
                      pinned=True, manual=True, fifo_reason=item.fifo_reason)
        PlanItem.objects.create(**target)
        if qty == item.qty:
            item.delete()
        else:
            item.qty -= qty
            item.save(update_fields=['qty'])
        services.log_event(plan, request.user, 'item_moved', f'{item.jc_number} {qty:,} to {machine.name} {date_text} {shift}')
        messages.success(request, f'Moved {qty:,} impressions of {item.jc_number} to {machine.name}.')
    except (ValueError, Machine.DoesNotExist) as exc:
        messages.error(request, str(exc) or 'Could not move the job.')
    return _back(item.lane)


@login_required
@require_POST
def item_pin(request, plan_id, item_id):
    denied = _require_editor(request)
    if denied:
        return denied
    plan = get_object_or_404(ProductionPlan, pk=plan_id)
    item = get_object_or_404(PlanItem, pk=item_id, plan=plan)
    item.pinned = not item.pinned
    item.save(update_fields=['pinned'])
    services.log_event(plan, request.user, 'item_pinned' if item.pinned else 'item_unpinned', item.jc_number)
    return _back(item.lane)


@login_required
@require_POST
def set_priority(request, plan_id):
    """Priority lives on the planning job (Urgent..Low); the engine puts higher priority ahead of age."""
    denied = _require_editor(request)
    if denied:
        return denied
    plan = get_object_or_404(ProductionPlan, pk=plan_id)
    job = get_object_or_404(PlanningJob, pk=int(request.POST.get('job_id') or 0))
    try:
        priority = int(request.POST.get('priority'))
        if priority not in PRIORITY_LABELS:
            raise ValueError
    except (TypeError, ValueError):
        messages.error(request, 'Pick a priority from Low to Urgent.')
        return _back(request.POST.get('lane', ''))
    old = job.priority
    job.priority = priority
    job.save(update_fields=['priority'])
    plan.items.filter(planning_job=job).update(priority=priority)
    services.log_event(plan, request.user, 'priority', f'{job.jc_number}: {PRIORITY_LABELS.get(old, old)} -> {PRIORITY_LABELS[priority]}')
    messages.success(request, f'{job.jc_number} is now {PRIORITY_LABELS[priority]} priority. Recalculate to apply it to free capacity.')
    return _back(request.POST.get('lane', ''))


@login_required
@require_POST
def block_add(request, plan_id):
    denied = _require_editor(request)
    if denied:
        return denied
    plan = get_object_or_404(ProductionPlan, pk=plan_id)
    try:
        machine = Machine.objects.get(pk=int(request.POST.get('machine_id')))
        day = datetime.date.fromisoformat(request.POST.get('date'))
        shift = request.POST.get('shift') if request.POST.get('shift') in ('D', 'N', 'B') else 'B'
        PlanBlock.objects.get_or_create(machine=machine, date=day, shift=shift, defaults={
            'reason': (request.POST.get('reason') or 'Blocked by planner')[:200], 'created_by': request.user})
        services.log_event(plan, request.user, 'block_added', f'{machine.name} {day} {shift}')
        messages.success(request, f'{machine.name} is blocked on {day:%a %d %b}. Recalculate to move its jobs.')
    except (ValueError, TypeError, Machine.DoesNotExist):
        messages.error(request, 'Pick a machine and a date.')
    return _back(request.POST.get('lane', ''))


@login_required
@require_POST
def block_remove(request, plan_id, block_id):
    denied = _require_editor(request)
    if denied:
        return denied
    plan = get_object_or_404(ProductionPlan, pk=plan_id)
    PlanBlock.objects.filter(pk=block_id).delete()
    services.log_event(plan, request.user, 'block_removed', str(block_id))
    return _back(request.POST.get('lane', ''))


@login_required
def job_search(request):
    """Type-ahead for the add-job box: eligible jobs matching a JC or SKU."""
    if not services.can_edit_plan(request.user):
        return JsonResponse({'results': []}, status=403)
    q = (request.GET.get('q') or '').strip().lower()
    if len(q) < 2:
        return JsonResponse({'results': []})
    info = engine.load_jobs(PlanSettings.get(), timezone.localdate())
    rows = []
    for job_id, v in info.items():
        job = v['job']
        if v['lane'] and v['remaining'] > 0 and (q in job.jc_number.lower() or q in (job.sku or '').lower()):
            rows.append({'id': job_id, 'jc': job.jc_number, 'sku': job.sku, 'lane': v['lane'], 'remaining': v['remaining'],
                         'new': v['is_new'], 'order_date': v['order_date'].isoformat()})
    rows.sort(key=lambda r: (r['order_date'], r['jc']))
    return JsonResponse({'results': rows[:12]})


@login_required
def export(request, plan_id, kind):
    from reports.export.services import export_as_pdf, export_as_xlsx

    plan = get_object_or_404(ProductionPlan, pk=plan_id)
    payload = services.export_payload(plan)
    stem = f'production-plan-{plan.morning_date.isoformat()}'
    if kind == 'xlsx':
        response = HttpResponse(export_as_xlsx(payload), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    elif kind == 'pdf':
        response = HttpResponse(export_as_pdf(payload), content_type='application/pdf')
    else:
        return HttpResponse(status=404)
    response['Content-Disposition'] = f'attachment; filename="{stem}.{kind}"'
    return response


@login_required
def share_text(request, plan_id):
    """The WhatsApp message for a plan (text); the picture comes from share_image."""
    plan = get_object_or_404(ProductionPlan, pk=plan_id)
    return JsonResponse({'text': services.whatsapp_text(plan), 'image_url': reverse('production_plan:share_image', args=[plan.pk])})


@login_required
def share_image(request, plan_id):
    plan = get_object_or_404(ProductionPlan, pk=plan_id)
    response = HttpResponse(services.render_plan_image(plan), content_type='image/png')
    response['Content-Disposition'] = 'inline; filename="production-plan.png"'
    response['Cache-Control'] = 'no-store'
    return response


@login_required
@require_POST
def email_now(request, plan_id):
    """Admin: send the plan email right now (for a test or a resend)."""
    if not (request.user.is_superuser or request.user.is_staff):
        return HttpResponseForbidden('Admin only.')
    plan = get_object_or_404(ProductionPlan, pk=plan_id)
    try:
        services.send_plan_email(plan)
        messages.success(request, 'Plan email sent to the planners with printing in CC.')
    except Exception as exc:
        messages.error(request, f'Could not send the email: {exc}')
    return _back()


@login_required
def settings_page(request):
    if not (request.user.is_superuser or request.user.is_staff):
        return HttpResponseForbidden('Admin only.')
    from production_plan.forms import PlanSettingsForm

    obj = PlanSettings.get()
    form = PlanSettingsForm(request.POST or None, instance=obj)
    if request.method == 'POST' and form.is_valid():
        form.save()
        messages.success(request, 'Plan settings saved.')
        return redirect('production_plan:settings')
    return render(request, 'production_plan/settings.html', {'form': form})
