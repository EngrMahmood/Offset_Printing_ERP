import csv
import json
from datetime import date, datetime, timedelta

from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.db.models import Count, Q, Sum
from django.http import HttpResponse, JsonResponse
from django.shortcuts import get_object_or_404, redirect, render
from django.utils import timezone

from core.models import JobCard, Machine, Operator
from core.views import permission_required
from workflow.services import start_production

from . import services
from .models import CoatingRun


def _int(value, default=0):
    text = str(value or '').replace(',', '').strip()
    if not text:
        return default
    try:
        return int(text)
    except ValueError:
        raise ValueError(f'"{value}" is not a whole number.')


def _parse_time(value):
    value = (value or '').strip()
    return datetime.strptime(value, '%H:%M').time() if value else None


def _minutes_between(start, end):
    if not start or not end:
        return 0
    delta = datetime.combine(date.today(), end) - datetime.combine(date.today(), start)
    if delta.total_seconds() < 0:
        delta += timedelta(days=1)
    return round(delta.total_seconds() / 60, 1)


def _save_from_post(request, record):
    """Validate the form and create or update a run. Raises ValueError with a message for the user."""
    post = request.POST
    job_card = get_object_or_404(JobCard, pk=post.get('job_card'), is_active=True)
    shift = (post.get('shift') or '').strip()
    if shift not in ('A', 'B'):
        raise ValueError('Shift is required.')
    if not post.get('date'):
        raise ValueError('Date is required.')
    sheets = _int(post.get('sheets_coated'))
    waste = _int(post.get('waste_sheets'))
    coats = _int(post.get('coats'), 1)
    if sheets < 0 or waste < 0 or coats < 1:
        raise ValueError('Quantities cannot be negative and coats must be at least 1.')
    if sheets == 0 and waste == 0:
        raise ValueError('Enter the sheets coated and/or the waste sheets.')
    if not job_card.printing_productions.exists():
        raise ValueError('This job card has no printing entry yet: coating can only follow printing.')
    if record and not (post.get('change_reason') or '').strip():
        raise ValueError('Change reason is required when editing a coating record.')

    sides = 2 if (post.get('sides') == '2' or coats >= 2) else 1
    info = services.summary(job_card, exclude_run_id=record.pk if record else None, new_coats=coats, new_sides=sides)
    if sheets > info['sheet_cap']:
        raise ValueError(
            f'{sheets:,} sheets coated in one run is more than the {info["printed"]:,} good sheets printed for this job card.'
        )
    if sheets * coats > info['allowed_more']:
        raise ValueError(
            f'Coating passes would reach {info["passes"] + sheets * coats:,} (sheets x coats) but the limit is {info["pass_limit"]:,} '
            f'({info["printed"]:,} printed sheets x {info["sides"]} side{"s" if info["sides"] > 1 else ""}); '
            f'{info["allowed_more"]:,} can still be entered. If this job is coated on front and back, choose "Front and back" in "Coated on".'
        )

    start, end = _parse_time(post.get('start_time')), _parse_time(post.get('end_time'))
    run_minutes = float(post.get('run_minutes') or 0) or _minutes_between(start, end)
    machine = Machine.objects.filter(pk=post.get('machine')).first() if post.get('machine') else None
    operator = Operator.objects.filter(pk=post.get('operator'), is_active=True).first() if post.get('operator') else None
    values = {
        'job_card': job_card, 'date': post.get('date'), 'shift': shift, 'machine': machine, 'operator': operator,
        'coating_type': post.get('coating_type') or services.coating_kind(services.job_application(job_card)) or 'other',
        'coats': coats, 'sides': sides, 'sheets_coated': sheets, 'waste_sheets': waste, 'start_time': start, 'end_time': end,
        'run_minutes': run_minutes, 'status': post.get('status') or 'in_progress',
        'remark_notes': (post.get('remarks') or '').strip(),
    }
    if record:
        for name, value in values.items():
            setattr(record, name, value)
        record.change_reason = post.get('change_reason').strip()
        record.save()
        return record, False
    if job_card.workflow_status == 'released':
        start_production(job_card, actor=request.user, reason='Coating record created')
    return CoatingRun.objects.create(created_by=request.user, **values), True


@login_required
@permission_required('can_edit_production')
def coating_entry(request):
    view_id, edit_id = request.GET.get('view'), request.POST.get('edit_id') or request.GET.get('edit')
    record = get_object_or_404(CoatingRun, pk=view_id or edit_id, is_active=True) if (view_id or edit_id) else None
    is_view = bool(view_id)

    if request.method == 'POST' and not is_view:
        try:
            saved, created = _save_from_post(request, record)
            messages.success(request, f'Coating {"saved" if created else "updated"} for Job Card {saved.job_card.job_card_no}.')
            return redirect('coating_entry' if created else 'coating_records')
        except ValueError as exc:
            messages.error(request, str(exc))
        except Exception as exc:  # noqa: BLE001 - shown to the operator rather than a 500 page
            messages.error(request, f'Error saving coating: {exc}')

    show_all = request.GET.get('all') == '1'
    cards = list(services.coating_jobs(include_all=show_all).order_by('-created_at')[:300])
    prefill = str((request.POST.get('job_card') if request.method == 'POST' else request.GET.get('job_card')) or '').strip()
    wanted = record.job_card_id if record else (int(prefill) if prefill.isdigit() else None)
    if wanted and all(c.pk != wanted for c in cards):
        extra = JobCard.objects.filter(pk=wanted).select_related('planning_job', 'material').first()
        if extra:
            cards.insert(0, extra)
    info = {}
    for card in cards:
        data = services.summary(card, exclude_run_id=record.pk if record else None)
        data['kind'] = services.coating_kind(services.job_application(card)) or 'other'
        data['application'] = services.job_application(card) or '-'
        info[card.pk] = data
    return render(request, 'coating/entry.html', {
        'record': record, 'is_view': is_view, 'cards': cards, 'show_all': show_all,
        'info_json': json.dumps(info), 'prefill': wanted,
        'machines': Machine.objects.filter(is_active=True).order_by('name'),
        'operators': Operator.objects.filter(is_active=True).order_by('name'),
        'types': CoatingRun.TYPE_CHOICES, 'today': record.date if record else timezone.now().date(),
    })


@login_required
@permission_required('can_edit_production')
def coating_job_search(request):
    """Search printed job cards by job card no, SKU, PO, customer or job name (coating jobs first)."""
    query = (request.GET.get('q') or '').strip()
    if len(query) < 2:
        return JsonResponse({'results': []})
    match = (Q(job_card_no__icontains=query) | Q(SKU__icontains=query) | Q(PO_No__icontains=query)
             | Q(destination__icontains=query) | Q(planning_job__job_name__icontains=query))
    coating_ids = set(services.coating_jobs().values_list('pk', flat=True))
    cards = list(services.coating_jobs(include_all=True).filter(match).order_by('-created_at')[:40])
    cards.sort(key=lambda c: c.pk not in coating_ids)
    results = []
    for card in cards:
        data = services.summary(card)
        data['kind'] = services.coating_kind(services.job_application(card)) or 'other'
        data['application'] = services.job_application(card) or '-'
        results.append({
            'id': card.pk,
            'label': f'{card.job_card_no} - {card.SKU}' + (f' - {card.planning_job.job_name[:40]}' if card.planning_job else ''),
            'is_coating': card.pk in coating_ids, 'info': data,
        })
    return JsonResponse({'results': results})


@login_required
@permission_required('can_view_production_records')
def coating_records(request):
    can_edit = request.user.profile.can_edit_production()
    if request.method == 'POST':
        if not can_edit:
            messages.error(request, 'You do not have permission to remove coating records.')
            return redirect('coating_records')
        run = get_object_or_404(CoatingRun, pk=request.POST.get('run_id'), is_active=True)
        reason = (request.POST.get('reason') or '').strip()
        if not reason:
            messages.error(request, 'A reason is required to remove a coating record.')
        else:
            run.is_active, run.change_reason = False, f'Removed: {reason}'
            run.save(update_fields=['is_active', 'change_reason', 'updated_at'])
            messages.success(request, 'Coating record removed.')
        return redirect('coating_records')

    get = request.GET
    tab = 'pending' if get.get('tab') == 'pending' else 'records'
    today = timezone.now().date()
    date_from = get.get('from') or (today - timedelta(days=30)).isoformat()
    date_to = get.get('to') or today.isoformat()
    runs = CoatingRun.objects.filter(is_active=True, date__range=(date_from, date_to)).select_related('job_card', 'machine', 'operator')
    for key, field in (('machine', 'machine_id'), ('operator', 'operator_id'), ('type', 'coating_type'), ('shift', 'shift')):
        if get.get(key):
            runs = runs.filter(**{field: get[key]})
    if get.get('q'):
        runs = runs.filter(job_card__job_card_no__icontains=get['q'].strip())

    if get.get('export') == 'csv':
        response = HttpResponse(content_type='text/csv')
        response['Content-Disposition'] = 'attachment; filename="coating_records.csv"'
        out = csv.writer(response)
        out.writerow(['Date', 'Shift', 'Job Card', 'Coating', 'Machine', 'Operator', 'Coats', 'Sheets coated', 'Waste', 'Impressions', 'Run minutes', 'Status', 'Remarks'])
        for run in runs.order_by('date', 'id'):
            out.writerow([run.date, run.shift, run.job_card.job_card_no, run.get_coating_type_display(), run.machine or '', run.operator or '',
                          run.coats, run.sheets_coated, run.waste_sheets, run.impressions, run.run_minutes, run.get_status_display(), run.remark_notes])
        return response

    ctx = {
        'tab': tab, 'date_from': date_from, 'date_to': date_to, 'filters': get, 'can_edit': can_edit,
        'machines': Machine.objects.filter(is_active=True).order_by('name'),
        'operators': Operator.objects.filter(is_active=True).order_by('name'),
        'types': CoatingRun.TYPE_CHOICES,
    }
    if tab == 'pending':
        rows = services.pending_rows()
        ctx.update(pending=rows, pending_sheets=sum(r['remaining'] for r in rows))
    else:
        ctx.update(
            runs=list(runs.order_by('-date', '-id')[:500]),
            totals=runs.aggregate(sheets=Sum('sheets_coated'), waste=Sum('waste_sheets'), n=Count('pk'), minutes=Sum('run_minutes')),
            impressions=sum(r.impressions for r in runs),
            by_day=list(runs.values('date').annotate(sheets=Sum('sheets_coated'), waste=Sum('waste_sheets'), n=Count('pk')).order_by('-date')[:31]),
            by_operator=list(runs.values('operator__name').annotate(sheets=Sum('sheets_coated'), waste=Sum('waste_sheets'), minutes=Sum('run_minutes'), n=Count('pk')).order_by('-sheets')),
            by_machine=list(runs.values('machine__name').annotate(sheets=Sum('sheets_coated'), waste=Sum('waste_sheets'), minutes=Sum('run_minutes'), n=Count('pk')).order_by('-sheets')),
        )
    return render(request, 'coating/records.html', ctx)
