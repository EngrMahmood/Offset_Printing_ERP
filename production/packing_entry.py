"""Packing production entry — pieces packed + sorting waste."""

from __future__ import annotations

import json

from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.core.exceptions import ValidationError
from django.db import transaction
from django.db.models import Exists, OuterRef, Prefetch, Q
from django.http import JsonResponse
from django.shortcuts import get_object_or_404, redirect, render
from django.utils import timezone
from django.views.decorators.http import require_GET

from core.models import JOB_CARD_PRODUCTION_START_STATUSES, JobCard, Production, Sorter
from core.services import (
    build_audit_snapshot,
    ensure_edit_lock_allowed,
    get_active_record_or_404,
    get_record_edit_lock_days,
    log_change,
    record_is_time_locked,
)
from core.views import permission_required
from workflow.services import start_production


def _has_printing_entry_subquery():
    return Production.objects.filter(
        job_card_id=OuterRef('pk'),
        is_active=True,
        entry_type='printing',
    )


def _packing_eligible_job_cards_queryset(edit_record=None):
    has_printing = _has_printing_entry_subquery()
    if edit_record:
        qs = JobCard.objects.filter(is_active=True).filter(
            Q(pk=edit_record.job_card_id)
            | (
                Q(status__in=JOB_CARD_PRODUCTION_START_STATUSES)
                & (Q(is_print_job=False) | Exists(has_printing))
            )
        ).distinct()
    else:
        qs = JobCard.objects.filter(
            is_active=True,
            status__in=JOB_CARD_PRODUCTION_START_STATUSES,
        ).filter(
            Q(is_print_job=False) | Exists(has_printing),
        )
    return qs.select_related('planning_job', 'material').prefetch_related(
        Prefetch(
            'productions',
            queryset=Production.objects.filter(is_active=True).select_related('sorter', 'created_by'),
        )
    )


def _build_packing_job_info(job_card):
    packed = job_card.total_packed_pcs
    waste = job_card.total_sorting_waste_pcs
    used = job_card.total_packing_used_pcs
    limit = job_card.packing_limit_pcs
    history_qs = [row for row in job_card.productions.all() if row.entry_type == 'packing']
    history_qs.sort(key=lambda row: (row.date or timezone.now().date(), row.created_at), reverse=True)
    history = []
    for row in history_qs[:5]:
        history.append({
            'date': row.date.strftime('%d-%b') if row.date else '',
            'shift': row.shift,
            'packing_qty': f'{row.packing_qty:,}',
            'sorting_waste': f'{row.sorting_waste_qty:,}',
            'sorter': row.sorter.name if row.sorter else '-',
            'entered_by': row.created_by.get_full_name() if row.created_by else '-',
        })
    return {
        'job_card_no': job_card.job_card_no,
        'sku': job_card.SKU or '-',
        'product': job_card.planning_job.job_name if job_card.planning_job else (job_card.SKU or '-'),
        'customer': job_card.destination or '-',
        'po_no': job_card.PO_No or '-',
        'process_type': job_card.process_type_label,
        'is_print_job': job_card.is_print_job,
        'order_qty': f'{job_card.order_qty:,}',
        'produced_pcs': f'{job_card.total_printed_pcs:,}' if job_card.is_print_job else 'N/A',
        'pack_limit': f'{limit:,}',
        'already_packed': f'{packed:,}',
        'already_sort_waste': f'{waste:,}',
        'already_used': f'{used:,}',
        'remaining_allowed': f'{max(0, limit - used):,}',
        'remaining': max(0, limit - used),
        'remaining_display': f'{max(0, limit - used):,}',
        'dispatched': f'{job_card.total_dispatch:,}',
        'delivery_date': (
            job_card.planning_job.delivery_date.strftime('%Y-%m-%d')
            if job_card.planning_job and job_card.planning_job.delivery_date else '-'
        ),
        'material': job_card.material.name if job_card.material else '-',
        'destination': job_card.destination or '-',
        'ups': str(job_card.ups) if job_card.ups else '-',
        'history': history,
    }


def _build_packing_job_info_map(job_cards):
    return {str(job.id): _build_packing_job_info(job) for job in job_cards}


@login_required
@permission_required('can_edit_production')
@require_GET
def packing_job_card_search(request):
    """Server-side job card search for packing entry."""
    query = (request.GET.get('q') or '').strip()
    edit_id = (request.GET.get('edit_id') or '').strip()
    edit_record = get_active_record_or_404(Production, edit_id) if edit_id else None

    if len(query) < 2:
        return JsonResponse({'results': []})

    qs = _packing_eligible_job_cards_queryset(edit_record).filter(
        Q(job_card_no__icontains=query)
        | Q(SKU__icontains=query)
        | Q(PO_No__icontains=query)
        | Q(destination__icontains=query)
        | Q(planning_job__job_name__icontains=query)
    ).order_by('-created_at')[:30]

    results = []
    for job_card in qs:
        info = _build_packing_job_info(job_card)
        suffix = '' if job_card.is_print_job else ' (Cut & Pack)'
        results.append({
            'id': job_card.id,
            'label': f'{job_card.job_card_no} - {job_card.SKU}{suffix}',
            'job_card_no': job_card.job_card_no,
            'sku': info['sku'],
            'customer': info['customer'],
            'remaining_display': info['remaining_display'],
            'info': info,
        })
    return JsonResponse({'results': results})


@login_required
@permission_required('can_edit_production')
def packing_production_entry(request):
    view_id = (request.GET.get('view') or '').strip()
    is_view_mode = bool(view_id)
    edit_id = '' if is_view_mode else (request.POST.get('edit_id') or request.GET.get('edit') or '').strip()
    edit_record = (
        get_active_record_or_404(Production, view_id)
        if is_view_mode
        else (get_active_record_or_404(Production, edit_id) if edit_id else None)
    )
    if edit_record and edit_record.entry_type != 'packing':
        from django.urls import reverse
        if is_view_mode:
            return redirect(f"{reverse('printing_production_entry')}?view={edit_record.pk}")
        return redirect(f"{reverse('printing_production_entry')}?edit={edit_record.pk}")
    if edit_record and not is_view_mode and not ensure_edit_lock_allowed(request, 'production', edit_record):
        return redirect('production_records')

    if request.method == 'POST' and not is_view_mode:
        try:
            change_reason = (request.POST.get('change_reason') or '').strip()
            job_card = get_active_record_or_404(JobCard, request.POST.get('job_card'))
            if job_card.is_print_job and not job_card.printing_productions.exists():
                raise ValueError('Selected job card has no printing entry yet. Log printing production first.')
            packing_qty = int(request.POST.get('packing_qty') or 0)
            sorting_waste_qty = int(request.POST.get('sorting_waste_qty') or 0)
            shift = (request.POST.get('shift') or '').strip()
            packing_date = request.POST.get('date')
            sorter = get_object_or_404(Sorter, pk=request.POST.get('sorter'), is_active=True)
            remarks = (request.POST.get('remarks') or '').strip()

            if edit_record and not change_reason:
                raise ValueError('Change reason is required when editing packing data.')
            if not shift:
                raise ValueError('Shift is required.')
            if packing_qty < 0 or sorting_waste_qty < 0:
                raise ValueError('Quantities cannot be negative.')
            if packing_qty == 0 and sorting_waste_qty == 0:
                raise ValueError('Enter packing qty and/or sorting waste qty.')

            payload = {
                'entry_type': 'packing',
                'job_card': job_card,
                'machine': None,
                'operator': None,
                'shift': shift,
                'date': packing_date,
                'impressions': 0,
                'output_sheets': 0,
                'waste_sheets': 0,
                'intermediate_pass': False,
                'planned_time': 0,
                'run_time': 0,
                'make_ready_time': 0,
                'downtime_minutes': 0,
                'packing_qty': packing_qty,
                'sorting_waste_qty': sorting_waste_qty,
                'sorter': sorter,
                'remark_notes': remarks,
                'change_reason': change_reason,
                'status': request.POST.get('status') or 'in_progress',
            }

            if edit_record:
                before_snapshot = build_audit_snapshot('production', edit_record)
                with transaction.atomic():
                    for field_name, value in payload.items():
                        setattr(edit_record, field_name, value)
                    edit_record.save()
                if log_change('production', edit_record, before_snapshot, request.user, 'update', change_reason):
                    messages.success(request, f'Packing record updated for Job Card {job_card.job_card_no}')
                else:
                    messages.success(request, f'No changes detected for Job Card {job_card.job_card_no}')
                return redirect('production_records')

            with transaction.atomic():
                if job_card.workflow_status == 'released':
                    start_production(job_card, actor=request.user, reason='Packing record created')
                record = Production.objects.create(**payload)
                record.created_by = request.user
                record.save(update_fields=['created_by'])
            log_change('production', record, {}, request.user, 'create', 'Packing entry created')
            messages.success(request, f'Packing data saved for Job Card {job_card.job_card_no}')
            return redirect('packing_production_entry')

        except ValidationError as e:
            if hasattr(e, 'message_dict'):
                error_message = '; '.join(
                    f'{k}: {", ".join(str(m) for m in v)}' if isinstance(v, (list, tuple)) else str(v)
                    for k, v in e.message_dict.items()
                )
            else:
                error_message = ' '.join(str(m) for m in getattr(e, 'messages', [str(e)]))
            messages.error(request, f'Error saving packing data: {error_message}')
        except Exception as e:
            messages.error(request, f'Error saving packing data: {str(e)}')

    job_cards = list(_packing_eligible_job_cards_queryset(edit_record).order_by('-created_at')[:200])
    sorters = Sorter.objects.filter(is_active=True).order_by('name')
    info_map = _build_packing_job_info_map(job_cards)
    context = {
        'job_cards': job_cards,
        'sorters': sorters,
        'job_card_info_json': json.dumps(info_map),
        'today': edit_record.date if edit_record else timezone.now().date(),
        'edit_record': edit_record,
        'edit_lock_days': get_record_edit_lock_days(),
        'edit_lock_applies': bool(edit_record and record_is_time_locked('production', edit_record)),
        'is_view_mode': is_view_mode,
        'current_user_display': request.user.get_full_name() or request.user.username,
    }
    return render(request, 'production/packing_entry.html', context)


@login_required
@permission_required('can_edit_production')
def packing_records(request):
    """Packing production records ledger."""
    from datetime import datetime, timedelta
    from django.db.models import Sum, F
    from django.core.paginator import Paginator
    from core.models import EditOverrideRequest, Sorter
    from core.services import (
        get_record_edit_lock_days,
        get_record_edit_lock_cutoff,
        user_can_bypass_edit_lock,
        run_bulk_permanent_delete,
        user_can_archive_records,
    )
    from core.views import add_unique_message

    if request.method == 'POST':
        action = (request.POST.get('action') or '').strip()
        if action == 'bulk_delete':
            if request.user.profile.role != 'admin':
                add_unique_message(request, messages.ERROR, '❌ Only admin can run bulk delete.')
                return redirect('packing_records')
            if not user_can_archive_records(request.user):
                add_unique_message(request, messages.ERROR, '❌ You do not have permission to delete records.')
                return redirect('packing_records')

            selected_ids = request.POST.getlist('selected_ids')
            deleted_count, failures = run_bulk_permanent_delete(request, 'production', selected_ids)
            if deleted_count:
                add_unique_message(request, messages.SUCCESS, f'Deleted {deleted_count} packing record(s) permanently.')
            if failures:
                add_unique_message(request, messages.ERROR, f'Bulk delete completed with issues: {"; ".join(failures[:5])}')
            return redirect('packing_records')

    query = (request.GET.get('q') or '').strip()
    shift = (request.GET.get('shift') or '').strip()
    sorter_filter = (request.GET.get('sorter') or '').strip()
    date_from_raw = (request.GET.get('date_from') or '').strip()
    date_to_raw = (request.GET.get('date_to') or '').strip()
    sort = (request.GET.get('sort') or 'date').strip()
    direction = (request.GET.get('dir') or 'desc').strip().lower()
    per_page = request.GET.get('per_page') or '50'

    try:
        per_page = int(per_page)
    except (TypeError, ValueError):
        per_page = 50
    if per_page not in (50, 100):
        per_page = 50

    records = Production.objects.filter(
        is_active=True,
        job_card__is_active=True,
        entry_type='packing'
    ).select_related(
        'job_card', 'sorter', 'created_by'
    ).order_by('-date', '-id')

    if query:
        records = records.filter(
            Q(job_card__job_card_no__icontains=query) |
            Q(job_card__SKU__icontains=query) |
            Q(job_card__destination__icontains=query) |
            Q(sorter__name__icontains=query)
        )

    if shift:
        records = records.filter(shift=shift)

    if sorter_filter:
        try:
            records = records.filter(sorter_id=int(sorter_filter))
        except (TypeError, ValueError):
            sorter_filter = ''

    date_from = None
    date_to = None
    if date_from_raw:
        try:
            date_from = datetime.strptime(date_from_raw, '%Y-%m-%d').date()
            records = records.filter(date__gte=date_from)
        except ValueError:
            add_unique_message(request, messages.ERROR, 'Invalid From date format. Use YYYY-MM-DD.')
            date_from_raw = ''
    if date_to_raw:
        try:
            date_to = datetime.strptime(date_to_raw, '%Y-%m-%d').date()
            records = records.filter(date__lte=date_to)
        except ValueError:
            add_unique_message(request, messages.ERROR, 'Invalid To date format. Use YYYY-MM-DD.')
            date_to_raw = ''

    if date_from and date_to and date_from > date_to:
        add_unique_message(request, messages.ERROR, 'From date cannot be later than To date.')
        records = records.none()

    sortable_fields = {
        'date': 'date',
        'job_card': 'job_card__job_card_no',
        'sorter': 'sorter__name',
        'shift': 'shift',
        'packed_qty': 'packing_qty',
        'waste_qty': 'sorting_waste_qty',
        'added_by': 'created_by__username',
    }
    order_field = sortable_fields.get(sort, 'date')
    if direction not in ('asc', 'desc'):
        direction = 'desc'
    ordering = order_field if direction == 'asc' else f'-{order_field}'
    records = records.order_by(ordering)

    total_count = records.count()
    filtered_totals = records.aggregate(
        packed_total=Sum('packing_qty'),
        waste_total=Sum('sorting_waste_qty'),
    )

    paginator = Paginator(records, per_page)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)

    records_list = list(page_obj.object_list)
    page_packed_total = sum(row.packing_qty or 0 for row in records_list)
    page_waste_total = sum(row.sorting_waste_qty or 0 for row in records_list)

    cutoff = get_record_edit_lock_cutoff()
    pending_ids = set()
    approved_ids = set()
    if cutoff and not user_can_bypass_edit_lock(request.user):
        user_overrides = EditOverrideRequest.objects.filter(
            entity_type='production',
            requested_by=request.user,
        ).values('record_id', 'status', 'expires_at')
        for ov in user_overrides:
            if ov['status'] == 'pending':
                pending_ids.add(ov['record_id'])
            elif ov['status'] == 'approved' and ov['expires_at'] and ov['expires_at'] > timezone.now():
                approved_ids.add(ov['record_id'])

    context = {
        'edit_lock_days': get_record_edit_lock_days(),
        'edit_lock_cutoff': cutoff,
        'can_bypass_edit_lock': user_can_bypass_edit_lock(request.user),
        'pending_override_ids': pending_ids,
        'approved_override_ids': approved_ids,
        'records': records_list,
        'page_obj': page_obj,
        'total_count': total_count,
        'filtered_packed_total': filtered_totals['packed_total'] or 0,
        'filtered_waste_total': filtered_totals['waste_total'] or 0,
        'page_packed_total': page_packed_total,
        'page_waste_total': page_waste_total,
        'has_active_filters': bool(
            query or shift or sorter_filter or date_from_raw or date_to_raw
        ),
        'sorters': Sorter.objects.filter(is_active=True).order_by('name'),
        'q': query,
        'shift': shift,
        'sorter': sorter_filter,
        'today': timezone.now().date().isoformat(),
        'week_start': (timezone.now().date() - timedelta(days=timezone.now().weekday())).isoformat(),
        'month_start': timezone.now().date().replace(day=1).isoformat(),
        'date_from': date_from_raw,
        'date_to': date_to_raw,
        'sort': sort,
        'dir': direction,
        'per_page': per_page,
    }
    return render(request, 'production/packing_records.html', context)
