from django.shortcuts import render, get_object_or_404, redirect
from django.http import JsonResponse, Http404, HttpResponse
from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.utils import timezone
from django.db.models import Q, Sum, Count, F
from django.db import transaction
from django.core.exceptions import ValidationError
from django.core.paginator import Paginator
from django.urls import reverse
from django.views.decorators.http import require_GET
import json
import re
from datetime import datetime, timedelta
from django.contrib.auth import get_user_model

from core.models import (
    Dispatch,
    EditOverrideRequest,
    JOB_CARD_PRODUCTION_START_STATUSES,
    JobCard,
    JobCardWipStatus,
    Machine,
    Operator,
    Production,
    ProductionDowntime,
    ProductionWipStatus,
    Supervisor,
    Sorter,
)
from core.views import permission_required
from core.services import (
    add_unique_message,
    build_audit_snapshot,
    compute_planned_minutes,
    ensure_edit_lock_allowed,
    get_active_record_or_404,
    get_record_edit_lock_cutoff,
    get_record_edit_lock_days,
    log_change,
    record_is_time_locked,
    run_bulk_permanent_delete,
    user_can_archive_records,
    user_can_bypass_edit_lock,
)
from workflow.services import start_production
from production.services import OEECalculator
from production.printing_entry_helpers import (
    build_printing_job_card_maps,
    get_effective_job_card_plan,
    get_remaining_planned_for_job_card,
    printing_job_cards_queryset,
    resolve_related_machine,
)
from production.printing_pass_helpers import (
    effective_print_pass_number,
    get_job_card_pass_count,
    get_max_print_passes,
    validate_print_pass_number,
)

# Resolve the active user model
User = get_user_model()


@login_required
def create_operator_ajax(request):
    if request.method != 'POST':
        return JsonResponse({'error': 'invalid_method'}, status=405)
    if not getattr(request.user, 'profile', None) or not request.user.profile.can_manage_masters():
        return JsonResponse({'error': 'forbidden'}, status=403)
    name = (request.POST.get('name') or '').strip()
    emp = (request.POST.get('employee_code') or '').strip()
    if not name:
        return JsonResponse({'error': 'name_required'}, status=400)
    op = None
    try:
        op = Operator.objects.create(name=name, employee_code=emp)
    except Exception as e:
        return JsonResponse({'error': str(e)}, status=500)
    return JsonResponse({'id': op.id, 'name': op.name})


@login_required
def create_machine_ajax(request):
    if request.method != 'POST':
        return JsonResponse({'error': 'invalid_method'}, status=405)
    if not getattr(request.user, 'profile', None) or not request.user.profile.can_manage_masters():
        return JsonResponse({'error': 'forbidden'}, status=403)
    name = (request.POST.get('name') or '').strip()
    sip = request.POST.get('standard_impressions_per_hour') or ''
    setup = request.POST.get('standard_setup_minutes_per_color') or ''
    try:
        sip_val = float(sip) if sip else 0
    except ValueError:
        return JsonResponse({'error': 'invalid_speed'}, status=400)
    try:
        setup_val = float(setup) if setup else 0
    except ValueError:
        return JsonResponse({'error': 'invalid_setup_minutes'}, status=400)
    if not name:
        return JsonResponse({'error': 'name_required'}, status=400)
    standard_speed = sip_val if sip else 4000
    setup_minutes = setup_val if setup else 15
    if standard_speed <= 0:
        return JsonResponse({'error': 'speed_must_be_positive'}, status=400)
    if setup_minutes < 0:
        return JsonResponse({'error': 'setup_minutes_must_be_non_negative'}, status=400)
    try:
        m = Machine.objects.create(
            name=name,
            standard_impressions_per_hour=standard_speed,
            standard_setup_minutes_per_color=setup_minutes,
        )
    except Exception as e:
        return JsonResponse({'error': str(e)}, status=500)
    return JsonResponse({
        'id': m.id,
        'name': m.name,
        'standard_impressions_per_hour': m.standard_impressions_per_hour,
        'standard_setup_minutes_per_color': m.standard_setup_minutes_per_color,
    })


@login_required
def create_supervisor_ajax(request):
    if request.method != 'POST':
        return JsonResponse({'error': 'invalid_method'}, status=405)
    if not getattr(request.user, 'profile', None) or not request.user.profile.can_manage_masters():
        return JsonResponse({'error': 'forbidden'}, status=403)
    name = (request.POST.get('name') or '').strip()
    emp = (request.POST.get('employee_code') or '').strip()
    if not name:
        return JsonResponse({'error': 'name_required'}, status=400)

    existing = None
    if emp:
        existing = Supervisor.objects.filter(employee_code__iexact=emp).first()
    if existing is None:
        existing = Supervisor.objects.filter(
            name__iexact=name,
            employee_code__iexact=emp or ''
        ).first() if emp else Supervisor.objects.filter(name__iexact=name, employee_code__isnull=True).first()

    if existing:
        display_name = f"{existing.name} ({existing.employee_code})" if existing.employee_code else existing.name
        return JsonResponse({
            'id': existing.id,
            'name': existing.name,
            'display_name': display_name,
            'employee_code': existing.employee_code,
            'created': False,
        })

    try:
        sup = Supervisor.objects.create(name=name, employee_code=emp or None)
    except Exception as e:
        return JsonResponse({'error': str(e)}, status=500)

    display_name = f"{sup.name} ({sup.employee_code})" if sup.employee_code else sup.name
    return JsonResponse({
        'id': sup.id,
        'name': sup.name,
        'display_name': display_name,
        'employee_code': sup.employee_code,
        'created': True,
    })


@login_required
def create_sorter_ajax(request):
    if request.method != 'POST':
        return JsonResponse({'error': 'invalid_method'}, status=405)
    if not getattr(request.user, 'profile', None) or not request.user.profile.can_manage_masters():
        return JsonResponse({'error': 'forbidden'}, status=403)
    name = (request.POST.get('name') or '').strip()
    emp = (request.POST.get('employee_code') or '').strip()
    if not name:
        return JsonResponse({'error': 'name_required'}, status=400)
    try:
        sorter = Sorter.objects.create(name=name, employee_code=emp or None)
    except Exception as e:
        return JsonResponse({'error': str(e)}, status=500)
    display_name = f'{sorter.name} ({sorter.employee_code})' if sorter.employee_code else sorter.name
    return JsonResponse({'id': sorter.id, 'name': sorter.name, 'display_name': display_name})


@login_required
@permission_required('can_edit_production')
def production_entry(request):
    """Backward-compatible alias for printing production entry."""
    return printing_production_entry(request)


@login_required
@permission_required('can_edit_production')
def printing_production_entry(request):
    """Production data entry form for operators"""
    view_id = (request.GET.get('view') or '').strip()
    is_view_mode = bool(view_id)
    edit_id = '' if is_view_mode else (request.POST.get('edit_id') or request.GET.get('edit') or '').strip()
    edit_record = get_active_record_or_404(Production, view_id) if is_view_mode else (get_active_record_or_404(Production, edit_id) if edit_id else None)
    if edit_record and edit_record.entry_type == 'packing':
        target = 'packing_production_entry'
        if is_view_mode:
            return redirect(f'{reverse(target)}?view={edit_record.pk}')
        if edit_id:
            return redirect(f'{reverse(target)}?edit={edit_record.pk}')
    if edit_record and not is_view_mode and not ensure_edit_lock_allowed(request, 'production', edit_record):
        return redirect('production_records')

    if request.method == 'POST' and not is_view_mode:
        job_card_id = request.POST.get('job_card')
        machine_id = request.POST.get('machine')
        machine_override = request.POST.get('machine_override') == 'on'
        operator_id = request.POST.get('operator')
        shift = request.POST.get('shift')
        date = request.POST.get('date')
        impressions = request.POST.get('impressions')
        output_sheets = request.POST.get('output_sheets')
        waste_sheets = request.POST.get('waste_sheets')
        print_pass_number_raw = (request.POST.get('print_pass_number') or '').strip()
        downtime_minutes = request.POST.get('downtime_minutes')
        make_ready_time = request.POST.get('make_ready_time')
        planned_time = request.POST.get('planned_time')
        run_time = request.POST.get('run_time')
        downtime_category = request.POST.get('downtime_category')
        downtime_categories = request.POST.getlist('downtime_category[]')
        downtime_minutes_rows = request.POST.getlist('downtime_minutes_detail[]')
        downtime_notes = request.POST.getlist('downtime_note[]')
        waste_reason = request.POST.get('waste_reason')
        remarks = request.POST.get('remarks')
        overrun_reason_select = (request.POST.get('overrun_reason_select') or '').strip()
        overrun_reason_other = (request.POST.get('overrun_reason_other') or '').strip()

        try:
            change_reason = (request.POST.get('change_reason') or '').strip()
            output_sheets_val = int(output_sheets) if output_sheets else 0
            waste_sheets_val = int(waste_sheets) if waste_sheets else 0
            legacy_downtime_val = float(downtime_minutes) if downtime_minutes else 0
            planned_time_val = float(planned_time) if planned_time else 0
            # run_time_val calculated below from timestamps
            make_ready_time_val = float(make_ready_time) if make_ready_time else 0
            impressions_val = int(impressions) if impressions else 0

            # New fields
            counter_start_val = int(request.POST.get('counter_start') or 0)
            counter_end_val = int(request.POST.get('counter_end') or 0)
            start_time_val = request.POST.get('start_time') or None
            end_time_val = request.POST.get('end_time') or None
            supervisor_id = request.POST.get('supervisor') or None
            production_status = request.POST.get('status') or 'in_progress'
            waste_reason_other = request.POST.get('waste_reason_other') or ''
            downtime_category_other = request.POST.get('downtime_category_other') or ''

            # Server-side calculations
            # Impressions from counters if provided
            if counter_end_val > counter_start_val:
                impressions_val = counter_end_val - counter_start_val

            # Runtime from timestamps
            run_time_val = float(run_time) if run_time else 0
            if start_time_val and end_time_val:
                from datetime import datetime
                t1 = datetime.strptime(start_time_val, '%H:%M')
                t2 = datetime.strptime(end_time_val, '%H:%M')
                tdelta = t2 - t1
                if tdelta.total_seconds() > 0:
                    run_time_val = tdelta.total_seconds() / 60.0 # in minutes
                elif tdelta.total_seconds() < 0:
                    raise ValueError('End time cannot be earlier than start time')

            downtime_entries = []
            max_rows = max(len(downtime_categories), len(downtime_minutes_rows), len(downtime_notes))
            for idx in range(max_rows):
                category = (downtime_categories[idx] if idx < len(downtime_categories) else '').strip()
                minute_raw = (downtime_minutes_rows[idx] if idx < len(downtime_minutes_rows) else '').strip()
                note = (downtime_notes[idx] if idx < len(downtime_notes) else '').strip()

                if not category and not minute_raw and not note:
                    continue

                if not category:
                    raise ValueError('Downtime category is required for each downtime row')

                try:
                    minutes = float(minute_raw or 0)
                except ValueError:
                    raise ValueError('Downtime minutes must be numeric')

                if minutes <= 0:
                    raise ValueError('Downtime minutes must be greater than 0 for each downtime row')

                downtime_entries.append({
                    'category': category,
                    'minutes': minutes,
                    'note': note or None,
                })

            downtime_val = float(sum(item['minutes'] for item in downtime_entries)) if downtime_entries else legacy_downtime_val
            primary_downtime_category = downtime_entries[0]['category'] if downtime_entries else downtime_category

            job_card = get_active_record_or_404(JobCard, job_card_id)
            try:
                print_pass_number = int(print_pass_number_raw or 1)
            except (TypeError, ValueError):
                raise ValueError('Print pass is required.')
            pass_meta = validate_print_pass_number(
                job_card,
                print_pass_number,
                exclude_production_id=edit_record.pk if edit_record else None,
            )
            intermediate_pass = pass_meta['intermediate_pass']
            if intermediate_pass:
                output_sheets_val = 0

            if edit_record and not change_reason:
                raise ValueError('Change reason is required when editing production data')
            if planned_time_val < 0:
                raise ValueError('Planned time must be greater than or equal to 0')
            if run_time_val < 0:
                raise ValueError('Run time must be greater than or equal to 0')
            if make_ready_time_val < 0:
                raise ValueError('Make ready time must be greater than or equal to 0')
            if downtime_val < 0:
                raise ValueError('Downtime minutes must be greater than or equal to 0')
            if impressions_val < 0:
                raise ValueError('Impressions must be greater than or equal to 0')
            if output_sheets_val < 0:
                raise ValueError('Good sheets produced cannot be negative')
            if waste_sheets_val < 0:
                raise ValueError('Waste sheets cannot be negative')
            if not intermediate_pass and output_sheets_val <= 0:
                raise ValueError('Good sheets must be greater than 0 for the final print pass')
            if downtime_val > 0 and not primary_downtime_category:
                raise ValueError('Downtime category is required when downtime is greater than 0')
            if waste_sheets_val > 0 and not waste_reason:
                raise ValueError('Waste reason is required when waste sheets are greater than 0')
            if impressions_val <= 0:
                raise ValueError('Impressions must be greater than 0')
            if waste_sheets_val < 0:
                raise ValueError('Waste sheets cannot be negative')

            if machine_override and machine_id:
                machine = get_object_or_404(Machine, pk=machine_id)
            elif job_card.machine_name_id:
                machine = job_card.machine_name
            elif job_card.machine_name_display:
                machine = resolve_related_machine(job_card)
                if not machine:
                    raise ValueError('No machine mapped on selected Job Card. Please set machine in Job Card or choose fallback machine.')
            elif machine_id:
                machine = get_object_or_404(Machine, pk=machine_id)
            elif edit_record and edit_record.machine_id:
                machine = edit_record.machine
            else:
                raise ValueError('No machine mapped on selected Job Card. Please set machine in Job Card or choose fallback machine.')
            operator = get_object_or_404(Operator, pk=operator_id)

            job_card_plan = get_effective_job_card_plan(job_card)
            remaining_planned = get_remaining_planned_for_job_card(
                job_card,
                job_card_plan['planned_total'],
                exclude_production_id=edit_record.pk if edit_record else None,
            )

            if job_card_plan['planned_total'] > 0:
                if edit_record:
                    planned_time_val = float(edit_record.planned_time or 0)
                else:
                    if remaining_planned > 0:
                        planned_time_val = remaining_planned
                    elif planned_time_val <= 0:
                        planned_time_val = 0

            overrun_reason_map = {
                'extra_setup': 'Extra Setup / Make Ready',
                'machine_slowdown': 'Machine Slowdown',
                'operator_learning': 'Operator Learning / New Team',
                'material_issue': 'Material-related Delay',
                'quality_rework': 'Quality Rework',
                'job_complexity': 'Higher Job Complexity',
                'other': 'Other',
            }
            overrun_reason = ''
            if overrun_reason_select == 'other':
                overrun_reason = overrun_reason_other
            elif overrun_reason_select in overrun_reason_map:
                overrun_reason = overrun_reason_map[overrun_reason_select]

            if not edit_record and planned_time_val > remaining_planned:
                if not overrun_reason_select:
                    raise ValueError(
                        'Overrun allocation reason is required when planned minutes exceed remaining planned time.'
                    )
                if overrun_reason_select == 'other' and not overrun_reason_other:
                    raise ValueError('Please specify overrun reason details for "Other".')

            payload = {
                'entry_type': 'printing',
                'job_card': job_card,
                'machine': machine,
                'operator': operator,
                'shift': shift,
                'date': date,
                'impressions': impressions_val,
                'output_sheets': output_sheets_val,
                'waste_sheets': waste_sheets_val,
                'intermediate_pass': intermediate_pass,
                'print_pass_number': print_pass_number,
                'planned_time': planned_time_val,
                'run_time': run_time_val,
                'make_ready_time': make_ready_time_val,
                'downtime_minutes': downtime_val,
                'downtime_category': primary_downtime_category,
                'waste_reason': waste_reason,
                'counter_start': counter_start_val,
                'counter_end': counter_end_val,
                'start_time': start_time_val,
                'end_time': end_time_val,
                'supervisor_id': supervisor_id,
                'status': production_status,
                'waste_reason_other': waste_reason_other,
                'downtime_category_other': downtime_category_other,
                'remark_notes': remarks,
                'change_reason': change_reason,
            }

            if edit_record:
                before_snapshot = build_audit_snapshot('production', edit_record)
                with transaction.atomic():
                    for field_name, value in payload.items():
                        setattr(edit_record, field_name, value)
                    edit_record.save()
                    edit_record.downtime_entries.all().delete()
                    if downtime_entries:
                        ProductionDowntime.objects.bulk_create([
                            ProductionDowntime(
                                production=edit_record,
                                category=item['category'],
                                minutes=item['minutes'],
                                note=item['note'],
                            )
                            for item in downtime_entries
                        ])

                if log_change('production', edit_record, before_snapshot, request.user, 'update', change_reason):
                    messages.success(request, f'Production record updated for Job Card {job_card.job_card_no}')
                else:
                    messages.success(request, f'No changes detected for Job Card {job_card.job_card_no}')
                # Consume any approval this edit was made under, so it drops off
                # the requester's "My Override Requests" list.
                EditOverrideRequest.objects.filter(
                    entity_type='production',
                    record_id=edit_record.pk,
                    requested_by=request.user,
                    status='approved',
                    consumed_at__isnull=True,
                    expires_at__gt=timezone.now(),
                ).update(consumed_at=timezone.now())
                return redirect('production_records')

            with transaction.atomic():
                if job_card.workflow_status == 'released':
                    start_production(job_card, actor=request.user, reason='Production record created')
                record = Production.objects.create(**payload)
                if downtime_entries:
                    ProductionDowntime.objects.bulk_create([
                        ProductionDowntime(
                            production=record,
                            category=item['category'],
                            minutes=item['minutes'],
                            note=item['note'],
                        )
                        for item in downtime_entries
                    ])
                record.created_by = request.user
                record.save(update_fields=['created_by'])
            create_reason = overrun_reason or 'Initial entry created'
            log_change('production', record, {}, request.user, 'create', create_reason)

            messages.success(request, f'Production data saved successfully for Job Card {job_card.job_card_no}')
            return redirect('printing_production_entry')

        except ValidationError as e:
            if hasattr(e, 'message_dict'):
                error_parts = []
                for field_name, field_messages in e.message_dict.items():
                    if isinstance(field_messages, (list, tuple)):
                        error_parts.append(f"{field_name}: {', '.join(str(msg) for msg in field_messages)}")
                    else:
                        error_parts.append(str(field_messages))
                error_message = '; '.join(error_parts)
            else:
                error_message = ' '.join(str(msg) for msg in getattr(e, 'messages', [str(e)]))
            messages.error(request, f'Error saving production data: {error_message}')
        except Exception as e:
            messages.error(request, f'Error saving production data: {str(e)}')

    job_cards = list(printing_job_cards_queryset(edit_record)[:200])
    # The queryset is ordered by -created_at and sliced to the 200 most recent
    # cards, which can drop an *old* job card being edited (e.g. legacy records).
    # Guarantee the edited record's job card is present so it pre-selects and its
    # summary/pass data load.
    if edit_record and edit_record.job_card_id and not any(
        jc.id == edit_record.job_card_id for jc in job_cards
    ):
        edited_jc = JobCard.objects.filter(pk=edit_record.job_card_id).first()
        if edited_jc:
            job_cards.insert(0, edited_jc)
    machines = Machine.objects.filter(is_active=True)
    operators = Operator.objects.all()
    supervisors = Supervisor.objects.filter(is_active=True).order_by('name')

    job_card_plan_map, job_card_machine_map, job_card_info_map = build_printing_job_card_maps(
        job_cards,
        edit_record=edit_record,
    )

    context = {
        'job_cards': job_cards,
        'machines': machines,
        'operators': operators,
        'supervisors': supervisors,
        'job_card_plan_json': json.dumps(job_card_plan_map),
        'job_card_machine_json': json.dumps(job_card_machine_map),
        'job_card_info_json': json.dumps(job_card_info_map),
        'today': edit_record.date if edit_record else timezone.now().date(),
        'edit_record': edit_record,
        'edit_downtime_rows_json': json.dumps([
            {
                'category': row.category,
                'minutes': float(row.minutes or 0),
                'note': row.note or '',
            }
            for row in (edit_record.downtime_entries.all() if edit_record else [])
        ]),
        'edit_lock_days': get_record_edit_lock_days(),
        'edit_lock_applies': bool(edit_record and record_is_time_locked('production', edit_record)),
        'is_view_mode': is_view_mode,
        'max_print_passes': get_max_print_passes(),
    }

    return render(request, 'production/production_entry.html', context)


@login_required
@permission_required('can_edit_production')
@require_GET
def printing_job_card_search(request):
    """Server-side job card search for printing production entry."""
    query = (request.GET.get('q') or '').strip()
    edit_id = (request.GET.get('edit_id') or '').strip()
    edit_record = get_active_record_or_404(Production, edit_id) if edit_id else None

    if len(query) < 2:
        return JsonResponse({'results': []})

    qs = printing_job_cards_queryset(edit_record).filter(
        Q(job_card_no__icontains=query)
        | Q(SKU__icontains=query)
        | Q(PO_No__icontains=query)
        | Q(destination__icontains=query)
        | Q(planning_job__job_name__icontains=query)
    ).select_related('planning_job', 'material', 'machine_name')[:60]

    plan_map, machine_map, info_map = build_printing_job_card_maps(list(qs), edit_record=edit_record)
    results = []
    for job_card in qs:
        job_id = str(job_card.id)
        info = info_map.get(job_id, {})
        results.append({
            'id': job_card.id,
            'label': f'{job_card.job_card_no} - {job_card.SKU}',
            'job_card_no': job_card.job_card_no,
            'sku': job_card.SKU or '-',
            'customer': info.get('customer', job_card.destination or '-'),
            'remaining_display': info.get('remaining_qty', '-'),
            'info': info,
            'machine': machine_map.get(job_id),
            'plan': plan_map.get(job_id),
        })

    completed_matches = []
    if not results:
        from core.services import find_completed_job_card_matches
        completed_matches = find_completed_job_card_matches(query)

    return JsonResponse({'results': results, 'completed_matches': completed_matches})


@login_required
@permission_required('can_view_production_records')
def production_records(request):
    """Production records list page"""
    if request.method == 'POST':
        action = (request.POST.get('action') or '').strip()
        if action == 'bulk_delete':
            if not request.user.is_superuser:
                add_unique_message(request, messages.ERROR, '❌ Only a superuser can run bulk delete.')
                return redirect('production_records')

            selected_ids = request.POST.getlist('selected_ids')
            deleted_count, failures = run_bulk_permanent_delete(request, 'production', selected_ids)
            if deleted_count:
                add_unique_message(request, messages.SUCCESS, f'Deleted {deleted_count} production record(s) permanently.')
            if failures:
                add_unique_message(request, messages.ERROR, f'Bulk delete completed with issues: {"; ".join(failures[:5])}')
            return redirect('production_records')

    query = (request.GET.get('q') or '').strip()
    shift = (request.GET.get('shift') or '').strip()
    machine_filter = (request.GET.get('machine') or '').strip()
    performance_status = (request.GET.get('performance_status') or '').strip().lower()
    date_from_raw = (request.GET.get('date_from') or '').strip()
    date_to_raw = (request.GET.get('date_to') or '').strip()
    logged_from_raw = (request.GET.get('logged_from') or '').strip()
    logged_to_raw = (request.GET.get('logged_to') or '').strip()
    sort = (request.GET.get('sort') or 'date').strip()
    direction = (request.GET.get('dir') or 'desc').strip().lower()
    per_page = request.GET.get('per_page') or '50'
    try:
        per_page = int(per_page)
    except (TypeError, ValueError):
        per_page = 50
    if per_page not in (50, 100):
        per_page = 50

    records = Production.objects.filter(is_active=True, job_card__is_active=True, entry_type='printing').select_related(
        'job_card', 'machine', 'operator', 'supervisor', 'sorter', 'created_by'
    ).order_by('-date', '-id')

    if query:
        records = records.filter(
            Q(job_card__job_card_no__icontains=query) |
            Q(job_card__SKU__icontains=query) |
            Q(job_card__destination__icontains=query) |
            Q(machine__name__icontains=query) |
            Q(operator__name__icontains=query) |
            Q(supervisor__name__icontains=query)
        )

    if shift:
        records = records.filter(shift=shift)

    if machine_filter:
        try:
            records = records.filter(machine_id=int(machine_filter))
        except (TypeError, ValueError):
            machine_filter = ''

    records = records.annotate(
        actual_total_minutes=F('run_time') + F('downtime_minutes') + F('make_ready_time'),
    )
    if performance_status == 'overrun':
        records = records.filter(actual_total_minutes__gt=F('planned_time'))
    elif performance_status == 'under_plan':
        records = records.filter(actual_total_minutes__lt=F('planned_time'))
    elif performance_status == 'on_plan':
        records = records.filter(actual_total_minutes=F('planned_time'))

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

    logged_from = None
    logged_to = None
    if logged_from_raw:
        try:
            logged_from = datetime.strptime(logged_from_raw, '%Y-%m-%d').date()
            records = records.filter(created_at__date__gte=logged_from)
        except ValueError:
            add_unique_message(request, messages.ERROR, 'Invalid Logged From date format. Use YYYY-MM-DD.')
            logged_from_raw = ''
    if logged_to_raw:
        try:
            logged_to = datetime.strptime(logged_to_raw, '%Y-%m-%d').date()
            records = records.filter(created_at__date__lte=logged_to)
        except ValueError:
            add_unique_message(request, messages.ERROR, 'Invalid Logged To date format. Use YYYY-MM-DD.')
            logged_to_raw = ''

    if logged_from and logged_to and logged_from > logged_to:
        add_unique_message(request, messages.ERROR, 'Logged From date cannot be later than Logged To date.')
        records = records.none()

    sortable_fields = {
        'date': 'date',
        'job_card': 'job_card__job_card_no',
        'machine': 'machine__name',
        'operator': 'operator__name',
        'shift': 'shift',
        'impressions': 'impressions',
        'output': 'output_sheets',
        'waste': 'waste_sheets',
        'planned': 'planned_time',
        'added_by': 'created_by__username',
    }
    order_field = sortable_fields.get(sort, 'date')
    if direction not in ('asc', 'desc'):
        direction = 'desc'
    ordering = order_field if direction == 'asc' else f'-{order_field}'
    records = records.order_by(ordering)

    total_count = records.count()
    filtered_totals = records.aggregate(
        impressions_total=Sum('impressions'),
        output_total=Sum('output_sheets'),
        waste_total=Sum('waste_sheets'),
    )
    paginator = Paginator(records, per_page)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)

    records = list(page_obj.object_list)
    page_impressions_total = sum(row.impressions or 0 for row in records)
    page_output_total = sum(row.output_sheets or 0 for row in records)
    page_waste_total = sum(row.waste_sheets or 0 for row in records)
    page_overrun_count = sum(1 for row in records if row.overrun_minutes > 0)
    page_tolerance_alert_count = 0
    job_card_ids = list({row.job_card_id for row in records})
    consumption_map = {
        item['job_card_id']: int((item['total_output'] or 0) + (item['total_waste'] or 0))
        for item in Production.objects.filter(is_active=True, job_card_id__in=job_card_ids).values('job_card_id').annotate(
            total_output=Sum('output_sheets'),
            total_waste=Sum('waste_sheets'),
        )
    }
    for row in records:
        consumed = consumption_map.get(row.job_card_id, 0)
        row.job_card_extra_sheets_used = max(consumed - row.job_card.total_sheets_planned, 0)
        row.job_card_tolerance_sheets = row.job_card.tolerance_sheets
        if row.job_card_extra_sheets_used > row.job_card_tolerance_sheets:
            page_tolerance_alert_count += 1

    cutoff = get_record_edit_lock_cutoff()
    pending_ids: set = set()
    approved_ids: set = set()
    if cutoff and not user_can_bypass_edit_lock(request.user):
        user_overrides = EditOverrideRequest.objects.filter(
            entity_type='production',
            requested_by=request.user,
        ).values('record_id', 'status', 'expires_at', 'consumed_at')
        for ov in user_overrides:
            if ov['status'] == 'pending':
                pending_ids.add(ov['record_id'])
            elif (
                ov['status'] == 'approved'
                and ov['consumed_at'] is None
                and ov['expires_at'] and ov['expires_at'] > timezone.now()
            ):
                approved_ids.add(ov['record_id'])

    context = {
        'edit_lock_days': get_record_edit_lock_days(),
        'edit_lock_cutoff': cutoff,
        'can_bypass_edit_lock': user_can_bypass_edit_lock(request.user),
        'pending_override_ids': pending_ids,
        'approved_override_ids': approved_ids,
        'records': records,
        'page_obj': page_obj,
        'total_count': total_count,
        'filtered_impressions_total': filtered_totals['impressions_total'] or 0,
        'filtered_output_total': filtered_totals['output_total'] or 0,
        'filtered_waste_total': filtered_totals['waste_total'] or 0,
        'page_impressions_total': page_impressions_total,
        'page_output_total': page_output_total,
        'page_waste_total': page_waste_total,
        'page_overrun_count': page_overrun_count,
        'page_tolerance_alert_count': page_tolerance_alert_count,
        'has_active_filters': bool(
            query or shift or machine_filter or performance_status
            or date_from_raw or date_to_raw or logged_from_raw or logged_to_raw
        ),
        'machines': Machine.objects.filter(is_active=True).order_by('name'),
        'q': query,
        'shift': shift,
        'machine': machine_filter,
        'performance_status': performance_status,
        'today': timezone.now().date().isoformat(),
        'week_start': (timezone.now().date() - timedelta(days=timezone.now().weekday())).isoformat(),
        'month_start': timezone.now().date().replace(day=1).isoformat(),
        'date_from': date_from_raw,
        'date_to': date_to_raw,
        'logged_from': logged_from_raw,
        'logged_to': logged_to_raw,
        'sort': sort,
        'dir': direction,
        'per_page': per_page,
    }
    return render(request, 'production/production_records.html', context)


def user_can_set_pass_override(user):
    """Supervisory roles allowed to set a per-job pass-count override.

    Mirrors the `nav.can_set_pass_override` gate (core.navigation), both of
    which defer to the same soft-coded action.set_pass_override permission.
    """
    if getattr(user, 'is_staff', False):
        return True
    profile = getattr(user, 'profile', None)
    return bool(profile and profile.can_set_pass_override())


@login_required
def set_pass_override(request):
    """Authorised supervisor override of a job's print pass count.

    Used when a machine runs at reduced colour capacity (a unit under
    maintenance) so a job planned for e.g. 2 passes must physically run in 4.
    Requires a supervisory role, a reason, and is audit-logged. Setting the
    override also lifts the impression ceiling proportionally (see JobCard).
    """
    if request.method != 'POST':
        return redirect('printing_production_entry')
    if not user_can_set_pass_override(request.user):
        messages.error(request, '❌ You do not have permission to override the pass count.')
        return redirect('printing_production_entry')

    job_card = get_active_record_or_404(JobCard, request.POST.get('job_card_id'))
    reason = (request.POST.get('reason') or '').strip()
    raw = (request.POST.get('pass_count') or '').strip()

    before_snapshot = build_audit_snapshot('job_card', job_card)

    if not raw:
        # Clearing the override (revert to planned passes).
        if job_card.pass_count_override:
            if not reason:
                messages.error(request, 'A reason is required to clear the pass override.')
                return redirect('printing_production_entry')
            job_card.pass_count_override = None
            job_card.pass_count_override_reason = None
            job_card.pass_count_override_by = None
            job_card.pass_count_override_at = None
            job_card.save(update_fields=[
                'pass_count_override', 'pass_count_override_reason',
                'pass_count_override_by', 'pass_count_override_at',
            ])
            log_change('job_card', job_card, before_snapshot, request.user, 'update', reason)
            messages.success(request, f'Pass override cleared for Job Card {job_card.job_card_no}.')
        return redirect('printing_production_entry')

    try:
        pass_count = int(raw)
    except (TypeError, ValueError):
        messages.error(request, 'Pass count must be a whole number.')
        return redirect('printing_production_entry')

    max_passes = get_max_print_passes()
    if pass_count < 1 or pass_count > max_passes:
        messages.error(request, f'Pass count must be between 1 and {max_passes}.')
        return redirect('printing_production_entry')
    if not reason:
        messages.error(request, 'A reason is required to override the pass count.')
        return redirect('printing_production_entry')

    job_card.pass_count_override = pass_count
    job_card.pass_count_override_reason = reason
    job_card.pass_count_override_by = request.user
    job_card.pass_count_override_at = timezone.now()
    job_card.save(update_fields=[
        'pass_count_override', 'pass_count_override_reason',
        'pass_count_override_by', 'pass_count_override_at',
    ])
    log_change('job_card', job_card, before_snapshot, request.user, 'update', reason)
    messages.success(
        request,
        f'Pass count for Job Card {job_card.job_card_no} overridden to {pass_count} passes.',
    )
    return redirect('printing_production_entry')


@permission_required('can_edit_production')
def production_data_anomalies(request):
    """Read-only review list of printing entries that look mis-recorded.

    Two categories:
      * Zero-output jobs — a job card with impressions logged across its
        printing entries but a total good-sheet output of 0. This is the
        legacy "operator entered impressions but forgot output" case. It is
        detected at the JOB-CARD level (not per row) so it also catches jobs
        whose rows are all marked as intermediate passes with no final row.
      * Missing impressions — an entry with good sheets recorded but 0
        impressions.

    Each entry links to the standard edit screen. Records older than the edit
    lock are corrected either directly (staff/managers who bypass the lock) or
    via the "Override" request flow on Production Records for other users —
    nothing is changed automatically here.
    """
    printing_qs = Production.objects.filter(is_active=True, entry_type='printing')

    def _row(row, total_passes):
        pass_no = effective_print_pass_number(row, total_passes)
        return {
            'id': row.id,
            'job_card_no': row.job_card.job_card_no if row.job_card else '-',
            'date': row.date,
            'shift': row.shift,
            'machine': row.machine.name if row.machine else '-',
            'pass_label': f'Pass {pass_no} of {total_passes}' + (' (final)' if pass_no >= total_passes else ''),
            'impressions': row.impressions or 0,
            'output_sheets': row.output_sheets or 0,
            'status': row.get_status_display(),
            'edit_url': f"{reverse('printing_production_entry')}?edit={row.id}",
        }

    # --- Zero-output jobs (job-card level) ---
    # Aggregate output + impressions per job card, then keep the job cards that
    # printed something (impressions > 0) but recorded no good sheets at all.
    agg = printing_qs.values('job_card_id').annotate(
        total_output=Sum('output_sheets'),
        total_impressions=Sum('impressions'),
    )
    zero_output_jc_ids = [
        a['job_card_id']
        for a in agg
        if (a['total_impressions'] or 0) > 0 and (a['total_output'] or 0) == 0
    ]

    zero_output_jobs = []
    if zero_output_jc_ids:
        jobs = JobCard.objects.filter(id__in=zero_output_jc_ids).order_by('job_card_no')
        rows_by_jc = {}
        for row in printing_qs.filter(job_card_id__in=zero_output_jc_ids).select_related(
            'job_card', 'machine'
        ).order_by('date', 'created_at'):
            rows_by_jc.setdefault(row.job_card_id, []).append(row)
        for jc in jobs:
            total_passes = get_job_card_pass_count(jc)
            jc_rows = rows_by_jc.get(jc.id, [])
            zero_output_jobs.append({
                'job_card_no': jc.job_card_no,
                'sku': jc.SKU or '-',
                'order_qty': jc.order_qty,
                'total_impressions': sum(r.impressions or 0 for r in jc_rows),
                'pass_count': total_passes,
                'has_completed': any(r.status == 'completed' for r in jc_rows),
                'rows': [_row(r, total_passes) for r in jc_rows],
            })

    # --- Missing impressions (row level) ---
    missing_impression_rows = []
    for row in printing_qs.filter(output_sheets__gt=0, impressions=0).select_related(
        'job_card', 'machine'
    ).order_by('job_card__job_card_no', 'date'):
        if not row.job_card:
            continue
        missing_impression_rows.append(_row(row, get_job_card_pass_count(row.job_card)))

    context = {
        'zero_output_jobs': zero_output_jobs,
        'missing_impression_rows': missing_impression_rows,
        'zero_output_count': len(zero_output_jobs),
        'missing_impression_count': len(missing_impression_rows),
        'can_bypass_edit_lock': user_can_bypass_edit_lock(request.user),
    }
    return render(request, 'production/production_anomalies.html', context)


@login_required
@permission_required('can_view_production_wip')
def production_wip(request):

    default_status_names = ['Printing', 'Printing Completed', 'Sorting / Packing', 'Ready for Dispatch', 'Partial Dispatch', 'Completed']
    for status_name in default_status_names:
        ProductionWipStatus.objects.get_or_create(
            name=status_name,
            defaults={'created_by': request.user},
        )

    if request.method == 'POST':
        action = (request.POST.get('action') or '').strip()
        if action == 'add_status':
            if not request.user.profile.can_manage_production_wip_statuses():
                add_unique_message(request, messages.ERROR, '❌ Only admin can add WIP statuses.')
                return redirect('production_wip')
            new_status_name = (request.POST.get('status_name') or '').strip()
            if not new_status_name:
                add_unique_message(request, messages.ERROR, 'Status name cannot be blank.')
            else:
                ProductionWipStatus.objects.get_or_create(
                    name=new_status_name,
                    defaults={'created_by': request.user},
                )
                add_unique_message(request, messages.SUCCESS, f'Added status: {new_status_name}')
            return redirect('production_wip')

        if action == 'set_job_status':
            job_card_id = request.POST.get('job_card_id')
            status_id = request.POST.get('status_id')
            if not job_card_id or not status_id:
                add_unique_message(request, messages.ERROR, 'Job and status selection are required.')
                return redirect('production_wip')
            job_card = get_object_or_404(
                JobCard,
                id=job_card_id,
                is_active=True,
                status__in=['released', 'in_production'],
            )
            status = get_object_or_404(ProductionWipStatus, id=status_id, is_active=True)
            from production.wip_service import update_wip_status_for_job
            update_wip_status_for_job(job_card, status.name, user=request.user, is_manual=True, force=True)
            add_unique_message(request, messages.SUCCESS, f"{job_card.job_card_no} set to {status.name} (Manual Override).")
            return redirect('production_wip')

    query = (request.GET.get('q') or '').strip()
    status_filter = (request.GET.get('wip_status') or '').strip()
    wip_mode_filter = (request.GET.get('wip_mode') or '').strip()
    calculated_status_filter = (request.GET.get('calculated_status') or '').strip()
    machine_filter = (request.GET.get('machine_id') or '').strip()

    job_cards_qs = JobCard.objects.filter(is_active=True, status__in=['released', 'in_production']).select_related(
        'planning_job', 'machine_name', 'production_wip_status__status'
    )
    from production.wip_service import evaluate_and_update_job_wip_status, get_system_calculated_status_name
    for job_card in job_cards_qs.filter(production_wip_status__isnull=True):
        evaluate_and_update_job_wip_status(job_card, user=request.user)

    if status_filter:
        job_cards_qs = job_cards_qs.filter(production_wip_status__status_id=status_filter)

    if machine_filter:
        job_cards_qs = job_cards_qs.filter(machine_name_id=machine_filter)

    if query:
        job_cards_qs = job_cards_qs.filter(
            Q(job_card_no__icontains=query) |
            Q(PO_No__icontains=query) |
            Q(SKU__icontains=query) |
            Q(planning_job__job_name__icontains=query)
        )

    filtered_job_cards = []
    for job in job_cards_qs:
        calc_status = get_system_calculated_status_name(job)
        job.calculated_status = calc_status

        if calculated_status_filter and calc_status != calculated_status_filter:
            continue

        is_manual = getattr(job.production_wip_status, 'is_manual', False)
        if wip_mode_filter == 'manual' and not is_manual:
            continue
        if wip_mode_filter == 'auto' and is_manual:
            continue

        filtered_job_cards.append(job)

    export_type = (request.GET.get('export') or '').strip().lower()
    if export_type in ('xlsx', 'pdf'):
        from reports.export.services import export_as_pdf, export_as_xlsx
        export_rows = []
        for idx, job in enumerate(filtered_job_cards, start=1):
            export_rows.append({
                'row_number': idx,
                'job_card_no': job.job_card_no,
                'po_number': job.PO_No or '-',
                'sku': job.SKU or '-',
                'job_name': getattr(job.planning_job, 'job_name', '-') or '-',
                'machine_name': job.machine_name_display or '-',
                'order_qty': job.order_qty,
                'production_status': job.workflow_status_label,
                'wip_status': job.wip_status_name,
                'wip_mode': 'Manual' if getattr(job.production_wip_status, 'is_manual', False) else 'Auto',
                'calculated_status': job.calculated_status,
            })
        payload = {
            'report': {'title': 'Production WIP Status Report'},
            'generated_at': timezone.localtime().strftime('%Y-%m-%d %H:%M:%S'),
            'data': export_rows,
            'headers': ['row_number', 'job_card_no', 'po_number', 'sku', 'job_name', 'machine_name', 'order_qty', 'production_status', 'wip_status', 'wip_mode', 'calculated_status'],
            'header_labels': {
                'row_number': '#',
                'job_card_no': 'Job Card No',
                'po_number': 'PO Number',
                'sku': 'SKU',
                'job_name': 'Job Name',
                'machine_name': 'Machine',
                'order_qty': 'Order Qty',
                'production_status': 'Production Status',
                'wip_status': 'Supervisor WIP Status',
                'wip_mode': 'Mode',
                'calculated_status': 'System Calculated Status',
            }
        }
        if export_type == 'xlsx':
            content = export_as_xlsx(payload)
            response = HttpResponse(content, content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            response['Content-Disposition'] = 'attachment; filename="production-wip.xlsx"'
            return response
        else:
            content = export_as_pdf(payload)
            response = HttpResponse(content, content_type='application/pdf')
            response['Content-Disposition'] = 'attachment; filename="production-wip.pdf"'
            return response

    statuses = ProductionWipStatus.objects.filter(is_active=True).order_by('name')
    from core.models import Machine
    machines = Machine.objects.filter(is_active=True).order_by('name')

    params = request.GET.copy()
    params.pop('export', None)
    export_query = params.urlencode()

    context = {
        'statuses': statuses,
        'machines': machines,
        'job_cards': filtered_job_cards,
        'q': query,
        'status_filter': status_filter,
        'wip_mode_filter': wip_mode_filter,
        'calculated_status_filter': calculated_status_filter,
        'machine_filter': machine_filter,
        'export_query': export_query,
    }
    return render(request, 'production/production_wip.html', context)


@login_required
@permission_required('can_view_analytics')
def production_dashboard(request):
    """Real-time production dashboard with OEE metrics"""
    today = timezone.now().date()
    start_date_input = (request.GET.get('start_date') or '').strip()
    end_date_input = (request.GET.get('end_date') or '').strip()

    start_date = None
    end_date = None
    if start_date_input and end_date_input:
        try:
            start_date = datetime.strptime(start_date_input, '%Y-%m-%d').date()
            end_date = datetime.strptime(end_date_input, '%Y-%m-%d').date()
            if start_date > end_date:
                start_date, end_date = end_date, start_date
        except ValueError:
            start_date = None
            end_date = None

    if start_date is None or end_date is None:
        try:
            days = int(request.GET.get('days', 7))
        except (TypeError, ValueError):
            days = 7
        if days < 1:
            days = 1
        end_date = today
        start_date = end_date - timedelta(days=days - 1)
    else:
        days = max((end_date - start_date).days + 1, 1)

    period_productions = Production.objects.filter(
        is_active=True,
        job_card__is_active=True,
        entry_type='printing',
        date__gte=start_date,
        date__lte=end_date,
    )
    period_dispatches = Dispatch.objects.filter(is_active=True, job_card__is_active=True, dispatch_date__gte=start_date, dispatch_date__lte=end_date)

    total_impressions = period_productions.aggregate(total=Sum('impressions'))['total'] or 0
    total_downtime = period_productions.aggregate(total=Sum('downtime_minutes'))['total'] or 0
    total_output = period_productions.aggregate(total=Sum('output_sheets'))['total'] or 0
    total_waste = period_productions.aggregate(total=Sum('waste_sheets'))['total'] or 0
    total_planned_minutes = period_productions.aggregate(total=Sum('planned_time'))['total'] or 0
    total_run_minutes = period_productions.aggregate(total=Sum('run_time'))['total'] or 0
    total_make_ready_minutes = period_productions.aggregate(total=Sum('make_ready_time'))['total'] or 0
    total_actual_minutes = float(total_run_minutes or 0) + float(total_make_ready_minutes or 0) + float(total_downtime or 0)
    planned_variance_minutes = total_actual_minutes - float(total_planned_minutes or 0)
    overrun_setup_minutes = float(total_make_ready_minutes or 0)
    availability_value = round(OEECalculator.availability(total_run_minutes, total_downtime) * 100, 2)
    performance_value = round(OEECalculator.performance(total_impressions, period_productions.aggregate(standard=Sum('machine__standard_impressions_per_hour'))['standard'] or 0, total_run_minutes) * 100, 2) if total_run_minutes else 0
    quality_value = 0
    if total_output + total_waste > 0:
        quality_value = round(OEECalculator.quality(total_output, total_waste, None) * 100, 2)
    oee_value = round(OEECalculator.oee(availability_value / 100, performance_value / 100, quality_value / 100), 2)
    overrun_downtime_minutes = float(sum(
        p.unplanned_downtime_minutes for p in period_productions
    ))
    overrun_run_perf_minutes = max(float(planned_variance_minutes) - overrun_setup_minutes - overrun_downtime_minutes, 0)

    total_dispatch_qty = period_dispatches.aggregate(total=Sum('dispatch_qty'))['total'] or 0
    dispatch_count = period_dispatches.count()
    dispatched_job_cards_count = period_dispatches.values('job_card').distinct().count()
    avg_dispatch_qty = (total_dispatch_qty / dispatch_count) if dispatch_count else 0
    dispatch_fulfillment_pct = (total_dispatch_qty / total_output * 100) if total_output > 0 else 0

    context = {
        'start_date': start_date,
        'end_date': end_date,
        'days': days,
        'period_label': 'Period',
        'oee_value': oee_value,
        'availability_value': availability_value,
        'performance_value': performance_value,
        'quality_value': quality_value,
        'total_impressions': total_impressions,
        'total_output': total_output,
        'total_waste': total_waste,
        'total_downtime': total_downtime,
        'total_planned_minutes': total_planned_minutes,
        'total_run_minutes': total_run_minutes,
        'total_make_ready_minutes': total_make_ready_minutes,
        'total_dispatch_qty': total_dispatch_qty,
        'dispatch_count': dispatch_count,
        'dispatched_job_cards_count': dispatched_job_cards_count,
        'avg_dispatch_qty': avg_dispatch_qty,
        'dispatch_fulfillment_pct': dispatch_fulfillment_pct,
        'shift_comparison_data': [],
        'at_risk_jobs': [],
        'pending_start_count': 0,
        'pending_start_in_period_count': 0,
        'pending_start_jobs': [],
    }

    return render(request, 'production/production_dashboard.html', context)
