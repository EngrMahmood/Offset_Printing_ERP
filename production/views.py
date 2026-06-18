from django.shortcuts import render, get_object_or_404, redirect
from django.http import JsonResponse, Http404
from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.utils import timezone
from django.db.models import Q, Sum, Count
from django.db import transaction
from django.core.exceptions import ValidationError
from django.core.paginator import Paginator
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
    user_can_archive_records,
    user_can_bypass_edit_lock,
)
from workflow.services import start_production
from production.services import OEECalculator

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
@permission_required('can_edit_production')
def production_entry(request):
    """Production data entry form for operators"""
    view_id = (request.GET.get('view') or '').strip()
    is_view_mode = bool(view_id)
    edit_id = '' if is_view_mode else (request.POST.get('edit_id') or request.GET.get('edit') or '').strip()
    edit_record = get_active_record_or_404(Production, view_id) if is_view_mode else (get_active_record_or_404(Production, edit_id) if edit_id else None)
    if edit_record and not is_view_mode and not ensure_edit_lock_allowed(request, 'production', edit_record):
        return redirect('production_records')

    def resolve_related_machine(job_card):
        if job_card.machine_name_id:
            return job_card.machine_name

        display_name = (job_card.machine_name_display or '').strip()
        if not display_name:
            return None

        machine = Machine.objects.filter(name__iexact=display_name).first()
        if machine:
            return machine

        machine = Machine.objects.filter(name__istartswith=display_name).first()
        if machine:
            return machine

        machine = Machine.objects.filter(name__icontains=display_name).first()
        if machine:
            return machine

        normalized = re.sub(r'[^A-Za-z0-9 ]+', ' ', display_name).strip()
        if normalized and normalized != display_name:
            machine = Machine.objects.filter(name__icontains=normalized).first()
            if machine:
                return machine

        return None

    def get_effective_job_card_plan(job_card):
        planned_run = float(job_card.estimated_run_time_minutes or 0)
        planned_setup = float(job_card.estimated_setup_time_minutes or 0)
        planned_total = float(job_card.estimated_total_time_minutes or 0)
        machine_obj = resolve_related_machine(job_card)

        if planned_total <= 0 and job_card.total_impressions_required and machine_obj:
            fallback_run, fallback_setup, fallback_total = compute_planned_minutes(
                job_card.total_impressions_required,
                machine_obj,
                job_card.colour,
            )
            planned_run = float(planned_run or fallback_run or 0)
            planned_setup = float(planned_setup or fallback_setup or 0)
            planned_total = float(planned_total or fallback_total or 0)

        return {
            'machine_obj': machine_obj,
            'planned_total': planned_total,
            'planned_run': planned_run,
            'planned_setup': planned_setup,
        }

    def get_remaining_planned_for_job_card(job_card, planned_total, exclude_production_id=None):
        if planned_total <= 0:
            return 0
        allocated_qs = job_card.productions.filter(is_active=True)
        if exclude_production_id:
            allocated_qs = allocated_qs.exclude(pk=exclude_production_id)
        allocated = float(allocated_qs.aggregate(total=Sum('planned_time'))['total'] or 0)
        return max(planned_total - allocated, 0)

    def get_job_card_pass_count(job_card):
        if job_card.planning_job:
            front_pass = int(job_card.planning_job.front_pass or 0)
            back_pass = int(job_card.planning_job.back_pass or 0)
            if front_pass > 0 or back_pass > 0:
                return max(1, front_pass + back_pass)
        colour = (job_card.colour or '').strip()
        match = re.fullmatch(r'(\d+)\s*\+\s*(\d+)', colour)
        if match:
            front = int(match.group(1))
            back = int(match.group(2))
            return max(1, front + back)
        match = re.search(r'(\d+)', colour)
        if match:
            return max(1, int(match.group(1)))
        return 1

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
        intermediate_pass = request.POST.get('intermediate_pass') == 'on'
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
                raise ValueError('Good sheets must be greater than 0 unless this is an intermediate pass')
            if downtime_val > 0 and not primary_downtime_category:
                raise ValueError('Downtime category is required when downtime is greater than 0')
            if waste_sheets_val > 0 and not waste_reason:
                raise ValueError('Waste reason is required when waste sheets are greater than 0')
            if intermediate_pass and impressions_val <= 0:
                raise ValueError('Impressions must be greater than 0 for intermediate pass entry')
            if intermediate_pass and waste_sheets_val < 0:
                raise ValueError('Waste sheets cannot be negative')

            job_card = get_active_record_or_404(JobCard, job_card_id)
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
                'job_card': job_card,
                'machine': machine,
                'operator': operator,
                'shift': shift,
                'date': date,
                'impressions': impressions_val,
                'output_sheets': output_sheets_val,
                'waste_sheets': waste_sheets_val,
                'intermediate_pass': intermediate_pass,
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
            return redirect('production_entry')

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

    job_cards = JobCard.objects.filter(is_active=True, status__in=JOB_CARD_PRODUCTION_START_STATUSES).order_by('-created_at')
    if edit_record:
        job_cards = JobCard.objects.filter(is_active=True).filter(Q(status__in=JOB_CARD_PRODUCTION_START_STATUSES) | Q(pk=edit_record.job_card_id)).distinct().order_by('-created_at')
    machines = Machine.objects.filter(is_active=True)
    operators = Operator.objects.all()
    supervisors = Supervisor.objects.filter(is_active=True).order_by('name')

    job_card_plan_map = {}
    job_card_machine_map = {}
    job_card_info_map = {} # for Section 1 Read-only card
    for j in job_cards:
        plan = get_effective_job_card_plan(j)
        remaining_planned = get_remaining_planned_for_job_card(
            j,
            plan['planned_total'],
            exclude_production_id=edit_record.pk if edit_record else None,
        )

        job_card_plan_map[str(j.id)] = {
            'planned_total': plan['planned_total'],
            'planned_setup': plan['planned_setup'],
            'planned_run': plan['planned_run'],
            'remaining_planned': remaining_planned,
        }
        resolved_machine = plan['machine_obj']
        job_card_machine_map[str(j.id)] = {
            'machine_id': str(resolved_machine.id) if resolved_machine else (j.machine_name_id or ''),
            'mapped_machine_name': resolved_machine.name if resolved_machine else '',
            'job_card_machine_name': j.machine_name_display or '',
        }

        # Dynamic production history for last 5 entries
        history_qs = j.productions.filter(is_active=True).order_by('-date', '-created_at')[:5]
        history_data = []
        for h in history_qs:
            history_data.append({
                'date': h.date.strftime('%d-%b') if h.date else '',
                'shift': h.shift,
                'impressions': f"{h.impressions:,}",
                'output': f"{h.output_sheets:,}",
                'waste': f"{h.waste_sheets:,}",
                'intermediate': 'Yes' if h.intermediate_pass else 'No',
                'runtime': h.run_time,
                'make_ready': f"{float(h.make_ready_time or 0):g}",
                'downtime': f"{float(h.downtime_minutes or 0):g}",
                'status': h.get_status_display(),
            })

        pass_count = get_job_card_pass_count(j)
        total_impressions_used = j.productions.filter(is_active=True).aggregate(total=Sum('impressions'))['total'] or 0
        allowed_impressions = j.total_impressions_allowed_with_tolerance or 0
        remaining_impressions = max(0, allowed_impressions - total_impressions_used)

        job_card_info_map[str(j.id)] = {
            'job_card_no': j.job_card_no,
            'customer': j.destination or '-',
            'product': j.planning_job.job_name if j.planning_job else (j.SKU or '-'),
            'machine': j.machine_name_display or '-',
            'paper': j.material.name if j.material else '-',
            'gsm': '-', # Assuming it's in the name for now per subagent
            'colors': str(j.total_colors) if j.total_colors else '-',
            'order_qty': f"{j.order_qty:,}",
            'required_sheets': f"{int(j.total_sheet_quantity_display or 0):,}",
            'produced_qty': f"{int(j.total_production_pcs or 0):,}",
            'remaining_qty': f"{max(0, j.order_qty - (j.total_production_pcs or 0)):,}",
            'due_date': j.planning_job.delivery_date.strftime('%Y-%m-%d') if j.planning_job and j.planning_job.delivery_date else '-',
            'job_type': (j.planning_job.repeat_flag if j.planning_job and getattr(j.planning_job, 'repeat_flag', None) else "New Job"),
            'pass_count': pass_count,
            'pass_type': f"{pass_count}-pass" if pass_count > 1 else 'Single-pass',
            'allowed_impressions': f"{allowed_impressions:,}",
            'used_impressions': f"{total_impressions_used:,}",
            'remaining_impressions': f"{remaining_impressions:,}",
            'history': history_data
        }

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
    }

    return render(request, 'production/production_entry.html', context)


@login_required
@permission_required('can_edit_production')
def production_records(request):
    """Production records list page"""
    if request.method == 'POST':
        action = (request.POST.get('action') or '').strip()
        if action == 'bulk_delete':
            if request.user.profile.role != 'admin':
                add_unique_message(request, messages.ERROR, '❌ Only admin can run bulk delete.')
                return redirect('production_records')
            if not user_can_archive_records(request.user):
                add_unique_message(request, messages.ERROR, '❌ You do not have permission to delete records.')
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

    records = Production.objects.filter(is_active=True, job_card__is_active=True).select_related('job_card', 'machine', 'operator', 'created_by').order_by('-date', '-id')

    if query:
        records = records.filter(
            Q(job_card__job_card_no__icontains=query) |
            Q(machine__name__icontains=query) |
            Q(operator__name__icontains=query)
        )

    if shift:
        records = records.filter(shift=shift)

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
    paginator = Paginator(records, per_page)
    page_number = request.GET.get('page')
    page_obj = paginator.get_page(page_number)

    records = list(page_obj.object_list)
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

    cutoff = get_record_edit_lock_cutoff()
    pending_ids: set = set()
    approved_ids: set = set()
    if cutoff and not user_can_bypass_edit_lock(request.user):
        user_overrides = EditOverrideRequest.objects.filter(
            entity_type='job_card',
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
        'records': records,
        'page_obj': page_obj,
        'total_count': total_count,
        'q': query,
        'shift': shift,
        'date_from': date_from_raw,
        'date_to': date_to_raw,
        'sort': sort,
        'dir': direction,
        'per_page': per_page,
    }
    return render(request, 'production/production_records.html', context)


@login_required
def production_wip(request):
    profile = getattr(request.user, 'profile', None)
    if not profile or profile.normalized_role not in ('admin', 'manager', 'planner', 'production_manager', 'production', 'operator'):
        messages.error(request, '❌ You do not have permission to access this feature.')
        return redirect('planning:home')

    default_status_names = ['Printing', 'Dispatch']
    for status_name in default_status_names:
        ProductionWipStatus.objects.get_or_create(
            name=status_name,
            defaults={'created_by': request.user},
        )

    if request.method == 'POST':
        action = (request.POST.get('action') or '').strip()
        if action == 'add_status':
            if request.user.profile.role != 'admin':
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
            JobCardWipStatus.objects.update_or_create(
                job_card=job_card,
                defaults={
                    'status': status,
                    'updated_by': request.user,
                },
            )
            add_unique_message(request, messages.SUCCESS, f"{job_card.job_card_no} set to {status.name}.")
            return redirect('production_wip')

    query = (request.GET.get('q') or '').strip()
    status_filter = (request.GET.get('wip_status') or '').strip()

    printing_status = ProductionWipStatus.objects.filter(name='Printing', is_active=True).first()
    job_cards = JobCard.objects.filter(is_active=True, status__in=['released', 'in_production']).select_related('planning_job')
    if printing_status:
        for job_card in job_cards.filter(production_wip_status__isnull=True, status='released'):
            JobCardWipStatus.objects.update_or_create(
                job_card=job_card,
                defaults={
                    'status': printing_status,
                    'updated_by': request.user,
                },
            )

    if status_filter:
        job_cards = job_cards.filter(production_wip_status__status_id=status_filter)

    if query:
        job_cards = job_cards.filter(
            Q(job_card_no__icontains=query) |
            Q(SKU__icontains=query) |
            Q(planning_job__job_name__icontains=query)
        )

    statuses = ProductionWipStatus.objects.filter(is_active=True).order_by('name')

    context = {
        'statuses': statuses,
        'job_cards': job_cards,
        'q': query,
        'status_filter': status_filter,
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

    period_productions = Production.objects.filter(is_active=True, job_card__is_active=True, date__gte=start_date, date__lte=end_date)
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
