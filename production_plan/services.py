"""Plan lifecycle: evening draft, planner approval, morning re-balance, the 8:45 email, plan vs actual.

    18:30  build_draft()     -> next working day's plan (draft); planners get a notification and a red tag
           (planners are blocked from the planning pages until someone approves it)
    08:30  rebalance_today() -> carries the plan forward against what the night actually printed
    08:45  send_plan_email() -> To planners, CC printing; marked "not re-approved" if nobody approved it
"""
from __future__ import annotations

import datetime
import logging
from collections import defaultdict

from django.contrib.auth import get_user_model
from django.core.mail import EmailMultiAlternatives, get_connection
from django.db import transaction
from django.db.models import Sum
from django.urls import reverse
from django.utils import timezone

from core.models import Notification, Production
from production_plan import engine
from production_plan.models import PRIORITY_LABELS, PlanEvent, PlanItem, PlanSettings, ProductionPlan

logger = logging.getLogger(__name__)
User = get_user_model()

ITEM_FIELDS = (
    'lane', 'machine_id', 'plan_date', 'shift', 'sequence', 'planning_job_id', 'jc_number', 'sku', 'order_date', 'qty',
    'remaining_at_plan', 'priority', 'is_new', 'is_chunk', 'pinned', 'manual', 'fifo_reason',
)


# ----------------------------------------------------------------------------- who

def planner_users():
    return User.objects.filter(is_active=True, profile__role='planner')


def can_edit_plan(user):
    if not user.is_authenticated:
        return False
    if user.is_superuser or user.is_staff:
        return True
    profile = getattr(user, 'profile', None)
    return bool(profile and profile.role in ('planner', 'admin'))


def pending_draft():
    """The latest plan if it is a draft the planners have been asked to approve."""
    plan = ProductionPlan.objects.exclude(status='superseded').order_by('-morning_date', '-id').first()
    # A draft only waits for approval (red tag, planner lock) once the planners have been notified; a preview
    # built by an admin earlier in the day stays quiet until the evening run notifies them.
    if plan is not None and plan.status == 'draft' and (plan.summary or {}).get('notified'):
        return plan
    return None


def current_plan():
    return ProductionPlan.objects.exclude(status='superseded').order_by('-morning_date', '-id').first()


def log_event(plan, user, action, detail=''):
    PlanEvent.objects.create(plan=plan, user=user if getattr(user, 'pk', None) else None, action=action, detail=detail[:2000])


# ----------------------------------------------------------------------------- persistence

def _items_as_dicts(plan):
    return [{f: getattr(item, f) for f in ITEM_FIELDS} for item in plan.items.all()]


def _save_build(plan, build, replace=True):
    kept = {key: (plan.summary or {})[key] for key in ('notified',) if key in (plan.summary or {})}
    with transaction.atomic():
        if replace:
            plan.items.all().delete()
        PlanItem.objects.bulk_create([PlanItem(plan=plan, **{k: v for k, v in item.items() if k in ITEM_FIELDS}) for item in build.items])
        plan.horizon_dates = [d.isoformat() for d in build.dates]
        plan.summary = {**build.summary, 'diff': build.diff, **kept}
        plan.save(update_fields=['horizon_dates', 'summary'])


# ----------------------------------------------------------------------------- lifecycle

def build_draft(today=None, settings_obj=None, notify=True):
    """Evening: the plan for the next working day's morning (and the day after), carried forward from the
    plan in force. Returns the new plan, or the existing one if it was already built."""
    settings_obj = settings_obj or PlanSettings.get()
    today = today or timezone.localdate()
    morning = engine.next_working_day(today, settings_obj)
    existing = ProductionPlan.objects.filter(morning_date=morning).exclude(status='superseded').first()
    if existing is not None:
        if notify and existing.status == 'draft' and not (existing.summary or {}).get('notified'):
            rebalance(existing, settings_obj, today)                 # a preview built earlier: refresh it, then ask for approval
            notify_draft(existing)
        return existing
    base = ProductionPlan.objects.exclude(status='superseded').filter(morning_date__lt=morning).order_by('-morning_date', '-id').first()
    build = engine.build_plan(morning, settings_obj, base_items=_items_as_dicts(base) if base else None, today=today)
    with transaction.atomic():
        plan = ProductionPlan.objects.create(morning_date=morning, base_plan=base, status='draft')
        _save_build(plan, build, replace=False)
        ProductionPlan.objects.filter(status='draft').exclude(pk=plan.pk).update(status='superseded')
        log_event(plan, None, 'draft_built', f'{len(build.items)} items for {", ".join(d.isoformat() for d in build.dates)}')
    if notify:
        notify_draft(plan)
    return plan


def notify_draft(plan):
    plan.summary = {**(plan.summary or {}), 'notified': True}
    plan.save(update_fields=['summary'])
    link = reverse('production_plan:board')
    for user in planner_users():
        Notification.objects.create(
            user=user, event_type='production_plan.draft', title='Production plan waiting for your approval',
            message=(f'The plan for {plan.morning_date:%a %d %b} and the next working day is ready. '
                     'Check it and approve it so it can go out in the morning. Planning is locked until you do.'),
            link=link, entity_type='production_plan', entity_id=plan.pk,
        )


def approve(plan, user):
    if plan.status == 'superseded':
        raise ValueError('This plan has been replaced by a newer one.')
    plan.status = 'approved'
    plan.approved_by = user
    plan.approved_at = timezone.now()
    plan.save(update_fields=['status', 'approved_by', 'approved_at'])
    log_event(plan, user, 'approved')
    Notification.objects.filter(entity_type='production_plan', entity_id=plan.pk, is_read=False).update(is_read=True)
    return plan


def rebalance(plan, settings_obj=None, today=None):
    """Morning: carry the plan forward against what has actually been printed. Planner changes stay;
    only printed jobs leave; the freed capacity is refilled by the rules."""
    settings_obj = settings_obj or PlanSettings.get()
    today = today or timezone.localdate()
    build = engine.build_plan(plan.morning_date, settings_obj, base_items=_items_as_dicts(plan), today=today)
    _save_build(plan, build, replace=True)
    plan.rebalanced_at = timezone.now()
    plan.save(update_fields=['rebalanced_at'])
    diff = build.diff
    log_event(plan, None, 'rebalanced',
              f"removed printed: {', '.join(diff['removed_printed'][:30]) or '-'}; "
              f"removed other: {', '.join(diff['removed_other'][:30]) or '-'}; reduced: {len(diff['reduced'])}")
    return plan


def rebalance_today(today=None, settings_obj=None):
    """08:30. Re-balance today's plan (building it first if the evening draft was missed)."""
    settings_obj = settings_obj or PlanSettings.get()
    today = today or timezone.localdate()
    if not engine.is_working_day(today, settings_obj):
        return None
    plan = ProductionPlan.objects.filter(morning_date=today).exclude(status='superseded').first()
    if plan is None:
        previous = ProductionPlan.objects.exclude(status='superseded').order_by('-morning_date', '-id').first()
        build = engine.build_plan(today, settings_obj, base_items=_items_as_dicts(previous) if previous else None, today=today)
        with transaction.atomic():
            plan = ProductionPlan.objects.create(morning_date=today, base_plan=previous, status='draft')
            _save_build(plan, build, replace=False)
            ProductionPlan.objects.filter(status='draft').exclude(pk=plan.pk).update(status='superseded')
            log_event(plan, None, 'draft_built', 'built in the morning: the evening draft was missed')
        notify_draft(plan)
    return rebalance(plan, settings_obj, today)


# ----------------------------------------------------------------------------- reading a plan

def slot_label(day, shift):
    return f"{day:%a %d %b} {engine.SHIFT_LABELS[shift]}"


def plan_grid(plan, settings_obj=None):
    """{lane: {'machines': [...], 'slots': [(date, shift)], 'cells': {(machine_id, date, shift): {...}}, 'items': [...]}}"""
    settings_obj = settings_obj or PlanSettings.get()
    dates = [datetime.date.fromisoformat(d) for d in plan.horizon_dates]
    machines_by_lane = engine.lane_machines(settings_obj)
    all_machines = [m for ms in machines_by_lane.values() for m in ms]
    blocked = engine.blocked_slots(all_machines, dates, settings_obj)
    items = list(plan.items.select_related('machine'))
    by_cell = defaultdict(list)
    for item in items:
        by_cell[(item.machine_id, item.plan_date, item.shift)].append(item)
    grid = {}
    for lane, machines in machines_by_lane.items():
        slots = [(d, sh) for d in dates for sh in engine.SHIFTS]
        cells = {}
        for m in machines:
            for d, sh in slots:
                cell_items = by_cell.get((m.id, d, sh), [])
                cap = 0 if (m.id, d, sh) in blocked else engine.slot_capacity(d, sh, settings_obj)
                total = sum(i.qty for i in cell_items)
                cells[(m.id, d, sh)] = {
                    'items': cell_items, 'total': total, 'cap': cap, 'blocked': (m.id, d, sh) in blocked,
                    'jobs': len(cell_items), 'pct': min(int(total * 100 / cap), 100) if cap else 0,
                    'over': total > cap,
                    'new': sum(1 for i in cell_items if i.is_new),
                }
        grid[lane] = {'machines': machines, 'slots': slots, 'cells': cells,
                      'planned': sum(c['total'] for c in cells.values()), 'capacity': sum(c['cap'] for c in cells.values())}
    return grid


def night_entries_missing(day, settings_obj=None):
    """Machines that had work planned for last night but have no night-shift printing entry logged yet."""
    settings_obj = settings_obj or PlanSettings.get()
    last_night = day - datetime.timedelta(days=1)
    while not engine.is_working_day(last_night, settings_obj) and last_night > day - datetime.timedelta(days=4):
        last_night -= datetime.timedelta(days=1)
    plan = ProductionPlan.objects.exclude(status='superseded').filter(items__plan_date=last_night, items__shift='N').distinct().order_by('-id').first()
    if plan is None:
        return last_night, []
    planned_machines = {i.machine for i in plan.items.filter(plan_date=last_night, shift='N').select_related('machine')}
    logged = set(Production.objects.filter(entry_type='printing', is_active=True, shift='B', date=last_night).values_list('machine_id', flat=True))
    return last_night, sorted(m.name for m in planned_machines if m.id not in logged)


def plan_vs_actual(plan, settings_obj=None, as_of=None):
    """Per machine and shift: impressions planned against impressions actually entered (shift D -> entry shift A, N -> B)."""
    settings_obj = settings_obj or PlanSettings.get()
    as_of = as_of or timezone.now()
    dates = [datetime.date.fromisoformat(d) for d in plan.horizon_dates]
    actual = {}
    for row in (Production.objects.filter(entry_type='printing', is_active=True, merge_parent__isnull=True, date__in=dates)
                .values('machine_id', 'date', 'shift').annotate(i=Sum('impressions'))):
        actual[(row['machine_id'], row['date'], 'D' if row['shift'] == 'A' else 'N')] = row['i'] or 0
    planned = defaultdict(int)
    for item in plan.items.select_related('machine'):
        planned[(item.machine_id, item.plan_date, item.shift, item.lane, item.machine.name)] += item.qty
    rows = []
    local_now = timezone.localtime(as_of)
    for (machine_id, d, sh, lane, name), qty in sorted(planned.items(), key=lambda kv: (kv[0][1], kv[0][2], kv[0][4])):
        shift_start = timezone.make_aware(datetime.datetime.combine(d, datetime.time(8 if sh == 'D' else 20)))
        started = local_now >= shift_start
        got = actual.get((machine_id, d, sh), 0)
        rows.append({
            'lane': lane, 'machine': name, 'date': d, 'shift': sh, 'planned': qty, 'actual': got, 'started': started,
            'pct': int(got * 100 / qty) if qty else 0,
        })
    return rows


# ----------------------------------------------------------------------------- export + email

EXPORT_HEADERS = ['lane', 'date', 'shift', 'machine', 'sequence', 'jc_number', 'sku', 'order_date', 'qty', 'priority', 'new_job', 'pinned', 'note']
EXPORT_LABELS = {
    'lane': 'Lane', 'date': 'Date', 'shift': 'Shift', 'machine': 'Machine', 'sequence': 'Seq', 'jc_number': 'Job Card', 'sku': 'SKU',
    'order_date': 'Order Date (FIFO)', 'qty': 'Impressions Planned', 'priority': 'Priority', 'new_job': 'New Job', 'pinned': 'Pinned', 'note': 'Note',
}


def export_payload(plan, settings_obj=None):
    settings_obj = settings_obj or PlanSettings.get()
    rows = []
    for item in plan.items.select_related('machine'):
        note = []
        if item.is_chunk:
            note.append('part of a job bigger than one shift')
        if item.fifo_reason:
            note.append(f'FIFO jump: {item.fifo_reason}')
        if item.manual:
            note.append('planner change')
        rows.append({
            'lane': engine.LANE_LABELS[item.lane], 'date': f'{item.plan_date:%a %d %b %Y}', 'shift': engine.SHIFT_LABELS[item.shift],
            'machine': item.machine.name, 'sequence': item.sequence, 'jc_number': item.jc_number, 'sku': item.sku,
            'order_date': item.order_date.isoformat() if item.order_date else '', 'qty': item.qty,
            'priority': PRIORITY_LABELS.get(item.priority, item.priority), 'new_job': 'Yes' if item.is_new else '',
            'pinned': 'Yes' if item.pinned else '', 'note': '; '.join(note),
        })
    summary_rows = []
    grid = plan_grid(plan, settings_obj)
    for lane, g in grid.items():
        for m in g['machines']:
            for d, sh in g['slots']:
                cell = g['cells'][(m.id, d, sh)]
                summary_rows.append({
                    'lane': engine.LANE_LABELS[lane], 'machine': m.name, 'slot': slot_label(d, sh), 'planned': cell['total'],
                    'capacity': cell['cap'], 'jobs': cell['jobs'], 'note': 'blocked' if cell['blocked'] else '',
                })
    status = {'approved': 'Approved', 'draft': 'NOT APPROVED by the planner'}.get(plan.status, plan.status)
    return {
        'report': {'title': f'Production Plan - {plan.morning_date:%a %d %b %Y}', 'slug': 'production-plan'},
        'generated_at': timezone.localtime(plan.rebalanced_at or plan.generated_at).strftime('%Y-%m-%d %H:%M'),
        'data': {
            'export_rows': rows, 'headers': EXPORT_HEADERS, 'header_labels': EXPORT_LABELS,
            'export_notes': [f'Status: {status}', 'Horizon: ' + ', '.join(plan.horizon_dates)],
            'extra_sheets': [{
                'title': 'Load by machine and shift',
                'headers': ['lane', 'machine', 'slot', 'planned', 'capacity', 'jobs', 'note'],
                'labels': {'lane': 'Lane', 'machine': 'Machine', 'slot': 'Shift', 'planned': 'Planned (impressions)',
                           'capacity': 'Capacity', 'jobs': 'Jobs', 'note': 'Note'},
                'rows': summary_rows,
            }],
        },
    }


def email_recipients(settings_obj):
    to = [u.email for u in User.objects.filter(is_active=True, profile__role__in=[r.lower() for r in engine._split_names(settings_obj.email_to_roles)]) if u.email]
    to += [a for a in engine._split_names(settings_obj.email_to_extra)]
    cc = engine._split_names(settings_obj.email_cc)
    to = list(dict.fromkeys(a for a in to if a))
    return to, [a for a in dict.fromkeys(cc) if a not in to]


def build_email(plan, settings_obj=None):
    settings_obj = settings_obj or PlanSettings.get()
    grid = plan_grid(plan, settings_obj)
    approved = plan.status == 'approved'
    banner = (
        f'Approved by {plan.approved_by.get_full_name() or plan.approved_by.username} on {timezone.localtime(plan.approved_at):%d %b %H:%M}.'
        if approved and plan.approved_by_id else
        'NOT RE-APPROVED by the planner: this plan was recalculated from the last approved plan. Planner changes are kept; '
        'only printed jobs were removed and the freed capacity was refilled by the rules.'
    )
    diff = (plan.summary or {}).get('diff') or {}
    night_date, missing = night_entries_missing(plan.morning_date, settings_obj)
    yesterday = []
    previous = ProductionPlan.objects.exclude(status='superseded').filter(morning_date__lt=plan.morning_date).order_by('-morning_date').first()
    if previous is not None:
        for row in plan_vs_actual(previous, settings_obj):
            if row['started'] and row['date'] < plan.morning_date:
                yesterday.append(row)
    lane_blocks = []
    for lane, g in grid.items():
        head = ''.join(f'<th style="padding:4px 8px;border:1px solid #cbd5e1;background:#f1f5f9">{slot_label(d, sh)}</th>' for d, sh in g['slots'])
        body = ''
        for m in g['machines']:
            cells = ''
            for d, sh in g['slots']:
                c = g['cells'][(m.id, d, sh)]
                text = 'blocked' if c['blocked'] else f"{c['total']:,} / {c['cap']:,}<br><small>{c['jobs']} job(s){' (' + str(c['new']) + ' new)' if c['new'] else ''}</small>"
                cells += f'<td style="padding:4px 8px;border:1px solid #cbd5e1;text-align:right">{text}</td>'
            body += f'<tr><td style="padding:4px 8px;border:1px solid #cbd5e1"><b>{m.name}</b></td>{cells}</tr>'
        lane_blocks.append(
            f'<h3 style="margin:16px 0 4px">{engine.LANE_LABELS[lane]}</h3>'
            f'<table style="border-collapse:collapse;font-size:13px"><tr><th style="padding:4px 8px;border:1px solid #cbd5e1;background:#f1f5f9">Machine</th>{head}</tr>{body}</table>'
        )
    extra = ''
    if diff.get('removed_printed'):
        extra += f"<p><b>Removed (already printed):</b> {', '.join(diff['removed_printed'][:40])}</p>"
    overflow = (plan.summary or {}).get('overflow_new') or []
    if overflow:
        extra += ('<p style="color:#b91c1c"><b>New jobs that do not fit the day shifts (never moved to night):</b> '
                  + ', '.join(f"{o['jc']} ({o['qty']:,})" for o in overflow[:30]) + '</p>')
    if missing:
        extra += (f'<p style="color:#b45309"><b>Night entries not logged for {night_date:%a %d %b}:</b> {", ".join(missing)}. '
                  'Their planned jobs are unconfirmed until supervisors log them.</p>')
    if yesterday:
        planned = sum(r['planned'] for r in yesterday)
        got = sum(r['actual'] for r in yesterday)
        extra += f'<p><b>Yesterday against plan:</b> {got:,} printed of {planned:,} planned ({int(got * 100 / planned) if planned else 0}%).</p>'
    html = (
        f'<div style="font-family:Arial,sans-serif"><h2 style="margin:0">Production Plan - {plan.morning_date:%A %d %B %Y}</h2>'
        f'<p style="color:{"#166534" if approved else "#b91c1c"}"><b>{banner}</b></p>'
        f'<p>Covers: {", ".join(datetime.date.fromisoformat(d).strftime("%a %d %b") for d in plan.horizon_dates)} - day and night shifts. '
        f'Full job list is in the attached Excel and PDF. Open the board: /production-plan/</p>{extra}{"".join(lane_blocks)}</div>'
    )
    text = f'Production Plan - {plan.morning_date:%d %b %Y}\n{banner}\nFull job list attached (Excel and PDF).'
    return html, text


def send_plan_email(plan, settings_obj=None):
    from reports.export.services import export_as_pdf, export_as_xlsx

    settings_obj = settings_obj or PlanSettings.get()
    to, cc = email_recipients(settings_obj)
    if not (to or cc):
        raise ValueError('No recipients: planners have no email address and no CC is set.')
    payload = export_payload(plan, settings_obj)
    html, text = build_email(plan, settings_obj)
    tag = '' if plan.status == 'approved' else ' (not re-approved)'
    message = EmailMultiAlternatives(
        subject=f'Production Plan - {plan.morning_date:%a %d %b %Y}{tag}', body=text, from_email=None, to=to, cc=cc,
        connection=get_connection(),
    )
    message.attach_alternative(html, 'text/html')
    stem = f'production-plan-{plan.morning_date.isoformat()}'
    message.attach(f'{stem}.xlsx', export_as_xlsx(payload), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    message.attach(f'{stem}.pdf', export_as_pdf(payload), 'application/pdf')
    if not message.send(fail_silently=False):
        raise RuntimeError('The email backend reported that nothing was sent.')
    plan.emailed_at = timezone.now()
    plan.email_status = 'sent' if plan.status == 'approved' else 'sent_unapproved'
    plan.email_error = ''
    plan.save(update_fields=['emailed_at', 'email_status', 'email_error'])
    log_event(plan, None, 'emailed', f"to {', '.join(to)}; cc {', '.join(cc)}")
    return plan


# ----------------------------------------------------------------------------- WhatsApp share (text + picture)

def whatsapp_text(plan, settings_obj=None):
    settings_obj = settings_obj or PlanSettings.get()
    grid = plan_grid(plan, settings_obj)
    approved = plan.status == 'approved'
    lines = [
        '*PRODUCTION PLAN*',
        f"{plan.morning_date:%a %d %b %Y} | " + ', '.join(datetime.date.fromisoformat(d).strftime('%a %d') for d in plan.horizon_dates),
        'Approved by the planner' if approved else '*NOT APPROVED yet*',
    ]
    for lane, g in grid.items():
        lines += ['', f'*{engine.LANE_LABELS[lane].upper()}*']
        for m in g['machines']:
            parts = []
            for d, sh in g['slots']:
                c = g['cells'][(m.id, d, sh)]
                parts.append(f"{d:%a} {engine.SHIFT_LABELS[sh][0]}: " + ('blocked' if c['blocked'] else f"{c['total'] / 1000:.0f}k/{c['jobs']}j"))
            lines.append(f'- {m.name}: ' + ' | '.join(parts))
    overflow = (plan.summary or {}).get('overflow_new') or []
    if overflow:
        lines += ['', f'New jobs not fitting the day shifts: {len(overflow)}']
    return '\n'.join(lines)


def render_plan_image(plan, settings_obj=None):
    """PNG of the plan: per lane, machine by shift, impressions planned and jobs (the job list is in the PDF/Excel)."""
    import io

    from PIL import Image, ImageDraw

    from reports.hourly_report import _font

    settings_obj = settings_obj or PlanSettings.get()
    grid = plan_grid(plan, settings_obj)
    width, pad = 1200, 36
    h1, h2, body, small, bold = _font(True, 36), _font(True, 28), _font(False, 24), _font(False, 20), _font(True, 24)
    ink, grey = (15, 23, 42), (100, 116, 139)
    colours = {'black': (30, 41, 59), 'colour_gto': (37, 99, 235), 'colour_sm': (124, 58, 237)}
    image = Image.new('RGB', (width, 4000), 'white')
    draw = ImageDraw.Draw(image)
    draw.rectangle([0, 0, width, 130], fill=(15, 23, 42))
    draw.text((pad, 22), 'PRODUCTION PLAN', font=h1, fill='white')
    status = 'Approved by the planner' if plan.status == 'approved' else 'NOT APPROVED yet'
    draw.text((pad, 78), f"{plan.morning_date:%a %d %b %Y}   |   {status}", font=body, fill=(203, 213, 225))
    y = 130 + 24
    for lane, g in grid.items():
        draw.rectangle([pad, y, width - pad, y + 50], fill=colours[lane])
        draw.text((pad + 16, y + 8), engine.LANE_LABELS[lane], font=h2, fill='white')
        y += 62
        col_w = (width - 2 * pad - 150) // max(len(g['slots']), 1)
        for idx, (d, sh) in enumerate(g['slots']):
            draw.text((pad + 150 + idx * col_w + 6, y), f"{d:%a %d} {engine.SHIFT_LABELS[sh]}", font=small, fill=grey)
        y += 32
        for m in g['machines']:
            draw.text((pad + 6, y + 6), m.name, font=bold, fill=ink)
            for idx, (d, sh) in enumerate(g['slots']):
                c = g['cells'][(m.id, d, sh)]
                x = pad + 150 + idx * col_w
                if c['blocked']:
                    draw.text((x + 6, y + 6), 'blocked', font=body, fill=(180, 83, 9))
                    continue
                fill = (254, 226, 226) if c['over'] else (220, 252, 231) if c['pct'] >= 90 else (254, 243, 199) if c['pct'] >= 50 else (241, 245, 249)
                draw.rectangle([x, y, x + col_w - 8, y + 54], fill=fill)
                draw.text((x + 8, y + 4), f"{c['total']:,}", font=bold, fill=ink)
                draw.text((x + 8, y + 32), f"{c['jobs']} job(s)" + (f", {c['new']} new" if c['new'] else ''), font=small, fill=grey)
            y += 62
        y += 18
    image = image.crop((0, 0, width, y + 10))
    buffer = io.BytesIO()
    image.save(buffer, format='PNG', optimize=True)
    return buffer.getvalue()
