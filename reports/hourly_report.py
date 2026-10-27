"""Hourly Report: a live snapshot of printing, packing, dispatch and planning for the shift supervisor.

The supervisor opens it, taps Share, and WhatsApp receives the message text with an image of the
same figures. Shifts: A is 8 am - 8 pm, B is 8 pm - 8 am. A night shift belongs to the day it
started on, and the business day runs from 8 am.

What each block counts (entries are logged in batches, not hour by hour, so this is a snapshot of
what has been *recorded*, not a per-hour output chart):
  * Printing: impressions of the shift's entries (their own date + shift), machine by machine, and
    the day's total over both shifts. Split entries copied to the members of a combined layout are
    left out, so a combined run's impressions count once.
  * Packing / dispatch, in pieces: what has been logged since 8 am of the business day. Packing is
    often entered the next morning for the previous day, so it is counted by when it was logged.
  * Planning: job cards released since 8 am.
"""
from __future__ import annotations

import datetime
import io
import os

from django.db.models import Count, ExpressionWrapper, F, IntegerField, Max, Sum, Value
from django.db.models.functions import Coalesce
from django.utils import timezone

from core.machine_scope import offset_machines
from core.models import ChangeLog, Dispatch, JobCard, Machine, Production

SHIFT_A_START = datetime.time(8, 0)
SHIFT_B_START = datetime.time(20, 0)
SHIFT_LABELS = {'A': 'Shift A (8 am - 8 pm)', 'B': 'Shift B (8 pm - 8 am)'}
RELEASED_LIST_LIMIT = 10


def shift_for(now: datetime.datetime) -> tuple[str, datetime.date]:
    """(shift, business date) for a local time. After midnight the night shift still belongs to the day before."""
    current = now.time()
    if SHIFT_A_START <= current < SHIFT_B_START:
        return 'A', now.date()
    if current >= SHIFT_B_START:
        return 'B', now.date()
    return 'B', now.date() - datetime.timedelta(days=1)


def _day_start(business_date: datetime.date) -> datetime.datetime:
    return timezone.make_aware(datetime.datetime.combine(business_date, SHIFT_A_START))


def _n(value) -> int:
    return int(value or 0)


def _fmt(value) -> str:
    return f'{int(round(value or 0)):,}'


def _short(text: str, limit: int = 34) -> str:
    text = (text or '').strip()
    return text if len(text) <= limit else text[: limit - 1].rstrip() + '…'


def build_hourly_report(now: datetime.datetime | None = None) -> dict:
    now = timezone.localtime(now or timezone.now())
    shift, business_date = shift_for(now)
    day_start = _day_start(business_date)

    # --- Printing: impressions, machine by machine, for this shift.
    printing = Production.objects.filter(
        is_active=True, entry_type='printing', merge_parent__isnull=True, date=business_date,
    )
    shift_rows = {
        (row['machine__name'] or 'No machine'): row
        for row in printing.filter(shift=shift).values('machine__name').annotate(
            impressions=Coalesce(Sum('impressions'), Value(0)), entries=Count('id'), last=Max('created_at'),
        )
    }
    shift_impressions = sum(_n(r['impressions']) for r in shift_rows.values())
    day_impressions = _n(printing.aggregate(total=Sum('impressions'))['total'])
    machines = [
        {
            'name': name, 'impressions': _n(row['impressions']), 'entries': row['entries'],
            'last_entry': timezone.localtime(row['last']).strftime('%H:%M') if row['last'] else '',
        }
        for name, row in shift_rows.items() if _n(row['impressions']) > 0 or row['entries']
    ]
    machines.sort(key=lambda m: (-m['impressions'], m['name']))
    active_names = set(shift_rows)
    idle_machines = sorted(
        m.name for m in offset_machines(Machine.objects.filter(is_active=True)) if m.name not in active_names
    )

    # --- Packing and dispatch, in pieces, as logged since 8 am.
    packing = Production.objects.filter(is_active=True, entry_type='packing', created_at__gte=day_start).aggregate(
        packed=Coalesce(Sum('packing_qty'), Value(0)), waste=Coalesce(Sum('sorting_waste_qty'), Value(0)),
        entries=Count('id'),
    )
    pcs_expr = ExpressionWrapper(
        F('dispatch_qty') * Coalesce(F('job_card__pcs_per_unit'), Value(1)), output_field=IntegerField(),
    )
    dispatch = Dispatch.objects.filter(is_active=True, created_at__gte=day_start).aggregate(
        pcs=Coalesce(Sum(pcs_expr), Value(0)), dcs=Count('dc_no', distinct=True), entries=Count('id'),
    )

    # --- Planning: job cards released since 8 am.
    logs = list(
        ChangeLog.objects.filter(entity_type='job_card', action='release', created_at__gte=day_start)
        .order_by('-created_at')
    )
    cards = {c.id: c for c in JobCard.objects.filter(id__in=[log.record_id for log in logs])}
    released, seen = [], set()
    for log in logs:
        card = cards.get(log.record_id)
        if card is None or card.id in seen:
            continue
        seen.add(card.id)
        released.append({'jc_no': card.job_card_no, 'sku': card.SKU or '', 'qty_pcs': int(card.order_qty_pcs)})

    totals = {
        'day_impressions': day_impressions,
        'shift_impressions': shift_impressions,
        'packed_pcs': _n(packing['packed']),
        'dispatched_pcs': _n(dispatch['pcs']),
        'released_jobs': len(released),
    }

    from reports.models import HourlyReportSnapshot

    previous = (
        HourlyReportSnapshot.objects.filter(business_date=business_date, taken_at__lt=now).order_by('-taken_at').first()
    )
    delta = None
    if previous is not None:
        before = previous.payload.get('totals', {})
        delta = {
            'since': timezone.localtime(previous.taken_at).strftime('%H:%M'),
            'impressions': totals['day_impressions'] - _n(before.get('day_impressions')),
            'packed_pcs': totals['packed_pcs'] - _n(before.get('packed_pcs')),
            'dispatched_pcs': totals['dispatched_pcs'] - _n(before.get('dispatched_pcs')),
            'released_jobs': totals['released_jobs'] - _n(before.get('released_jobs')),
        }

    report = {
        'generated_at': now.isoformat(),
        'date_label': now.strftime('%a %d %b %Y'),
        'time_label': now.strftime('%H:%M'),
        'shift': shift,
        'shift_label': SHIFT_LABELS[shift],
        'business_date': business_date.isoformat(),
        'printing': {
            'shift_impressions': shift_impressions, 'day_impressions': day_impressions,
            'machines': machines, 'idle_machines': idle_machines,
        },
        'packing': {'packed_pcs': _n(packing['packed']), 'sorting_waste_pcs': _n(packing['waste']), 'entries': packing['entries']},
        'dispatch': {'pcs': _n(dispatch['pcs']), 'dcs': dispatch['dcs'], 'entries': dispatch['entries']},
        'planning': {
            'released_jobs': len(released), 'released_qty_pcs': sum(r['qty_pcs'] for r in released),
            'released': released[:RELEASED_LIST_LIMIT], 'more': max(len(released) - RELEASED_LIST_LIMIT, 0),
        },
        'totals': totals,
        'delta': delta,
    }
    report['text'] = build_text(report)
    return report


def _signed(value: int) -> str:
    return f'+{value:,}' if value >= 0 else f'{value:,}'


def build_text(report: dict, shared_by: str = '') -> str:
    """The WhatsApp message. *bold* is WhatsApp formatting."""
    p, k, d, pl = report['printing'], report['packing'], report['dispatch'], report['planning']
    lines = [
        '*HOURLY PRODUCTION REPORT*',
        f"{report['date_label']} | {report['time_label']} | {report['shift_label']}",
        '',
        '*PRINTING - impressions*',
        f"This shift: {_fmt(p['shift_impressions'])}   |   Day: {_fmt(p['day_impressions'])}",
    ]
    for m in p['machines']:
        lines.append(f"- {m['name']}: {_fmt(m['impressions'])}" + (f" (last entry {m['last_entry']})" if m['last_entry'] else ''))
    if p['idle_machines']:
        lines.append(f"No entry yet: {', '.join(p['idle_machines'])}")
    lines += [
        '',
        '*PACKING - pcs (logged today)*',
        f"Packed: {_fmt(k['packed_pcs'])}   |   Sorting waste: {_fmt(k['sorting_waste_pcs'])}",
        '',
        '*DISPATCH - pcs (logged today)*',
        f"Dispatched: {_fmt(d['pcs'])} pcs in {d['dcs']} DC{'s' if d['dcs'] != 1 else ''}",
        '',
        '*PLANNING - released today*',
    ]
    if pl['released_jobs']:
        lines.append(f"{pl['released_jobs']} job{'s' if pl['released_jobs'] != 1 else ''} released ({_fmt(pl['released_qty_pcs'])} pcs)")
        for r in pl['released']:
            lines.append(f"- {r['jc_no']} {_short(r['sku'])} - {_fmt(r['qty_pcs'])}")
        if pl['more']:
            lines.append(f"...and {pl['more']} more")
    else:
        lines.append('No job released yet')
    delta = report.get('delta')
    if delta:
        lines += [
            '',
            f"*Since last report ({delta['since']})*",
            f"Impressions {_signed(delta['impressions'])} | Packed {_signed(delta['packed_pcs'])} | "
            f"Dispatched {_signed(delta['dispatched_pcs'])} | Released {_signed(delta['released_jobs'])}",
        ]
    if shared_by:
        lines += ['', f'_Shared by {shared_by}_']
    return '\n'.join(lines)


# --------------------------------------------------------------------------- image

def _font(bold: bool, size: int):
    import reportlab
    from PIL import ImageFont

    folder = os.path.join(os.path.dirname(reportlab.__file__), 'fonts')
    return ImageFont.truetype(os.path.join(folder, 'VeraBd.ttf' if bold else 'Vera.ttf'), size)


def render_image(report: dict, shared_by: str = '') -> bytes:
    """PNG of the report: the same figures as the text, laid out as cards, sized for a phone."""
    from PIL import Image, ImageDraw

    width, pad = 1080, 40
    ink, grey, line = (15, 23, 42), (100, 116, 139), (226, 232, 240)
    colours = {'printing': (37, 99, 235), 'packing': (22, 163, 74), 'dispatch': (234, 88, 12), 'planning': (124, 58, 237)}
    h1, h2, body, small = _font(True, 40), _font(True, 30), _font(False, 27), _font(False, 23)
    bold = _font(True, 27)

    image = Image.new('RGB', (width, 4000), 'white')
    draw = ImageDraw.Draw(image)
    y = 0
    draw.rectangle([0, 0, width, 150], fill=(15, 23, 42))
    draw.text((pad, 28), 'HOURLY PRODUCTION REPORT', font=h1, fill='white')
    draw.text((pad, 92), f"{report['date_label']}  |  {report['time_label']}  |  {report['shift_label']}", font=body, fill=(203, 213, 225))
    y = 150 + 26

    def section(title, key):
        nonlocal y
        draw.rectangle([pad, y, width - pad, y + 54], fill=colours[key])
        draw.text((pad + 18, y + 9), title, font=h2, fill='white')
        y += 54 + 14

    def kv(label, value, strong=True):
        nonlocal y
        draw.text((pad + 8, y), label, font=body, fill=grey)
        text = str(value)
        draw.text((width - pad - 8 - draw.textlength(text, font=bold if strong else body), y), text, font=bold if strong else body, fill=ink)
        y += 42

    def rule():
        nonlocal y
        draw.line([pad, y, width - pad, y], fill=line, width=2)
        y += 12

    p, k, d, pl = report['printing'], report['packing'], report['dispatch'], report['planning']
    section('PRINTING - impressions', 'printing')
    kv('This shift', _fmt(p['shift_impressions']))
    kv('Day total', _fmt(p['day_impressions']))
    if p['machines']:
        rule()
        for m in p['machines']:
            draw.text((pad + 8, y), m['name'], font=body, fill=ink)
            value = _fmt(m['impressions'])
            draw.text((width - pad - 8 - draw.textlength(value, font=bold), y), value, font=bold, fill=ink)
            if m['last_entry']:
                draw.text((pad + 8, y + 31), f"last entry {m['last_entry']}", font=small, fill=grey)
                y += 28
            y += 44
    if p['idle_machines']:
        draw.text((pad + 8, y), 'No entry yet: ' + ', '.join(p['idle_machines']), font=small, fill=grey)
        y += 34
    y += 18

    section('PACKING - pcs (logged today)', 'packing')
    kv('Packed', _fmt(k['packed_pcs']))
    kv('Sorting waste', _fmt(k['sorting_waste_pcs']))
    y += 18

    section('DISPATCH - pcs (logged today)', 'dispatch')
    kv('Dispatched', f"{_fmt(d['pcs'])} pcs")
    kv('Delivery challans', d['dcs'])
    y += 18

    section('PLANNING - released today', 'planning')
    if pl['released_jobs']:
        kv('Jobs released', f"{pl['released_jobs']}  ({_fmt(pl['released_qty_pcs'])} pcs)")
        rule()
        for r in pl['released']:
            draw.text((pad + 8, y), f"{r['jc_no']}  {_short(r['sku'], 28)}", font=small, fill=ink)
            value = _fmt(r['qty_pcs'])
            draw.text((width - pad - 8 - draw.textlength(value, font=small), y), value, font=small, fill=ink)
            y += 34
        if pl['more']:
            draw.text((pad + 8, y), f"...and {pl['more']} more", font=small, fill=grey)
            y += 34
    else:
        draw.text((pad + 8, y), 'No job released yet', font=body, fill=grey)
        y += 42
    y += 18

    delta = report.get('delta')
    if delta:
        draw.rectangle([pad, y, width - pad, y + 118], fill=(241, 245, 249))
        draw.text((pad + 18, y + 12), f"Since last report ({delta['since']})", font=bold, fill=ink)
        draw.text((pad + 18, y + 56),
                  f"Impr {_signed(delta['impressions'])}   Packed {_signed(delta['packed_pcs'])}   "
                  f"Disp {_signed(delta['dispatched_pcs'])}   Rel {_signed(delta['released_jobs'])}", font=small, fill=ink)
        y += 118 + 18
    if shared_by:
        draw.text((pad, y), f'Shared by {shared_by}', font=small, fill=grey)
        y += 36
    y += 20

    image = image.crop((0, 0, width, y))
    buffer = io.BytesIO()
    image.save(buffer, format='PNG', optimize=True)
    return buffer.getvalue()
