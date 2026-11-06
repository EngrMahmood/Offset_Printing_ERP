"""Delivery challan (DC): numbering and the printable A4 form (Store copy + Depart copy).

A DC is the set of active Dispatch lines that share one DC number, so nothing about how dispatch is
recorded changes; the challan is a print of those lines."""
from __future__ import annotations

import io
import os
import re
from datetime import date

from django.contrib.staticfiles import finders

DC_PREFIX = 'PP'
DC_FIRST_NUMBER = 1             # the ERP series starts from 0001
DC_PATTERN = re.compile(r'^PP-(\d+)-(\d{2})$', re.IGNORECASE)
COMPANY_NAME = 'UTOPIA PRINTING & PACKAGING'
COMPANY_SUPPLIER_CODE = 'SUP-1471'
COPIES = ('STORE COPY', 'DEPART COPY')
MIN_ROWS = 10


def next_dc_no(today: date | None = None) -> str:
    """Next number in the series, e.g. PP-0001-26. The number keeps counting across years; the suffix is the year."""
    from core.models import Dispatch

    today = today or date.today()
    highest = DC_FIRST_NUMBER - 1
    for number in Dispatch.objects.filter(dc_no__istartswith=f'{DC_PREFIX}-').values_list('dc_no', flat=True):
        match = DC_PATTERN.match((number or '').strip())
        if match:
            highest = max(highest, int(match.group(1)))
    return f'{DC_PREFIX}-{highest + 1:04d}-{today:%y}'


def dc_lines(dc_no: str):
    from core.models import Dispatch

    return list(
        Dispatch.objects.filter(is_active=True, dc_no__iexact=(dc_no or '').strip())
        .select_related('job_card', 'job_card__planning_job', 'created_by')
        .order_by('id')
    )


def _clean(value) -> str:
    return re.sub(r'\s+', ' ', str(value or '')).strip()


def dc_data(dc_no: str) -> dict | None:
    lines = dc_lines(dc_no)
    if not lines:
        return None
    po_numbers, departments, destinations = [], [], []
    for line in lines:
        card = line.job_card
        department = card.department.name if card.department_id else (getattr(card.planning_job, 'department', '') if card.planning_job else '')
        for bucket, value in ((po_numbers, card.PO_No), (departments, department), (destinations, card.destination)):
            value = _clean(value)
            if value and value not in bucket:
                bucket.append(value)
    preparer = next((l.created_by for l in lines if l.created_by), None)
    return {
        'dc_no': lines[0].dc_no,
        'date': min(l.dispatch_date for l in lines),
        'po_no': ', '.join(po_numbers),
        'department': ', '.join(departments),
        'destination': ', '.join(destinations),
        'prepared_by': (preparer.get_full_name() or preparer.username) if preparer else '',
        'rows': [
            {
                'jc': line.job_card.job_card_no,
                'item': _clean(line.job_card.SKU),
                'qty': int(line.dispatch_qty or 0),
            }
            for line in lines
        ],
    }


def _fit(canvas, text, font, size, width):
    """Shrink the font until the text fits the width (long SKU codes)."""
    while size > 5 and canvas.stringWidth(text, font, size) > width:
        size -= 0.5
    return size


def _draw_copy(c, data, label, left, top, width, height):
    """One challan copy inside the box whose top-left is (left, top)."""
    from reportlab.lib import colors

    bottom = top - height
    c.setStrokeColor(colors.black)
    c.setLineWidth(0.8)
    c.rect(left, bottom, width, height)

    # --- header
    header_h = 54
    logo = finders.find('core/dc_logo.png')
    if logo and os.path.exists(logo):
        c.drawImage(logo, left + 8, top - header_h + 6, width=48, height=32, mask='auto', preserveAspectRatio=True)
    c.setFont('Helvetica-Bold', 14)
    c.drawCentredString(left + width / 2, top - 18, COMPANY_NAME)
    c.setFont('Helvetica', 9)
    c.drawCentredString(left + width / 2, top - 30, COMPANY_SUPPLIER_CODE)
    c.setFont('Helvetica-Bold', 12)
    c.drawCentredString(left + width / 2, top - 46, 'DELIVERY CHALLAN')
    c.line(left, top - header_h, left + width, top - header_h)

    # --- reference row: Date | DC NO. | PO NO. | customer | copy label
    row_h = 22
    y = top - header_h
    cells = [
        ('Date', f'{data["date"]:%d-%m-%Y}', 0.14),
        ('DC NO.', data['dc_no'], 0.17),
        ('PO NO.', data['po_no'], 0.21),
        ('Department', data['department'], 0.24),
        ('Destination', data['destination'], 0.24),
    ]
    x = left
    for index, (caption, value, share) in enumerate(cells):
        w = width * share
        c.setFont('Helvetica', 6.5)
        c.drawString(x + 3, y - 8, caption)
        font = 'Helvetica-Bold'
        size = _fit(c, value, font, 9, w - 6)
        c.setFont(font, size)
        c.drawString(x + 3, y - 18, value)
        x += w
        if index < len(cells) - 1:
            c.line(x, y, x, y - row_h)
    c.line(left, y - row_h, left + width, y - row_h)
    c.setFont('Helvetica-Bold', 8)
    c.drawRightString(left + width - 6, top - 12, label)

    # --- item table
    cols = [(28, 'S.#', 'c'), (110, 'JC #', 'l'), (width - 28 - 110 - 90, 'ITEM NAME', 'l'), (90, 'Quantity', 'r')]
    y = y - row_h
    head_h = 16
    c.setFillColor(colors.HexColor('#e8edf3'))
    c.rect(left, y - head_h, width, head_h, fill=1, stroke=0)
    c.setFillColor(colors.black)
    x = left
    c.setFont('Helvetica-Bold', 8)
    for w, title, align in cols:
        if align == 'c':
            c.drawCentredString(x + w / 2, y - 11, title)
        elif align == 'r':
            c.drawRightString(x + w - 4, y - 11, title)
        else:
            c.drawString(x + 4, y - 11, title)
        x += w
    c.line(left, y - head_h, left + width, y - head_h)
    y -= head_h

    line_h = 17
    rows = data['rows']
    shown = max(MIN_ROWS, len(rows))
    for index in range(shown):
        row = rows[index] if index < len(rows) else None
        x = left
        c.setFont('Helvetica', 8)
        values = [str(index + 1), row['jc'] if row else '', row['item'] if row else '',
                  (f'{row["qty"]:,}' if row and row['qty'] else '')]
        for (w, _title, align), value in zip(cols, values):
            if value:
                size = _fit(c, value, 'Helvetica', 8, w - 8)
                c.setFont('Helvetica', size)
                if align == 'c':
                    c.drawCentredString(x + w / 2, y - 12, value)
                elif align == 'r':
                    c.drawRightString(x + w - 4, y - 12, value)
                else:
                    c.drawString(x + 4, y - 12, value)
            x += w
        c.setStrokeColor(colors.HexColor('#9aa7b5'))
        c.setLineWidth(0.4)
        c.line(left, y - line_h, left + width, y - line_h)
        y -= line_h
    c.setStrokeColor(colors.black)
    c.setLineWidth(0.8)
    table_top = top - header_h - row_h
    x = left
    for w, _t, _a in cols[:-1]:
        x += w
        c.line(x, table_top, x, y)

    # --- totals
    c.setFont('Helvetica-Bold', 8.5)
    c.line(left, y, left + width, y)
    c.drawString(left + 32, y - 12, 'TOTAL')
    total_qty = sum(r['qty'] for r in rows)
    c.drawRightString(left + width - 4, y - 12, f'{total_qty:,}')
    y -= 17
    c.line(left, y, left + width, y)

    # --- signatures
    sig_y = bottom + 22
    c.setFont('Helvetica', 8)
    c.line(left + 12, sig_y, left + 150, sig_y)
    c.drawString(left + 12, sig_y - 10, 'Prepared By' + (f': {data["prepared_by"]}' if data['prepared_by'] else ''))
    c.line(left + width - 150, sig_y, left + width - 12, sig_y)
    c.drawRightString(left + width - 12, sig_y - 10, 'Approved By')


def build_dc_pdf(dc_no: str) -> bytes | None:
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas

    data = dc_data(dc_no)
    if not data:
        return None
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4, pageCompression=1)
    c.setTitle(f'Delivery Challan {data["dc_no"]}')
    page_w, page_h = A4
    margin = 28
    width = page_w - 2 * margin
    if len(data['rows']) <= MIN_ROWS:
        half = (page_h - 3 * margin) / 2                 # both copies on one sheet, as the paper form
        _draw_copy(c, data, COPIES[0], margin, page_h - margin, width, half)
        _draw_copy(c, data, COPIES[1], margin, page_h - margin - half - margin, width, half)
        c.showPage()
    else:
        for label in COPIES:                              # a long challan: one full page per copy
            _draw_copy(c, data, label, margin, page_h - margin, width, page_h - 2 * margin)
            c.showPage()
    c.save()
    return buffer.getvalue()
