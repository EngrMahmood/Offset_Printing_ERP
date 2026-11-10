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


HEADER_H = 54            # company name block
HEAD_ROW_H = 16          # column titles
TOTAL_H = 17
FOOT_H = 38              # space for the signature lines
MIN_ROW_H = 17
LINE_GAP = 9.6
BODY_FONT = 8
MAX_HEADER_LINES = 8


def _wrap(canvas, text, font, size, width):
    """Wrap text to the width, breaking on spaces and, for a long unbroken code, on characters."""
    text = re.sub(r'\s+', ' ', str(text or '')).strip()
    if not text:
        return ['']
    lines, current = [], ''
    for word in text.split(' '):
        candidate = f'{current} {word}'.strip()
        if canvas.stringWidth(candidate, font, size) <= width:
            current = candidate
            continue
        if current:
            lines.append(current)
            current = ''
        while canvas.stringWidth(word, font, size) > width and len(word) > 1:
            cut = len(word) - 1
            while cut > 1 and canvas.stringWidth(word[:cut], font, size) > width:
                cut -= 1
            lines.append(word[:cut])
            word = word[cut:]
        current = word
    if current:
        lines.append(current)
    return lines


def _measure(c, data, width, body_size):
    """Work out the wrapped text and heights for one copy; returns a layout dict including the height needed."""
    cells = [
        ('Date', f'{data["date"]:%d-%m-%Y}', 0.14),
        ('DC NO.', data['dc_no'], 0.17),
        ('PO NO.', data['po_no'], 0.21),
        ('Department', data['department'], 0.24),
        ('Destination', data['destination'], 0.24),
    ]
    header_cells, most = [], 1
    for caption, value, share in cells:
        lines = _wrap(c, value, 'Helvetica-Bold', 8.5, width * share - 8)
        if len(lines) > MAX_HEADER_LINES:
            lines = lines[:MAX_HEADER_LINES]
            lines[-1] = lines[-1][:-1] + '…'
        header_cells.append((caption, lines, width * share))
        most = max(most, len(lines))
    ref_h = max(22, 10 + most * LINE_GAP + 3)

    cols = [(28, 'S.#', 'c'), (110, 'JC #', 'l'), (width - 28 - 110 - 90, 'ITEM NAME', 'l'), (90, 'Quantity', 'r')]
    line_gap = body_size * 1.2
    rows = []
    for index, row in enumerate(data['rows']):
        jc_lines = _wrap(c, row['jc'], 'Helvetica', body_size, cols[1][0] - 8)
        item_lines = _wrap(c, row['item'], 'Helvetica', body_size, cols[2][0] - 8)
        n = max(len(jc_lines), len(item_lines))
        rows.append({'no': str(index + 1), 'jc': jc_lines, 'item': item_lines, 'qty': f'{row["qty"]:,}' if row['qty'] else '',
                     'h': max(MIN_ROW_H, n * line_gap + 7)})
    body_h = sum(r['h'] for r in rows)
    blanks = max(0, -(-int(MIN_ROWS * MIN_ROW_H - body_h) // MIN_ROW_H))          # blank rows only fill up to the paper form's height
    body_h += blanks * MIN_ROW_H
    needed = HEADER_H + ref_h + HEAD_ROW_H + body_h + TOTAL_H + FOOT_H
    return {'cells': header_cells, 'ref_h': ref_h, 'cols': cols, 'rows': rows, 'blanks': blanks, 'needed': needed,
            'line_gap': line_gap, 'size': body_size}


def _draw_copy(c, data, layout, label, left, top, width, height):
    """One challan copy inside the box whose top-left is (left, top)."""
    from reportlab.lib import colors

    bottom = top - height
    c.setStrokeColor(colors.black)
    c.setLineWidth(0.8)
    c.rect(left, bottom, width, height)

    # --- header
    logo = finders.find('core/dc_logo.png')
    if logo and os.path.exists(logo):
        c.drawImage(logo, left + 8, top - HEADER_H + 6, width=48, height=32, mask='auto', preserveAspectRatio=True)
    c.setFont('Helvetica-Bold', 14)
    c.drawCentredString(left + width / 2, top - 18, COMPANY_NAME)
    c.setFont('Helvetica', 9)
    c.drawCentredString(left + width / 2, top - 30, COMPANY_SUPPLIER_CODE)
    c.setFont('Helvetica-Bold', 12)
    c.drawCentredString(left + width / 2, top - 46, 'DELIVERY CHALLAN')
    c.line(left, top - HEADER_H, left + width, top - HEADER_H)
    c.setFont('Helvetica-Bold', 8)
    c.drawRightString(left + width - 6, top - 12, label)

    # --- reference row (wraps for long department / destination / PO text)
    y = top - HEADER_H
    ref_h = layout['ref_h']
    x = left
    for index, (caption, lines, w) in enumerate(layout['cells']):
        c.setFont('Helvetica', 6.5)
        c.drawString(x + 3, y - 8, caption)
        c.setFont('Helvetica-Bold', 8.5)
        for n, text in enumerate(lines):
            c.drawString(x + 3, y - 17 - n * LINE_GAP, text)
        x += w
        if index < len(layout['cells']) - 1:
            c.line(x, y, x, y - ref_h)
    c.line(left, y - ref_h, left + width, y - ref_h)

    # --- item table
    cols = layout['cols']
    y -= ref_h
    c.setFillColor(colors.HexColor('#e8edf3'))
    c.rect(left, y - HEAD_ROW_H, width, HEAD_ROW_H, fill=1, stroke=0)
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
    c.line(left, y - HEAD_ROW_H, left + width, y - HEAD_ROW_H)
    y -= HEAD_ROW_H
    table_top = y

    size, gap = layout['size'], layout['line_gap']
    body_rows = list(layout['rows']) + [
        {'no': str(len(layout['rows']) + k + 1), 'jc': [''], 'item': [''], 'qty': '', 'h': MIN_ROW_H} for k in range(layout['blanks'])
    ]
    for row in body_rows:
        c.setFont('Helvetica', size)
        c.drawCentredString(left + cols[0][0] / 2, y - 3 - size, row['no'])
        x = left + cols[0][0]
        for n, text in enumerate(row['jc']):
            c.drawString(x + 4, y - 3 - size - n * gap, text)
        x += cols[1][0]
        for n, text in enumerate(row['item']):
            c.drawString(x + 4, y - 3 - size - n * gap, text)
        if row['qty']:
            c.drawRightString(left + width - 4, y - 3 - size, row['qty'])
        c.setStrokeColor(colors.HexColor('#9aa7b5'))
        c.setLineWidth(0.4)
        c.line(left, y - row['h'], left + width, y - row['h'])
        y -= row['h']
    c.setStrokeColor(colors.black)
    c.setLineWidth(0.8)
    x = left
    for w, _t, _a in cols[:-1]:
        x += w
        c.line(x, table_top, x, y)

    # --- totals
    c.setFont('Helvetica-Bold', 8.5)
    c.line(left, y, left + width, y)
    c.drawString(left + 32, y - 12, 'TOTAL')
    total_qty = sum(r['qty'] for r in data['rows'])
    c.drawRightString(left + width - 4, y - 12, f'{total_qty:,}')
    y -= TOTAL_H
    c.line(left, y, left + width, y)

    # --- signatures
    sig_y = bottom + 22
    c.setFont('Helvetica', 8)
    c.line(left + 12, sig_y, left + 150, sig_y)
    c.drawString(left + 12, sig_y - 10, 'Prepared By' + (f': {data["prepared_by"]}' if data['prepared_by'] else ''))
    c.line(left + width - 150, sig_y, left + width - 12, sig_y)
    c.drawRightString(left + width - 12, sig_y - 10, 'Approved By')


def build_dc_pdf(dc_no: str) -> bytes | None:
    """A4 portrait. Long item names wrap inside their row (the row grows, like a wide Excel row) so the page size
    never changes: Store copy and Depart copy share one sheet when they fit, otherwise each copy gets its own page."""
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
    half = (page_h - 3 * margin) / 2
    full = page_h - 2 * margin

    layout = _measure(c, data, width, BODY_FONT)
    if layout['needed'] <= half:
        _draw_copy(c, data, layout, COPIES[0], margin, page_h - margin, width, half)
        _draw_copy(c, data, layout, COPIES[1], margin, page_h - margin - half - margin, width, half)
        c.showPage()
    else:
        size = BODY_FONT
        while layout['needed'] > full and size > 5.5:        # a very long challan: shrink the text a little, last resort
            size -= 0.5
            layout = _measure(c, data, width, size)
        for label in COPIES:
            _draw_copy(c, data, layout, label, margin, page_h - margin, width, full)
            c.showPage()
    c.save()
    return buffer.getvalue()
