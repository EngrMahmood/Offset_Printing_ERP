from __future__ import annotations

import csv
import io
from typing import Any


def _json_primitive(value: Any) -> str:
    if value is None:
        return ''
    return str(value)


def _rowset_from_payload(payload: dict) -> list[dict[str, Any]]:
    data = payload.get('data') or {}
    if isinstance(data, list):
        return [item for item in data if isinstance(item, dict)]
    if not isinstance(data, dict):
        return []

    # Prefer common tabular keys from report contexts.
    for key in (
        'machine_rows',
        'actual_rows',
        'status_rows',
        'by_machine',
        'by_shift',
        'by_date',
        'by_dc',
        'cutting_request_rows',
        'recent_dispatches',
        'pending_qc_jobs',
        'approved_recent',
        'rejected_jobs',
        'overdue_jobs',
        'due_soon_jobs',
        'missing_readiness',
        'missing_plate_jobs',
        'ready_jobs',
    ):
        value = data.get(key)
        if isinstance(value, list) and value and isinstance(value[0], dict):
            return value

    # Fallback to summary if no rowset is available.
    summary = data.get('summary')
    if isinstance(summary, dict):
        return [summary]
    return []


def export_as_csv(payload: dict) -> bytes:
    rows = _rowset_from_payload(payload)
    report_title = (payload.get('report') or {}).get('title') or 'Report'

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow([report_title])
    writer.writerow(['generated_at', payload.get('generated_at', '')])
    writer.writerow([])

    if not rows:
        writer.writerow(['message', 'No tabular rows available for export'])
        return output.getvalue().encode('utf-8-sig')

    headers = sorted({key for row in rows for key in row.keys()})
    writer.writerow(headers)
    for row in rows:
        writer.writerow([_json_primitive(row.get(header)) for header in headers])
    return output.getvalue().encode('utf-8-sig')


def export_as_xlsx(payload: dict) -> bytes:
    try:
        from openpyxl import Workbook  # type: ignore
    except Exception as exc:
        raise RuntimeError('XLSX export requires openpyxl package.') from exc

    rows = _rowset_from_payload(payload)
    report_title = (payload.get('report') or {}).get('title') or 'Report'

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = 'Report'
    sheet.append([report_title])
    sheet.append(['generated_at', payload.get('generated_at', '')])
    sheet.append([])

    if not rows:
        sheet.append(['message', 'No tabular rows available for export'])
    else:
        headers = sorted({key for row in rows for key in row.keys()})
        sheet.append(headers)
        for row in rows:
            sheet.append([_json_primitive(row.get(header)) for header in headers])

    stream = io.BytesIO()
    workbook.save(stream)
    return stream.getvalue()


def export_as_pdf(payload: dict) -> bytes:
    try:
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib.units import mm
        from reportlab.pdfgen import canvas
    except Exception as exc:
        raise RuntimeError('PDF export requires reportlab package.') from exc

    rows = _rowset_from_payload(payload)
    report = payload.get('report') or {}
    report_title = report.get('title') or 'Report'
    generated_at = payload.get('generated_at', '')

    stream = io.BytesIO()
    c = canvas.Canvas(stream, pagesize=landscape(A4))
    page_w, page_h = landscape(A4)

    y = page_h - 12 * mm
    c.setFont('Helvetica-Bold', 14)
    c.drawString(12 * mm, y, f'Offset Printing ERP - {report_title}')
    y -= 6 * mm
    c.setFont('Helvetica', 9)
    c.drawString(12 * mm, y, f'Generated at: {generated_at}')
    y -= 8 * mm

    if not rows:
        c.drawString(12 * mm, y, 'No tabular rows available for export')
        c.showPage()
        c.save()
        return stream.getvalue()

    headers = sorted({key for row in rows for key in row.keys()})
    headers = headers[:10]
    col_width = (page_w - 24 * mm) / max(len(headers), 1)

    def _draw_header(current_y):
        c.setFont('Helvetica-Bold', 8)
        x = 12 * mm
        for header in headers:
            c.drawString(x, current_y, str(header)[:24])
            x += col_width

    _draw_header(y)
    y -= 5 * mm
    c.setFont('Helvetica', 8)

    for row in rows:
        if y < 12 * mm:
            c.showPage()
            y = page_h - 12 * mm
            _draw_header(y)
            y -= 5 * mm
            c.setFont('Helvetica', 8)

        x = 12 * mm
        for header in headers:
            text = _json_primitive(row.get(header))
            c.drawString(x, y, text[:28])
            x += col_width
        y -= 4.5 * mm

    c.showPage()
    c.save()
    return stream.getvalue()
