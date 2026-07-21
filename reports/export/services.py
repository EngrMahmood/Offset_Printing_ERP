from __future__ import annotations

import csv
import io
from typing import Any


def _json_primitive(value: Any) -> str:
    if value is None:
        return ''
    return str(value)


def _report_totals_line(payload: dict, rows: list[dict[str, Any]]) -> str | None:
    """Total Impressions / Hours summary line, Machine Planning only - other
    reports (e.g. Manual Working) have their own unrelated 'total_impressions'
    field, so this is gated on the report slug, not just key presence."""
    slug = (payload.get('report') or {}).get('slug')
    if slug != 'machine-planning' or not rows:
        return None
    total_impressions = sum(_to_number(row.get('total_impressions')) for row in rows)
    total_hours = sum(_to_number(row.get('estimated_hours')) for row in rows)
    return f'Total Impressions: {total_impressions:,.0f} — Total Hours: {total_hours:,.2f}h'


def _to_number(value: Any) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def _rowset_from_payload(payload: dict) -> list[dict[str, Any]]:
    data = payload.get('data') or {}
    if isinstance(data, list):
        return [item for item in data if isinstance(item, dict)]
    if not isinstance(data, dict):
        return []

    # Prefer common tabular keys from report contexts. 'export_rows' lets a
    # report nominate exactly which of its tables should be exported.
    for key in (
        'export_rows',
        'wastage_rows',
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
    machine_filter = (payload.get('filters') or {}).get('machine')
    if machine_filter:
        report_title = f"{report_title} - {machine_filter}"

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow([report_title])
    writer.writerow(['generated_at', payload.get('generated_at', '')])
    totals_line = _report_totals_line(payload, rows)
    if totals_line:
        writer.writerow([totals_line])
    writer.writerow([])

    if not rows:
        writer.writerow(['message', 'No tabular rows available for export'])
        return output.getvalue().encode('utf-8-sig')

    headers = payload.get('headers') or (payload.get('data') or {}).get('headers')
    if not headers:
        headers = sorted({key for row in rows for key in row.keys()})
    labels = payload.get('header_labels') or (payload.get('data') or {}).get('header_labels') or {}
    writer.writerow([labels.get(header, header) for header in headers])
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
    machine_filter = (payload.get('filters') or {}).get('machine')
    if machine_filter:
        report_title = f"{report_title} - {machine_filter}"

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = 'Report'
    sheet.append([report_title])
    sheet.append(['generated_at', payload.get('generated_at', '')])
    totals_line = _report_totals_line(payload, rows)
    if totals_line:
        sheet.append([totals_line])
    sheet.append([])

    if not rows:
        sheet.append(['message', 'No tabular rows available for export'])
    else:
        headers = payload.get('headers') or (payload.get('data') or {}).get('headers')
        if not headers:
            headers = sorted({key for row in rows for key in row.keys()})
        labels = payload.get('header_labels') or (payload.get('data') or {}).get('header_labels') or {}
        sheet.append([labels.get(header, header) for header in headers])
        for row in rows:
            sheet.append([_json_primitive(row.get(header)) for header in headers])

    stream = io.BytesIO()
    workbook.save(stream)
    return stream.getvalue()


def export_as_pdf(payload: dict) -> bytes:
    try:
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib.units import mm
        from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
        from reportlab.platypus import Paragraph, SimpleDocTemplate, Table, TableStyle, PageBreak
        from reportlab.lib import colors
    except Exception as exc:
        raise RuntimeError('PDF export requires reportlab package.') from exc

    rows = _rowset_from_payload(payload)
    report = payload.get('report') or {}
    report_title = report.get('title') or 'Report'
    machine_filter = (payload.get('filters') or {}).get('machine')
    if machine_filter:
        report_title = f"{report_title} - {machine_filter}"
    generated_at = payload.get('generated_at', '')
    row_count = len(rows)

    buffer = io.BytesIO()
    page_w, page_h = landscape(A4)
    left_margin = right_margin = top_margin = bottom_margin = 12 * mm
    usable_width = page_w - left_margin - right_margin

    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        leftMargin=left_margin,
        rightMargin=right_margin,
        topMargin=top_margin,
        bottomMargin=bottom_margin,
    )

    if not rows:
        stylesheet = getSampleStyleSheet()
        story = [
            Paragraph(f'<b>Offset Printing ERP - {report_title}</b>', stylesheet['Heading3']),
            Paragraph(f'Generated at: {generated_at} — Rows: {row_count}', stylesheet['Normal']),
            Paragraph('No tabular rows available for export', stylesheet['Normal']),
        ]
        doc.build(story)
        return buffer.getvalue()

    headers = payload.get('headers') or (payload.get('data') or {}).get('headers')
    if not headers:
        headers = sorted({key for row in rows for key in row.keys()})
    labels = payload.get('header_labels') or (payload.get('data') or {}).get('header_labels') or {}

    header_style = ParagraphStyle(
        'header',
        fontName='Helvetica-Bold',
        fontSize=7,
        leading=8,
        alignment=0,
    )
    cell_style = ParagraphStyle(
        'cell',
        fontName='Helvetica',
        fontSize=6,
        leading=7,
        alignment=0,
    )

    # Small enough that Machine Planning's 19 columns (with the narrow S#,
    # colour/pass/hour columns) all fit on one landscape A4 page instead of
    # spilling extra columns onto a second page - reduces paper usage.
    min_col_width = 9 * mm
    max_columns = max(1, min(len(headers), int(usable_width // min_col_width)))
    header_chunks = [headers] if len(headers) <= max_columns else [headers[i:i + max_columns] for i in range(0, len(headers), max_columns)]
    stylesheet = getSampleStyleSheet()
    story = [
        Paragraph(f'<b>Offset Printing ERP - {report_title}</b>', stylesheet['Heading3']),
        Paragraph(f'Generated at: {generated_at} — Rows: {row_count}', stylesheet['Normal']),
    ]
    totals_line = _report_totals_line(payload, rows)
    if totals_line:
        story.append(Paragraph(totals_line, stylesheet['Normal']))
    story.append(Paragraph('<br/>', stylesheet['Normal']))

    for chunk_index, chunk_headers in enumerate(header_chunks):
        table_data = [
            [Paragraph(str(labels.get(header, header)), header_style) for header in chunk_headers]
        ]
        for row in rows:
            table_data.append([
                Paragraph(_json_primitive(row.get(header)), cell_style)
                for header in chunk_headers
            ])

        # Dynamically set proportional column widths to prevent excessive
        # wrapping of text fields. Narrow numeric columns get a slim share so
        # all 19 Machine Planning columns fit on one page; kept <=1.0 total
        # (it previously summed to 1.06, silently overflowing the page).
        widths_map = {
            'sequence': 0.025,          # S# - very narrow
            'po_count': 0.03,
            'po_age_days': 0.03,
            'ai_score': 0.035,
            'colors': 0.03,
            'ups': 0.03,
            'passes': 0.03,
            'estimated_hours': 0.035,
            'total_impressions': 0.05,
            'status': 0.05,
            'priority_display': 0.055,
            'machine_name': 0.06,
            'print_sheet_size': 0.05,
            'print_sheet_quantity': 0.05,
            'finish_quantity': 0.05,
            'material': 0.06,
            'sku': 0.09,
            'po_numbers': 0.09,
            'job_card_numbers': 0.09,
        }
        
        col_widths = []
        total_mapped = 0.0
        unmapped_count = 0
        for h in chunk_headers:
            if h in widths_map:
                total_mapped += widths_map[h]
            else:
                unmapped_count += 1
        
        # Distribute remaining width to unmapped columns
        remaining_ratio = max(0.0, 1.0 - total_mapped)
        unmapped_ratio = remaining_ratio / max(unmapped_count, 1)
        
        for h in chunk_headers:
            ratio = widths_map.get(h, unmapped_ratio)
            col_widths.append(ratio * usable_width)

        table = Table(table_data, colWidths=col_widths, repeatRows=1)
        table_style = TableStyle([
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 6),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('GRID', (0, 0), (-1, -1), 0.25, colors.grey),
            ('BACKGROUND', (0, 0), (-1, 0), colors.whitesmoke),
            ('LEFTPADDING', (0, 0), (-1, -1), 3),
            ('RIGHTPADDING', (0, 0), (-1, -1), 3),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ])
        table.setStyle(table_style)
        story.append(table)
        if chunk_index < len(header_chunks) - 1:
            story.append(PageBreak())

    doc.build(story)
    return buffer.getvalue()
