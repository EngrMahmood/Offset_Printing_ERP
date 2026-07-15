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

    min_col_width = 16 * mm
    max_columns = max(1, min(len(headers), int(usable_width // min_col_width)))
    header_chunks = [headers] if len(headers) <= max_columns else [headers[i:i + max_columns] for i in range(0, len(headers), max_columns)]
    stylesheet = getSampleStyleSheet()
    story = [
        Paragraph(f'<b>Offset Printing ERP - {report_title}</b>', stylesheet['Heading3']),
        Paragraph(f'Generated at: {generated_at} — Rows: {row_count}', stylesheet['Normal']),
        Paragraph('<br/>', stylesheet['Normal']),
    ]

    for chunk_index, chunk_headers in enumerate(header_chunks):
        table_data = [
            [Paragraph(str(labels.get(header, header)), header_style) for header in chunk_headers]
        ]
        for row in rows:
            table_data.append([
                Paragraph(_json_primitive(row.get(header)), cell_style)
                for header in chunk_headers
            ])

        # Dynamically set proportional column widths to prevent excessive wrapping of text fields
        widths_map = {
            'sequence': 0.05,        # 5% of usable width
            'po_count': 0.04,        # 4%
            'po_age_days': 0.04,     # 4%
            'ai_score': 0.04,        # 4%
            'colors': 0.04,          # 4%
            'ups': 0.04,             # 4%
            'status': 0.06,          # 6%
            'priority_display': 0.07, # 7%
            'machine_name': 0.06,    # 6%
            'material': 0.08,        # 8%
            'print_sheet_size': 0.08, # 8%
            'print_sheet_quantity': 0.06, # 6%
            'finish_quantity': 0.07,  # 7%
            'sku': 0.11,             # 11%
            'po_numbers': 0.11,      # 11%
            'job_card_numbers': 0.11, # 11%
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
