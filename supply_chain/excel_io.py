import csv
import io
from datetime import datetime

from django.http import HttpResponse
from django.utils import timezone

from .models import StockDemand, StockTransaction, SupplyChainItem

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False


TRANSACTION_HEADERS = [
    'Month',
    'Date',
    'GIN / JC',
    'Item ID',
    'Item Type',
    'UOM',
    'Sheet Packing/Pcs',
    'Sheet Qty/Pcs',
    'Pkt/Rim Qty',
]

DEMAND_HEADERS = [
    'Item ID',
    'Item Type',
    'UOM',
    'Sheet Packing/Pcs',
    'Month',
    'Sheet Qty/Pcs',
    'Pkt/Rim Qty',
]

ITEM_HEADERS = [
    'Item ID',
    'Item Type',
    'UOM',
    'Sheet Packing/Pcs',
    'Unit Cost',
    'Safety Stock',
    'Max Stock Level',
    'Lead Time (Days)',
]

ITEM_WISE_CONSUMPTION_HEADERS = [
    'Item ID',
    'Item Type',
    'UOM',
    'Sheet Packing/Pcs',
    'Month',
    'Sheet Qty/Pcs',
    'Pkt/Rim Qty',
    'Consumption Value',
]

KPI_HEADERS = [
    'Item ID',
    'Item Type',
    'ABC',
    'FSN',
    'Closing Stock',
    'Monthly Demand',
    'Daily Demand',
    'Reorder Point',
    'Safety Stock',
    'Inventory Value',
    'Inventory Turnover',
    'Annual Consumption Value',
    'Stockout',
    'Reorder Alert',
    'Safety Stock Alert',
    'Overstock',
    'Dead Stock',
    'Slow Moving',
    'Inventory Accuracy %',
]

PHYSICAL_COUNT_HEADERS = [
    'Count Date',
    'Item ID',
    'Item Type',
    'Physical Sheet Qty/Pcs',
    'System Sheet Qty/Pcs',
    'Variance',
    'Inventory Accuracy %',
    'Notes',
]

MONTH_WISE_CONSUMPTION_HEADERS = [
    'Month',
    'Item ID',
    'Item Type',
    'UOM',
    'Sheet Packing/Pcs',
    'Sheet Qty/Pcs',
    'Pkt/Rim Qty',
    'Consumption Value',
]


def _parse_date(value):
    if value is None or value == '':
        return timezone.now().date()
    if hasattr(value, 'date'):
        return value.date()
    text = str(value).strip()
    for fmt in ('%Y-%m-%d', '%d-%m-%Y', '%d/%m/%Y', '%m/%d/%Y'):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return timezone.now().date()


def _get_item_by_id(item_id):
    item_id = (item_id or '').strip()
    if not item_id:
        return None
    return SupplyChainItem.objects.filter(item_id__iexact=item_id).select_related('material').first()


def _read_rows(upload_file):
    name = (upload_file.name or '').lower()
    rows = []

    if name.endswith('.csv'):
        decoded = upload_file.read().decode('utf-8-sig')
        reader = csv.DictReader(io.StringIO(decoded))
        for row in reader:
            rows.append({(k or '').strip(): v for k, v in row.items()})
        return rows

    if name.endswith('.xlsx'):
        if not EXCEL_AVAILABLE:
            raise ImportError('openpyxl is required for XLSX upload.')
        wb = openpyxl.load_workbook(upload_file, data_only=True)
        ws = wb.active
        header = None
        for row in ws.iter_rows(min_row=1, max_row=1, values_only=True):
            header = [str(c).strip() if c is not None else '' for c in row]
        if not header:
            return []
        for values in ws.iter_rows(min_row=2, values_only=True):
            if not any(values):
                continue
            row = {}
            for idx, key in enumerate(header):
                if key:
                    row[key] = values[idx] if idx < len(values) else None
            rows.append(row)
        return rows

    raise ValueError('Unsupported file type. Please upload CSV or XLSX.')


def _write_workbook(filename, headers, data_rows):
    if not EXCEL_AVAILABLE:
        raise ImportError('openpyxl is required for Excel export.')

    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    header_font = Font(bold=True, color='FFFFFF')

    for col_num, header in enumerate(headers, 1):
        cell = worksheet.cell(row=1, column=col_num, value=header)
        cell.fill = header_fill
        cell.font = header_font

    for row_num, row in enumerate(data_rows, 2):
        for col_num, value in enumerate(row, 1):
            worksheet.cell(row=row_num, column=col_num, value=value)

    response = HttpResponse(
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
    response['Content-Disposition'] = f'attachment; filename="{filename}"'
    workbook.save(response)
    return response


def export_transactions(transaction_type, queryset, filename):
    rows = []
    for txn in queryset:
        rows.append([
            txn.month_str or '',
            txn.date.isoformat() if txn.date else '',
            txn.gin_jc or '',
            txn.item.item_id or '',
            txn.item.material.name,
            txn.item.uom,
            txn.item.sheet_packing_pcs,
            txn.sheet_qty_pcs,
            txn.pkt_rim_qty,
        ])
    return _write_workbook(filename, TRANSACTION_HEADERS, rows)


def export_demands(queryset, filename):
    rows = []
    for demand in queryset:
        rows.append([
            demand.item.item_id or '',
            demand.item.material.name,
            demand.item.uom,
            demand.item.sheet_packing_pcs,
            demand.month_str or '',
            demand.sheet_qty_pcs,
            demand.pkt_rim_qty,
        ])
    return _write_workbook(filename, DEMAND_HEADERS, rows)


def export_items(queryset, filename):
    rows = []
    for item in queryset:
        rows.append([
            item.item_id or '',
            item.material.name,
            item.uom,
            item.sheet_packing_pcs,
            float(item.unit_cost),
            item.safety_stock,
            item.max_stock_level,
            item.lead_time_days,
        ])
    return _write_workbook(filename, ITEM_HEADERS, rows)


def export_item_wise_consumption(rows, filename='item_wise_monthly_consumption.xlsx'):
    data_rows = [
        [
            row['item_id'],
            row['item_type'],
            row['uom'],
            row['sheet_packing_pcs'],
            row['month'],
            row['sheet_qty_pcs'],
            row['pkt_rim_qty'],
            row['consumption_value'],
        ]
        for row in rows
    ]
    return _write_workbook(filename, ITEM_WISE_CONSUMPTION_HEADERS, data_rows)


def export_month_wise_consumption(rows, filename='month_wise_item_consumption.xlsx'):
    data_rows = [
        [
            row['month'],
            row['item_id'],
            row['item_type'],
            row['uom'],
            row['sheet_packing_pcs'],
            row['sheet_qty_pcs'],
            row['pkt_rim_qty'],
            row['consumption_value'],
        ]
        for row in rows
    ]
    return _write_workbook(filename, MONTH_WISE_CONSUMPTION_HEADERS, data_rows)


def export_kpi_dashboard(rows, filename='supply_chain_kpis.xlsx'):
    data_rows = [
        [
            row['item'].item_id or '',
            row['item'].material.name,
            row['abc_class'],
            row['fsn_class'],
            row['closing'],
            row['monthly_demand'],
            row['variable_daily_demand'],
            row['reorder_point'],
            row['item'].safety_stock,
            row['inventory_value'],
            row['inventory_turnover'],
            row['annual_consumption_value'],
            'Yes' if row['stockout'] else 'No',
            'Yes' if row['reorder'] else 'No',
            'Yes' if row['safety_stock_alert'] else 'No',
            'Yes' if row['overstock'] else 'No',
            'Yes' if row['dead_stock'] else 'No',
            'Yes' if row['slow_moving'] else 'No',
            row['inventory_accuracy'] if row.get('inventory_accuracy') is not None else '',
        ]
        for row in rows
    ]
    return _write_workbook(filename, KPI_HEADERS, data_rows)


def export_physical_counts(queryset, filename='physical_stock_counts.xlsx'):
    data_rows = []
    for count in queryset:
        variance = count.physical_sheet_qty - count.system_sheet_qty
        data_rows.append([
            count.count_date.isoformat() if count.count_date else '',
            count.item.item_id or '',
            count.item.material.name,
            count.physical_sheet_qty,
            count.system_sheet_qty,
            variance,
            float(count.accuracy_percent),
            count.notes,
        ])
    return _write_workbook(filename, PHYSICAL_COUNT_HEADERS, data_rows)


def import_transactions(transaction_type, upload_file):
    rows = _read_rows(upload_file)
    created = 0
    skipped = 0

    for row in rows:
        item = _get_item_by_id(row.get('Item ID'))
        if not item:
            skipped += 1
            continue

        StockTransaction.objects.create(
            item=item,
            transaction_type=transaction_type,
            month_str=(row.get('Month') or '').strip() or None,
            date=_parse_date(row.get('Date')),
            gin_jc=(row.get('GIN / JC') or row.get('GIN/JC') or '').strip() or None,
            sheet_qty_pcs=int(row.get('Sheet Qty/Pcs') or 0),
            pkt_rim_qty=int(row.get('Pkt/Rim Qty') or 0),
        )
        created += 1

    return created, skipped


def import_demands(upload_file):
    rows = _read_rows(upload_file)
    created = 0
    skipped = 0

    for row in rows:
        item = _get_item_by_id(row.get('Item ID'))
        if not item:
            skipped += 1
            continue

        StockDemand.objects.create(
            item=item,
            month_str=(row.get('Month') or '').strip() or None,
            sheet_qty_pcs=int(row.get('Sheet Qty/Pcs') or 0),
            pkt_rim_qty=int(row.get('Pkt/Rim Qty') or 0),
        )
        created += 1

    return created, skipped
