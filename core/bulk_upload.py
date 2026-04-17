import csv
import re
import math
import calendar
from io import BytesIO
from dateutil import parser
from datetime import datetime, date
from django.db import transaction

from .models import JobCard, Material, Machine, Department

try:
    import openpyxl
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False


# ----------------------------
# NORMALIZER (CORE ENGINE)
# ----------------------------
def normalize(value):
    if not value:
        return ""
    return re.sub(r'[^a-zA-Z0-9]', '', str(value)).lower()


def clean(row, key):
    value = row.get(key)
    if value is None:
        return ""
    return str(value).strip()


def parse_int(value):
    try:
        return int(value)
    except:
        return 0


def parse_bool(value):
    return str(value).strip().lower() in ["yes", "true", "1"]


def parse_month_hint(value):
    """Return month number (1-12) from numeric/text month input, else None."""
    if value is None:
        return None
    raw = str(value).strip().lower()
    if not raw:
        return None

    if raw.isdigit():
        month_num = int(raw)
        if 1 <= month_num <= 12:
            return month_num
        return None

    month_lookup = {m.lower(): i for i, m in enumerate(calendar.month_name) if m}
    month_abbr_lookup = {m.lower(): i for i, m in enumerate(calendar.month_abbr) if m}

    if raw in month_lookup:
        return month_lookup[raw]
    if raw in month_abbr_lookup:
        return month_abbr_lookup[raw]
    return None


def parse_date(value, month_hint=None):
    if value is None:
        return None

    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value

    raw = str(value).strip()
    if not raw:
        return None

    hinted_month = parse_month_hint(month_hint)

    # Preferred safe format (no ambiguity)
    if re.fullmatch(r'\d{4}-\d{2}-\d{2}', raw):
        parsed = datetime.strptime(raw, '%Y-%m-%d').date()
        if hinted_month and parsed.month != hinted_month:
            raise ValueError(
                f"PO Date '{raw}' does not match Month column '{month_hint}'."
            )
        return parsed

    # Handle dd-mm-yyyy / mm-dd-yyyy and dd/mm/yyyy / mm/dd/yyyy safely.
    compact_match = re.fullmatch(r'(\d{1,2})[-/](\d{1,2})[-/](\d{4})', raw)
    if compact_match:
        first = int(compact_match.group(1))
        second = int(compact_match.group(2))
        year = int(compact_match.group(3))

        # Unambiguous day-month.
        if first > 12 and 1 <= second <= 12:
            return date(year, second, first)
        # Unambiguous month-day.
        if second > 12 and 1 <= first <= 12:
            return date(year, first, second)
        # Ambiguous (both <= 12): use month hint if provided, otherwise reject.
        if first <= 12 and second <= 12:
            if hinted_month:
                if first == hinted_month and second != hinted_month:
                    # mm-dd-yyyy
                    return date(year, first, second)
                if second == hinted_month and first != hinted_month:
                    # dd-mm-yyyy
                    return date(year, second, first)
                raise ValueError(
                    f"Ambiguous date '{raw}' does not align with Month column '{month_hint}'. Use YYYY-MM-DD."
                )
            raise ValueError(
                f"Ambiguous date '{raw}'. Provide Month column or use ISO format YYYY-MM-DD (example: 2026-03-12)."
            )

    # Fallback for long month names etc.
    parsed = parser.parse(raw, dayfirst=True).date()
    if hinted_month and parsed.month != hinted_month:
        raise ValueError(
            f"PO Date '{raw}' does not match Month column '{month_hint}'."
        )
    return parsed


def normalize_colour_value(value):
    """Convert compact notation like 1+1 into readable front/back text."""
    raw = (value or '').strip()
    if not raw:
        return None

    match = re.fullmatch(r'(\d+)\s*\+\s*(\d+)', raw)
    if not match:
        return raw

    front = int(match.group(1))
    back = int(match.group(2))
    front_label = 'color' if front == 1 else 'colors'
    back_label = 'color' if back == 1 else 'colors'
    return f"{front} {front_label} front and {back} {back_label} back"


def extract_total_colors(value):
    """Extract total color units from supported color formats."""
    raw = (value or '').strip().lower()
    if not raw:
        return 0

    simple_plus = re.fullmatch(r'(\d+)\s*\+\s*(\d+)', raw)
    if simple_plus:
        return int(simple_plus.group(1)) + int(simple_plus.group(2))

    normalized_text = re.fullmatch(r'(\d+)\s*colors?\s*front\s*and\s*(\d+)\s*colors?\s*back', raw)
    if normalized_text:
        return int(normalized_text.group(1)) + int(normalized_text.group(2))

    digits_only = re.fullmatch(r'(\d+)', raw)
    if digits_only:
        return int(digits_only.group(1))

    return 0


def compute_estimated_minutes(total_impressions_required, machine, colour_value):
    """Return (run_minutes, setup_minutes, total_minutes) from machine+impressions+colors."""
    impressions = int(total_impressions_required or 0)
    if impressions <= 0 or not machine:
        return (None, None, None)

    speed = float(machine.standard_impressions_per_hour or 0)
    if speed <= 0:
        return (None, None, None)

    run_minutes = (impressions / speed) * 60
    total_colors = extract_total_colors(colour_value)
    setup_per_color = float(machine.standard_setup_minutes_per_color or 0)
    setup_minutes = total_colors * setup_per_color
    total_minutes = run_minutes + setup_minutes
    return (
        round(run_minutes, 2),
        round(setup_minutes, 2),
        round(total_minutes, 2),
    )


# ----------------------------
# MASTER CACHE BUILDER
# ----------------------------
def build_cache(model):
    cache = {}
    for obj in model.objects.all():
        key = normalize(getattr(obj, "name", ""))
        if key:
            cache[key] = obj
    return cache


# ----------------------------
# FILE READER (CSV & EXCEL)
# ----------------------------
def read_csv_file(file):
    """Read CSV file and return list of dictionaries"""
    decoded_file = file.read().decode('utf-8').splitlines()
    reader = csv.DictReader(decoded_file)
    return list(reader)


def read_excel_file(file):
    """Read Excel file and return list of dictionaries"""
    if not EXCEL_AVAILABLE:
        raise ImportError("openpyxl not installed. Install with: pip install openpyxl")
    
    file.seek(0)
    workbook = openpyxl.load_workbook(file)
    worksheet = workbook.active
    
    rows = list(worksheet.iter_rows(values_only=True))
    if not rows:
        return []
    
    headers = [str(h).strip() if h else f"Column_{i}" for i, h in enumerate(rows[0])]
    data = []
    for row in rows[1:]:
        row_dict = {headers[i]: row[i] for i in range(len(headers))}
        data.append(row_dict)
    
    return data


def read_upload_file(file):
    """Auto-detect file type and read accordingly"""
    file_name = file.name.lower()
    if file_name.endswith('.xlsx') or file_name.endswith('.xls'):
        return read_excel_file(file)
    return read_csv_file(file)


# ----------------------------
# SMART COLUMN MAPPER
# ----------------------------
def normalize_headers(raw_headers):
    """Map flexible column names to standard field names"""
    normalized = {}

    # Define multiple acceptable variations for each field
    field_mappings = {
        'job_card_no': ['job card number', 'job card no', 'jc number', 'jobcard number', 'job card'],
        'month': ['month', 'month name', 'po month'],
        'po_no': ['po number', 'po no', 'po', 'pono', 'po_no'],
        'po_date': ['po date', 'po_date', 'date'],
        'sku': ['sku', 'product code', 'product'],
        'material': ['material', 'material name'],
        'colour': ['colour', 'color', 'colour (count)', 'color count', 'colours'],
        'application': ['application', 'application type'],
        'order_qty': ['order quantity', 'order qty', 'quantity', 'order quantity (pcs)', 'orderqty'],
        'ups': ['ups', 'units per sheet', 'units', 'ups (units per sheet)'],
        'print_sheet_size': ['print sheet size', 'sheet size', 'print size', 'print_sheet_size'],
        'total_impressions_required': ['total impressions required', 'impressions required', 'impressions', 'total impressions'],
        'wastage': ['wastage', 'wastage sheets', 'waste sheets', 'wastage (sheets)'],
        'wastage_percent': ['wastage (%)', 'waste %', 'waste percentage', 'wastage percent'],
        'actual_sheet_required': ['actual sheet required', 'actual sheets required', 'actual sheet'],
        'purchase_sheet_size': ['purchase sheet size', 'purchase size', 'purchase_sheet_size'],
        'purchase_sheet_ups': ['purchase sheet ups', 'purchase ups', 'purchase_sheet_ups'],
        'remarks': ['remarks', 'notes', 'comments'],
        'destination': ['destination', 'delivery location'],
        'machine_name': ['machine', 'machine name', 'press'],
        'department': ['department', 'dept'],
        'die_cutting': ['die cutting', 'die cut', 'die cutting (yes/no)', 'die_cutting']
    }

    for raw_header in raw_headers:
        normalized_header = normalize(raw_header)
        best_match = None
        best_score = 0

        for field, variations in field_mappings.items():
            for variation in variations:
                normalized_variation = normalize(variation)
                if normalized_header == normalized_variation:
                    score = len(normalized_variation) + 100
                elif normalized_variation in normalized_header:
                    score = len(normalized_variation) + 10
                elif normalized_header in normalized_variation:
                    score = len(normalized_header)
                else:
                    score = 0

                if score > best_score:
                    best_score = score
                    best_match = field

        if best_match:
            normalized[raw_header] = best_match

    return normalized


def get_field_value(row, field_name, column_mapping):
    """Get value from row using flexible column names"""
    for raw_header, mapped_field in column_mapping.items():
        if mapped_field == field_name and raw_header in row:
            return clean(row, raw_header)

    # Fallback: support direct field name headers
    if field_name in row:
        return clean(row, field_name)
    normalized_field = normalize(field_name)
    for raw_header in row.keys():
        if normalize(raw_header) == normalized_field:
            return clean(row, raw_header)

    return ""


def calculate_wastage_sheets(row, column_mapping, order_qty, ups):
    """Convert incoming wastage inputs into sheet count for JobCard.wastage."""
    wastage_sheets_raw = get_field_value(row, 'wastage', column_mapping)
    if str(wastage_sheets_raw).strip():
        return max(parse_int(wastage_sheets_raw), 0)

    required_sheets = (order_qty / ups) if ups > 0 else 0

    actual_sheet_required_raw = get_field_value(row, 'actual_sheet_required', column_mapping)
    if str(actual_sheet_required_raw).strip():
        actual_sheet_required = max(parse_int(actual_sheet_required_raw), 0)
        return max(actual_sheet_required - math.floor(required_sheets), 0)

    wastage_percent_raw = get_field_value(row, 'wastage_percent', column_mapping)
    if str(wastage_percent_raw).strip():
        cleaned_percent = str(wastage_percent_raw).replace('%', '').strip()
        try:
            wastage_percent = max(float(cleaned_percent), 0)
        except Exception:
            wastage_percent = 0
        return int(round(required_sheets * (wastage_percent / 100)))

    return 0


# ----------------------------
# SMART RESOLVER (UNIVERSAL)
# ----------------------------
def resolve(value, cache, label, errors, row_no):
    key = normalize(value)

    if not key:
        errors.append({
            "row": row_no,
            "errors": f"{label} is missing"
        })
        return None

    # exact match
    if key in cache:
        return cache[key]

    # partial match
    for k, v in cache.items():
        if key in k or k in key:
            return v

    errors.append({
        "row": row_no,
        "errors": f"Invalid {label}: {value}"
    })
    return None


# ----------------------------
# MAIN IMPORT FUNCTION
# ----------------------------
def process_jobcard_upload(file, uploaded_by=None):
    try:
        # Read file (auto-detect CSV/Excel)
        rows = read_upload_file(file)
    except Exception as e:
        return {
            "success_count": 0,
            "error_count": 1,
            "errors": [{"row": 0, "errors": f"File reading error: {str(e)}"}]
        }

    if not rows:
        return {
            "success_count": 0,
            "error_count": 1,
            "errors": [{"row": 0, "errors": "No data found in file"}]
        }

    jobcards = []
    errors = []
    success_count = 0
    error_count = 0

    # ----------------------------
    # LOAD MASTER DATA
    # ----------------------------
    MATERIAL_MAP = build_cache(Material)
    MACHINE_MAP = build_cache(Machine)
    DEPARTMENT_MAP = build_cache(Department)
    
    # Map column names to standard fields
    column_mapping = normalize_headers(rows[0].keys()) if rows else {}
    required_columns = [
        'job_card_no',
        'sku',
        'po_date',
        'machine_name',
        'department',
        'material',
        'order_qty',
        'ups',
        'colour',
        'total_impressions_required',
    ]
    missing_columns = [col for col in required_columns if col not in column_mapping.values()]
    if missing_columns:
        return {
            "success_count": 0,
            "error_count": 1,
            "errors": [{
                "row": 0,
                "errors": f"Missing required columns: {', '.join(missing_columns)}"
            }]
        }

    for index, row in enumerate(rows, start=2):
        try:
            # Get values using flexible column mapping
            job_card_no = get_field_value(row, 'job_card_no', column_mapping)

            # ----------------------------
            # VALIDATION: JOB CARD
            # ----------------------------
            if not job_card_no:
                errors.append({
                    "row": index,
                    "errors": "Job Card Number missing"
                })
                error_count += 1
                continue

            if JobCard.objects.filter(job_card_no=job_card_no).exists():
                errors.append({
                    "row": index,
                    "errors": f"Duplicate Job Card: {job_card_no}"
                })
                error_count += 1
                continue

            # ----------------------------
            # FK RESOLUTION
            # ----------------------------
            material_value = get_field_value(row, 'material', column_mapping)
            machine_value = get_field_value(row, 'machine_name', column_mapping)
            department_value = get_field_value(row, 'department', column_mapping)
            
            material = resolve(material_value, MATERIAL_MAP, "Material", errors, index)
            machine = resolve(machine_value, MACHINE_MAP, "Machine", errors, index)
            department = resolve(department_value, DEPARTMENT_MAP, "Department", errors, index)

            if not (material and machine and department):
                error_count += 1
                continue

            # ----------------------------
            # FIELD PARSING
            # ----------------------------
            order_qty = parse_int(get_field_value(row, 'order_qty', column_mapping))
            ups = parse_int(get_field_value(row, 'ups', column_mapping))
            sku_value = get_field_value(row, 'sku', column_mapping)
            colour_raw = get_field_value(row, 'colour', column_mapping)
            colour_value = normalize_colour_value(colour_raw)
            total_impressions_required = parse_int(get_field_value(row, 'total_impressions_required', column_mapping))
            month_raw = get_field_value(row, 'month', column_mapping)
            po_date_raw = get_field_value(row, 'po_date', column_mapping)
            month_hint_num = parse_month_hint(month_raw)
            month_hint_name = calendar.month_name[month_hint_num] if month_hint_num else None

            if not po_date_raw:
                errors.append({
                    "row": index,
                    "errors": "PO Date is required"
                })
                error_count += 1
                continue

            try:
                po_date_value = parse_date(po_date_raw, month_hint=month_raw)
            except Exception as exc:
                errors.append({
                    "row": index,
                    "errors": f"PO Date error: {str(exc)}"
                })
                error_count += 1
                continue
            month_value = po_date_value.strftime('%B') if po_date_value else month_hint_name

            missing_row_fields = []
            if not sku_value:
                missing_row_fields.append('SKU')
            if order_qty <= 0:
                missing_row_fields.append('Order Qty (>0)')
            if ups <= 0:
                missing_row_fields.append('UPS (>0)')
            if not colour_value:
                missing_row_fields.append('Colour')
            if total_impressions_required <= 0:
                missing_row_fields.append('Total Impressions Required (>0)')

            if missing_row_fields:
                errors.append({
                    "row": index,
                    "errors": f"Missing/invalid mandatory fields: {', '.join(missing_row_fields)}"
                })
                error_count += 1
                continue

            if ups <= 0:
                errors.append({
                    "row": index,
                    "errors": "UPS must be greater than 0"
                })
                error_count += 1
                continue

            estimated_run_minutes, estimated_setup_minutes, estimated_total_minutes = compute_estimated_minutes(
                total_impressions_required,
                machine,
                colour_value,
            )

            # ----------------------------
            # CREATE OBJECT
            # ----------------------------
            jobcards.append(JobCard(
                job_card_no=job_card_no,

                month=month_value,
                po_date=po_date_value,
                PO_No=get_field_value(row, 'po_no', column_mapping),
                SKU=sku_value,

                material=material,
                colour=colour_value,
                application=get_field_value(row, 'application', column_mapping),

                order_qty=order_qty,
                ups=ups,

                print_sheet_size=get_field_value(row, 'print_sheet_size', column_mapping),
                total_impressions_required=total_impressions_required,
                estimated_run_time_minutes=estimated_run_minutes,
                estimated_setup_time_minutes=estimated_setup_minutes,
                estimated_total_time_minutes=estimated_total_minutes,
                wastage=calculate_wastage_sheets(row, column_mapping, order_qty, ups),

                purchase_sheet_size=get_field_value(row, 'purchase_sheet_size', column_mapping),
                purchase_sheet_ups=parse_int(get_field_value(row, 'purchase_sheet_ups', column_mapping)),

                remarks=get_field_value(row, 'remarks', column_mapping),
                destination=get_field_value(row, 'destination', column_mapping),

                machine_name=machine,
                department=department,

                die_cutting=get_field_value(row, 'die_cutting', column_mapping),
                created_by=uploaded_by,
            ))

            success_count += 1

        except Exception as e:
            errors.append({
                "row": index,
                "errors": str(e)
            })
            error_count += 1

    # ----------------------------
    # BULK INSERT (SAFE)
    # ----------------------------
    if jobcards:
        with transaction.atomic():
            JobCard.objects.bulk_create(jobcards)

    return {
        "success_count": success_count,
        "error_count": error_count,
        "errors": errors
    }


# ----------------------------
# TEMPLATE GENERATION
# ----------------------------
def get_template_headers():
    """Get the headers for the job card template"""
    return [
        'JC Number',
        'SKU',
        'Month',
        'PO Number',
        'PO Date (YYYY-MM-DD)',
        'Material',
        'Colour',
        'Application',
        'Order Quantity',
        'Ups',
        'Print Sheet Size',
        'Total Impressions Required',
        'Wastage (sheets)',
        'Purchase Sheet Size',
        'Purchase Sheet Ups',
        'Remarks',
        'Destination',
        'Machine',
        'Department',
        'Die Cutting (Yes/No)'
    ]


def get_template_example():
    """Get example data for the job card template"""
    return [
        'JC-26-1001',
        'SKU-01',
        'March',
        'PO-7788',
        '2026-04-13',
        'Bleach230',
        '4',
        'UV',
        '10000',
        '12',
        '20x30',
        '24000',
        '50',
        '20x30',
        '6',
        'Urgent job',
        'SITE 1',
        'GTO 1A',
        'Pillow',
        'Yes'
    ]