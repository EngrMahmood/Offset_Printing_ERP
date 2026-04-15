import csv
import re
from io import BytesIO
from dateutil import parser
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


def parse_date(value):
    try:
        return parser.parse(value, dayfirst=True).date()
    except:
        return None


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
# TEMPLATE HEADERS (AUTO-GENERATED)
# ----------------------------
def get_template_headers():
    """Dynamically generate template headers based on JobCard fields"""
    return [
        'Job Card Number',
        'SKU',
        'PO Number',
        'PO Date (dd/mm/yyyy)',
        'Month',
        'Material',
        'Colour (Count)',
        'Application',
        'Order Quantity (pcs)',
        'UPS (Units per Sheet)',
        'Print Sheet Size',
        'Total Impressions Required',
        'Wastage (%)',
        'Purchase Sheet Size',
        'Purchase Sheet UPS',
        'Remarks',
        'Destination',
        'Machine',
        'Department',
        'Die Cutting (Yes/No)'
    ]


def get_template_example():
    """Generate example row for template"""
    return [
        'JC-26-1001',
        'SKU-01',
        'PO-7788',
        '15/04/2026',
        'April',
        'Bleach230',
        '4',
        'UV',
        '10000',
        '12',
        '20x30',
        '400',  # total impressions = sheets x colours
        '5',
        '20x30',
        '6',
        'Urgent job',
        'SITE 1',
        'GTO 1A',
        'Pillow',
        'Yes'
    ]


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
        'wastage': ['wastage', 'wastage (%)', 'waste %', 'waste percentage'],
        'purchase_sheet_size': ['purchase sheet size', 'purchase size', 'purchase_sheet_size'],
        'purchase_sheet_ups': ['purchase sheet ups', 'purchase ups', 'purchase_sheet_ups'],
        'remarks': ['remarks', 'notes', 'comments'],
        'destination': ['destination', 'delivery location'],
        'machine_name': ['machine', 'machine name', 'press'],
        'department': ['department', 'dept'],
        'die_cutting': ['die cutting', 'die cut', 'die cutting (yes/no)', 'die_cutting'],
        'month': ['month']
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
def process_jobcard_upload(file):
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
    required_columns = ['job_card_no', 'machine_name', 'department', 'material', 'order_qty', 'ups']
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

            if ups <= 0:
                errors.append({
                    "row": index,
                    "errors": "UPS must be greater than 0"
                })
                error_count += 1
                continue

            # ----------------------------
            # CREATE OBJECT
            # ----------------------------
            jobcards.append(JobCard(
                job_card_no=job_card_no,

                month=get_field_value(row, 'month', column_mapping),
                po_date=parse_date(get_field_value(row, 'po_date', column_mapping)),
                PO_No=get_field_value(row, 'po_no', column_mapping),
                SKU=get_field_value(row, 'sku', column_mapping),

                material=material,
                colour=parse_int(get_field_value(row, 'colour', column_mapping)),
                application=get_field_value(row, 'application', column_mapping),

                order_qty=order_qty,
                ups=ups,

                print_sheet_size=get_field_value(row, 'print_sheet_size', column_mapping),
                total_impressions_required=parse_int(get_field_value(row, 'total_impressions_required', column_mapping)),
                wastage=parse_int(get_field_value(row, 'wastage', column_mapping)),

                purchase_sheet_size=get_field_value(row, 'purchase_sheet_size', column_mapping),
                purchase_sheet_ups=parse_int(get_field_value(row, 'purchase_sheet_ups', column_mapping)),

                remarks=get_field_value(row, 'remarks', column_mapping),
                destination=get_field_value(row, 'destination', column_mapping),

                machine_name=machine,
                department=department,

                die_cutting=get_field_value(row, 'die_cutting', column_mapping)
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
        'PO Number',
        'PO Date',
        'Month',
        'Material',
        'Colour',
        'Application',
        'Order Quantity',
        'Ups',
        'Print Sheet Size',
        'Wastage (%)',
        'Actual Sheet Required',
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
        'PO-7788',
        '4/13/2026',
        'April',
        'Bleach230',
        '4',
        'UV',
        '10000',
        '12',
        '20x30',
        '5',
        '10500',
        '20x30',
        '6',
        'Urgent job',
        'SITE 1',
        'GTO 1A',
        'Pillow',
        'Yes'
    ]