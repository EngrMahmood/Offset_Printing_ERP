import csv
import re
from dateutil import parser
from datetime import datetime


from .models import JobCard, Material, Machine, Department


# ----------------------------
# NORMALIZER (CORE ENGINE)
# ----------------------------
def normalize(value):
    if not value:
        return ""
    return re.sub(r'[^a-zA-Z0-9]', '', str(value)).lower()


def clean(row, key):
    return (row.get(key) or "").strip()


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

    # 1. exact match
    if key in cache:
        return cache[key]

    # 2. partial/flexible match
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

    decoded_file = file.read().decode('utf-8').splitlines()
    reader = csv.DictReader(decoded_file)

    jobcards = []
    errors = []

    success_count = 0
    error_count = 0

    # ----------------------------
    # LOAD ALL MASTER DATA ONCE
    # ----------------------------
    MATERIAL_MAP = build_cache(Material)
    MACHINE_MAP = build_cache(Machine)
    DEPARTMENT_MAP = build_cache(Department)

    for index, row in enumerate(reader, start=2):

        try:
            job_card_no = clean(row, "job_card_no")

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

            # duplicate check
            if JobCard.objects.filter(job_card_no=job_card_no).exists():
                errors.append({
                    "row": index,
                    "errors": f"Duplicate Job Card: {job_card_no}"
                })
                error_count += 1
                continue

            # ----------------------------
            # SMART FK RESOLUTION
            # ----------------------------
            material = resolve(
                clean(row, "material"),
                MATERIAL_MAP,
                "Material",
                errors,
                index
            )

            machine = resolve(
                clean(row, "machine_name"),
                MACHINE_MAP,
                "Machine",
                errors,
                index
            )

            department = resolve(
                clean(row, "department"),
                DEPARTMENT_MAP,
                "Department",
                errors,
                index
            )

            if not (material and machine and department):
                error_count += 1
                continue

            # ----------------------------
            # CREATE JOBCARD OBJECT
            # ----------------------------
            jobcards.append(JobCard(
                 job_card_no=job_card_no,

                month=clean(row, "month"),
                po_date=parse_date(clean(row, "po_date")),
                PO_No=parse_int(clean(row, "PO_No")),
                SKU=clean(row, "SKU"),

                material=material,
                colour=parse_int(clean(row, "colour")),
                application=clean(row, "application"),

                order_qty=parse_int(clean(row, "order_qty")),
                ups=parse_int(clean(row, "ups")),

                print_sheet_size=clean(row, "print_sheet_size"),
                wastage=parse_int(clean(row, "wastage")),
                actual_sheet_required=parse_int(clean(row, "actual_sheet_required")),

                purchase_sheet_size=clean(row, "purchase_sheet_size"),
                purchase_sheet_ups=parse_int(clean(row, "purchase_sheet_ups")),

                remarks=clean(row, "remarks"),
                destination=clean(row, "destination"),
                machine_name=machine,          # ✅ FIXED
                department=department,        # ✅ FIXED

                die_cutting=clean(row, "die_cutting")

                 ))

            success_count += 1

        except Exception as e:
            errors.append({
                "row": index,
                "errors": str(e)
            })
            error_count += 1

    # ----------------------------
    # BULK INSERT
    # ----------------------------
    JobCard.objects.bulk_create(jobcards)

    return {
        "success_count": success_count,
        "error_count": error_count,
        "errors": errors
    }