import os
import sys
import django
import csv
import json
import zipfile
from datetime import datetime, date, time
from decimal import Decimal
from openpyxl import Workbook


# Set up Django environment
sys.path.append(r"d:\Development\Offset_Printing_ERP")
os.environ.setdefault("DJANGO_SETTINGS_MODULE", "Offset_ERP.settings")
django.setup()

from django.db.models import Sum
from core.models import Machine, Department, Material, PrintColor, JobCard, Production, ShiftConfig, MachineWorkSchedule
from planning.models import PlanningJob, SkuRecipe
from supply_chain.models import RawMaterialSku, StockTransaction, StockDemand, PhysicalStockCount
from printing_plates.models import PlateRequest
from supply_chain.services import build_dashboard_data

# Create output directory
output_dir = r"d:\Development\Offset_Printing_ERP\scratch\erp_export"
os.makedirs(output_dir, exist_ok=True)

def clean_row(row):
    cleaned = {}
    for k, v in row.items():
        if isinstance(v, (datetime, date, time)):
            cleaned[k] = v.isoformat()
        elif isinstance(v, Decimal):
            cleaned[k] = float(v)
        else:
            cleaned[k] = v
    return cleaned

# Let's collect data for all models
data_sets = {}

# 1. Machines
machines = Machine.objects.all().values(
    'id', 'name', 'standard_impressions_per_hour', 'standard_setup_minutes_per_color', 
    'plate_life_impressions', 'is_active'
)
data_sets['machines'] = [clean_row(m) for m in machines]

# 2. Departments
departments = Department.objects.all().values('id', 'name')
data_sets['departments'] = [clean_row(d) for d in departments]

# 3. Materials
materials = Material.objects.all().values('id', 'name')
data_sets['materials'] = [clean_row(m) for m in materials]

# 4. Print Colors
print_colors = PrintColor.objects.all().values('id', 'name', 'is_active', 'sort_order')
data_sets['print_colors'] = [clean_row(pc) for pc in print_colors]

# 5. Job Cards (Production Orders)
job_cards = JobCard.objects.all().values(
    'id', 'job_card_no', 'planning_job_id', 'month', 'po_date', 'PO_No', 'SKU', 
    'material_id', 'colour', 'application', 'order_qty', 'total_impressions_required', 
    'estimated_run_time_minutes', 'estimated_setup_time_minutes', 'estimated_total_time_minutes', 
    'production_tolerance_percent', 'ups', 'print_sheet_size', 'plate_set_no', 'wastage', 
    'total_sheet_quantity', 'total_colors', 'purchase_sheet_size', 'purchase_sheet_ups', 
    'remarks', 'destination', 'machine_name_id', 'department_id', 'die_cutting', 
    'is_print_job', 'is_active', 'created_at', 'updated_at', 'status'
)
# Add calculated properties/labels
jc_list = []
for jc_val in job_cards:
    jc_obj = JobCard.objects.get(pk=jc_val['id'])
    row = dict(jc_val)
    row['workflow_status'] = jc_obj.workflow_status
    row['workflow_status_label'] = jc_obj.workflow_status_label
    row['total_sheets_planned'] = jc_obj.total_sheets_planned
    row['total_production'] = jc_obj.total_production
    row['total_dispatch'] = jc_obj.total_dispatch
    row['balance_qty'] = jc_obj.balance_qty
    row['job_status'] = jc_obj.job_status
    jc_list.append(clean_row(row))
data_sets['job_cards'] = jc_list

# 6. Planning Jobs
planning_jobs = PlanningJob.objects.all().values(
    'id', 'jc_number', 'plan_month', 'plan_date', 'po_approval_date', 'delivery_date', 
    'po_number', 'sku', 'job_name', 'repeat_flag', 'material', 'color_spec', 'application', 
    'size_w_mm', 'size_h_mm', 'order_qty', 'print_pcs', 'ups', 'print_sheet_size', 
    'print_sheets', 'wastage_sheets', 'actual_sheet_required', 'purchase_sheet_size', 
    'purchase_sheet_ups', 'purchase_sheet_required', 'pkt_value', 'remarks', 'requirement', 
    'front_colors', 'back_colors', 'total_colors', 'total_mr_time_minutes', 'front_pass', 
    'back_pass', 'job_process_type', 'print_passes', 'planned_total_impressions', 'status', 
    'planning_stage', 'pr_reference', 'priority', 'is_active', 'is_on_hold', 'hold_reason', 
    'created_at', 'updated_at'
)
data_sets['planning_jobs'] = [clean_row(pj) for pj in planning_jobs]

# 7. SKU Recipes (SKU Master)
sku_recipes = SkuRecipe.objects.all().values(
    'id', 'sku', 'job_name', 'material', 'color_spec', 'application', 'product_type', 
    'machine_name', 'job_process_type', 'print_passes', 'plate_set_no', 'size_w_mm', 
    'size_h_mm', 'ups', 'print_sheet_size', 'purchase_sheet_size', 'purchase_sheet_ups', 
    'default_unit_cost', 'daily_demand', 'awc_no', 'die_cutting', 'notes', 'remarks', 
    'is_active', 'master_data_status'
)
data_sets['sku_recipes'] = [clean_row(sr) for sr in sku_recipes]

# 8. Raw Material SKUs (Material Master / Inventory Info)
raw_material_skus = RawMaterialSku.objects.all().values(
    'id', 'sku', 'material_id', 'purchase_sheet_size', 'uom', 'sheet_packing_pcs', 
    'unit_cost', 'safety_stock', 'max_stock_level', 'lead_time_days', 'is_active'
)
rm_list = []
# Get inventory status using the supply chain dashboard calculation service
inventory_dashboard = build_dashboard_data()
inventory_map = {item['item'].pk: item for item in inventory_dashboard}

for rm in raw_material_skus:
    row = dict(rm)
    inv_info = inventory_map.get(rm['id'])
    if inv_info:
        row['available_inventory'] = inv_info['closing']
        row['opening'] = inv_info['opening']
        row['receiving'] = inv_info['receiving']
        row['issuance'] = inv_info['issuance']
        row['adjustment'] = inv_info['adjustment']
        row['monthly_demand'] = inv_info['monthly_demand']
        row['daily_demand'] = inv_info['variable_daily_demand']
        row['forecast_days'] = inv_info['stock_forecast_days']
    else:
        row['available_inventory'] = 0
        row['opening'] = 0
        row['receiving'] = 0
        row['issuance'] = 0
        row['adjustment'] = 0
        row['monthly_demand'] = 0
        row['daily_demand'] = 0
        row['forecast_days'] = 0
    rm_list.append(clean_row(row))
data_sets['raw_material_skus'] = rm_list

# 9. Stock Transactions
stock_transactions = StockTransaction.objects.all().values(
    'id', 'raw_material_sku_id', 'transaction_type', 'source', 'job_card_id', 
    'month_str', 'date', 'sheet_qty_pcs', 'pkt_rim_qty', 'is_active', 'is_approved', 
    'created_at', 'updated_at'
)
data_sets['stock_transactions'] = [clean_row(st) for st in stock_transactions]

# 10. Stock Demands
stock_demands = StockDemand.objects.all().values(
    'id', 'raw_material_sku_id', 'month_str', 'sheet_qty_pcs', 'pkt_rim_qty', 
    'is_active', 'created_at', 'updated_at'
)
data_sets['stock_demands'] = [clean_row(sd) for sd in stock_demands]

# 11. Physical Stock Counts
physical_stock_counts = PhysicalStockCount.objects.all().values(
    'id', 'raw_material_sku_id', 'count_date', 'physical_sheet_qty', 'system_sheet_qty', 
    'accuracy_percent', 'notes', 'is_active', 'created_at', 'updated_at'
)
data_sets['physical_stock_counts'] = [clean_row(psc) for psc in physical_stock_counts]

# 12. Plate Requests
plate_requests = PlateRequest.objects.all().values(
    'id', 'planning_job_id', 'job_card_id', 'sku_recipe_id', 'machine_id', 'department_id', 
    'status', 'set_no', 'new_set_no', 'awc_no', 'die_cutting', 'plate_quantity', 
    'sets_required', 'plate_color', 'vendor', 'remarks', 'source', 'replacement_reason', 
    'damaged_colors', 'progress', 'challan', 'chalan_sign', 'box', 'requested_at', 
    'sent_at', 'received_at', 'created_at', 'updated_at'
)
data_sets['plate_requests'] = [clean_row(pr) for pr in plate_requests]

# 13. Shift Configurations
shift_configs = ShiftConfig.objects.all().values(
    'id', 'day_of_week', 'shift', 'effective_from', 'effective_to', 'net_hours'
)
data_sets['shift_configs'] = [clean_row(sc) for sc in shift_configs]

# 14. Machine Work Schedules
machine_work_schedules = MachineWorkSchedule.objects.all().values(
    'id', 'machine_id', 'day_of_week', 'shift', 'effective_from', 'effective_to', 'is_working'
)
data_sets['machine_work_schedules'] = [clean_row(mws) for mws in machine_work_schedules]

# 15. Productions
productions = Production.objects.all().values(
    'id', 'entry_type', 'job_card_id', 'date', 'shift', 'machine_id', 'output_sheets', 
    'waste_sheets', 'intermediate_pass', 'print_pass_number', 'waste_reason', 'impressions', 
    'packing_qty', 'sorting_waste_qty', 'planned_time', 'run_time', 'downtime_minutes', 
    'make_ready_time', 'downtime_category', 'counter_start', 'counter_end', 'start_time', 
    'end_time', 'status', 'operator_id', 'supervisor_id', 'is_active', 'created_at'
)
data_sets['productions'] = [clean_row(p) for p in productions]


# Write to JSON
json_path = os.path.join(output_dir, "erp_data_extract.json")
with open(json_path, 'w', encoding='utf-8') as f:
    json.dump(data_sets, f, indent=2)

# Write to individual CSV files
csv_files = []
for name, data in data_sets.items():
    if not data:
        continue
    csv_path = os.path.join(output_dir, f"{name}.csv")
    csv_files.append(csv_path)
    
    headers = list(data[0].keys())
    with open(csv_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=headers)
        writer.writeheader()
        for row in data:
            writer.writerow(row)

# Zip CSV files
zip_path = os.path.join(output_dir, "erp_data_csvs.zip")
with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
    for csv_file in csv_files:
        zipf.write(csv_file, os.path.basename(csv_file))

# Write to Excel Workbook (multiple sheets)
excel_path = os.path.join(output_dir, "erp_data_workbook.xlsx")
wb = Workbook()
# Remove default sheet
wb.remove(wb.active)

for name, data in data_sets.items():
    if not data:
        continue
    # Sheet names max length is 31 chars
    sheet_name = name[:30]
    ws = wb.create_sheet(title=sheet_name)
    headers = list(data[0].keys())
    ws.append(headers)
    for row in data:
        ws.append([row[h] for h in headers])

wb.save(excel_path)

print(f"Data extraction complete!")
print(f"JSON: {json_path}")
print(f"Excel: {excel_path}")
print(f"Zip of CSVs: {zip_path}")
