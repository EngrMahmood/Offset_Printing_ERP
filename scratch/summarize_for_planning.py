import os
import sys
import django
import json
from datetime import datetime

# Set up Django environment
sys.path.append(r"d:\Development\Offset_Printing_ERP")
os.environ.setdefault("DJANGO_SETTINGS_MODULE", "Offset_ERP.settings")
django.setup()

from core.models import Machine, JobCard, Production
from planning.models import PlanningJob, SkuRecipe
from supply_chain.models import RawMaterialSku
from printing_plates.models import PlateRequest
from supply_chain.services import build_dashboard_data

# Create output summary path
summary_path = r"d:\Development\Offset_Printing_ERP\scratch\erp_export\erp_planning_summary.txt"

# 1. Fetch Machine Master
machines_list = []
for m in Machine.objects.filter(is_active=True):
    machines_list.append({
        "Name": m.name,
        "Speed_Imp_per_Hour": m.standard_impressions_per_hour,
        "Setup_Min_per_Color": m.standard_setup_minutes_per_color,
        "Plate_Life": m.plate_life_impressions
    })

# 2. Fetch Raw Materials Inventory (Closing Stock)
inventory_data = build_dashboard_data()
materials_list = []
for inv in inventory_data:
    item = inv['item']
    materials_list.append({
        "Material_SKU": item.sku,
        "Material_Name": item.material.name,
        "Sheet_Size": item.purchase_sheet_size,
        "Closing_Stock": inv['closing'],
        "Daily_Demand": inv['variable_daily_demand'],
        "Safety_Stock": item.safety_stock,
        "Unit_Cost": float(item.unit_cost)
    })

# 3. Fetch Active Job Cards (Exclude Completed/Closed/Archived unless they are in progress)
# We want open, released, planning approved, pending qc, and in production jobs
active_statuses = [
    'draft', 'pending_data', 'planning_approved', 'pending_qc', 'qc_approved',
    'pending_pm_approval', 'production_approved', 'released', 'in_production'
]
jobs_list = []
for jc in JobCard.objects.filter(status__in=active_statuses, is_active=True):
    # Find plate status for this SKU/job
    plate_statuses = list(PlateRequest.objects.filter(job_card=jc).values_list('status', flat=True))
    plate_status = plate_statuses[0] if plate_statuses else "No Request"

    jobs_list.append({
        "JC_No": jc.job_card_no,
        "PO_No": jc.PO_No,
        "SKU": jc.SKU,
        "Material": jc.material.name if jc.material else "N/A",
        "Colors": jc.total_colors or jc.number_of_colors,
        "Color_Spec": jc.colour,
        "Order_Qty": jc.order_qty,
        "Planned_Sheets": jc.total_sheets_planned,
        "Wastage_Sheets": jc.wastage,
        "Ups": jc.ups,
        "Sheet_Size": jc.print_sheet_size,
        "Target_Machine": jc.machine_name.name if jc.machine_name else (jc.planning_job.machine_name if jc.planning_job else "Unassigned"),
        "Status": jc.status,
        "Job_Status_Calc": jc.job_status,
        "Plate_Status": plate_status,
        "Required_Date": jc.po_date.isoformat() if jc.po_date else "N/A",
        "Est_Run_Time_Min": jc.estimated_run_time_minutes,
        "Est_Setup_Time_Min": jc.estimated_setup_time_minutes
    })

# Format the summary report in a clean, copy-pasteable Markdown format
summary_lines = []
summary_lines.append("# ERP Production Planning & Scheduling Summary")
summary_lines.append(f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")

summary_lines.append("## 1. Machine Capacities & Master")
summary_lines.append("| Machine Name | Speed (Impressions/Hr) | Setup (Min/Color) | Plate Life |")
summary_lines.append("| --- | --- | --- | --- |")
for m in machines_list:
    summary_lines.append(f"| {m['Name']} | {m['Speed_Imp_per_Hour']} | {m['Setup_Min_per_Color']} | {m['Plate_Life']} |")
summary_lines.append("")

summary_lines.append("## 2. Raw Material Inventory Status")
summary_lines.append("| SKU | Material Type | Sheet Size | Stock Qty | Daily Demand | Safety Stock | Unit Cost |")
summary_lines.append("| --- | --- | --- | --- | --- | --- | --- |")
for mat in materials_list:
    summary_lines.append(f"| {mat['Material_SKU']} | {mat['Material_Name']} | {mat['Sheet_Size']} | {mat['Closing_Stock']} | {mat['Daily_Demand']} | {mat['Safety_Stock']} | ${mat['Unit_Cost']:.2f} |")
summary_lines.append("")

summary_lines.append("## 3. Active & Pending Production Jobs")
summary_lines.append("| JC No | PO No | SKU | Material | Colors | Color Spec | Order Qty | Planned Sheets | Ups | Sheet Size | Machine | Status | Job Status | Plate Status | Req Date | Est Run (m) | Est Setup (m) |")
summary_lines.append("| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |")
for j in jobs_list:
    summary_lines.append(
        f"| {j['JC_No']} | {j['PO_No']} | {j['SKU']} | {j['Material']} | {j['Colors']} | {j['Color_Spec']} | "
        f"{j['Order_Qty']} | {j['Planned_Sheets']} | {j['Ups']} | {j['Sheet_Size']} | {j['Target_Machine']} | "
        f"{j['Status']} | {j['Job_Status_Calc']} | {j['Plate_Status']} | {j['Required_Date']} | "
        f"{j['Est_Run_Time_Min']} | {j['Est_Setup_Time_Min']} |"
    )
summary_lines.append("")

summary_lines.append("## 4. Key Planning Instructions for AI Scheduling")
summary_lines.append("1. **Same-SKU Consolidation**: Look for matching SKU codes across different PO/JC numbers to merge jobs or schedule them consecutively to eliminate setup times.")
summary_lines.append("2. **Material Batching**: Run jobs using the same material grade and sheet size sequentially to minimize substrate changeover delays.")
summary_lines.append("3. **Plate / Color Sequence**: Align jobs by color count (e.g. light to dark ink) and reuse plate sets if they share the same Artwork/AWC ID.")
summary_lines.append("4. **Machine Balancing**: Balance jobs across eligible machines. Ensure bottleneck machines are prioritized for high-priority/urgent jobs.")

# Write to file
with open(summary_path, 'w', encoding='utf-8') as f:
    f.write("\n".join(summary_lines))

print(f"Summary generated at: {summary_path}")
