"""Declares which models sync to which Google Sheets tab.

Each entry: (dotted "app_label.ModelName", tab_name, serializer_function_name,
header_row). The serializer function lives in sheets_sync.serializers and
takes a model instance, returning a dict of {header: value}.

Phase 1 scope: Job Cards only. Planning/Production/Dispatch (Phase 2),
Maintenance (Phase 3), and Supply Chain (Phase 4) are added incrementally.
"""

JOB_CARD_HEADERS = [
    'Job Card No', 'PO No', 'SKU', 'Material', 'Colour', 'Application',
    'Order Qty', 'Total Impressions Required', 'Machine', 'Department',
    'Status', 'Workflow Status', 'Destination', 'Remarks',
    'Balance Qty', 'Dispatch %', 'Job Status', 'Updated At', 'Deleted',
]

PLANNING_HEADERS = [
    'JC Number', 'PO No', 'SKU', 'Job Name', 'Material', 'Application',
    'Order Qty', 'Delivery Date', 'Machine', 'Department', 'Status',
    'Priority', 'Remarks', 'Updated At', 'Deleted',
]

PRODUCTION_HEADERS = [
    'Job Card No', 'Date', 'Shift', 'Entry Type', 'Machine', 'Output Sheets',
    'Waste Sheets', 'Impressions', 'Packing Qty', 'Downtime Minutes',
    'Status', 'Remarks', 'OEE %', 'Availability %', 'Performance %', 'Quality %',
    'Updated At', 'Deleted',
]

PRODUCTION_DOWNTIME_HEADERS = [
    'Job Card No', 'Production Date', 'Shift', 'Machine', 'Category',
    'Minutes', 'Note', 'Updated At', 'Deleted',
]

DISPATCH_HEADERS = [
    'Job Card No', 'DC No', 'Dispatch Date', 'Dispatch Qty', 'PO No', 'SKU',
    'Created By', 'Active', 'Updated At', 'Deleted',
]

MAINTENANCE_RECORD_HEADERS = [
    'Record No', 'Machine', 'Reported Date', 'Reported By', 'Maintenance Type',
    'Priority', 'Fault Description', 'Status', 'Assigned To', 'Work Start',
    'Work End', 'Remarks', 'Updated At', 'Deleted',
]

MAINTENANCE_SPARE_PART_HEADERS = [
    'Record No', 'Machine', 'Description', 'Quantity', 'UOM', 'Existing SKU',
    'Item Request', 'Updated At', 'Deleted',
]

MAINTENANCE_SERVICE_JOB_HEADERS = [
    'Record No', 'Machine', 'Vendor', 'Scope', 'Item Request', 'Sent Out Date',
    'Returned Date', 'Updated At', 'Deleted',
]

MACHINE_DOWNTIME_HEADERS = [
    'Machine', 'Record No', 'Start At', 'End At', 'Reason',
    'Scheduled Minutes Lost', 'Updated At', 'Deleted',
]

SUPPLY_CHAIN_ITEM_HEADERS = [
    'Item ID', 'Material', 'UOM', 'Sheet Packing/Pcs', 'Unit Cost',
    'Safety Stock', 'Max Stock Level', 'Lead Time Days', 'Updated At', 'Deleted',
]

STOCK_TRANSACTION_HEADERS = [
    'Job Card No', 'SKU', 'Transaction Type', 'Source', 'Date', 'Month',
    'GIN/JC', 'Sheet Qty/Pcs', 'Pkt/Rim Qty', 'Active', 'Approved',
    'Updated At', 'Deleted',
]

STOCK_DEMAND_HEADERS = [
    'SKU', 'Month', 'Sheet Qty/Pcs', 'Pkt/Rim Qty', 'Active',
    'Updated At', 'Deleted',
]

ITEM_REQUEST_HEADERS = [
    'IR-ID', 'Request Type', 'Request Date', 'Item Title', 'Machine',
    'UOM', 'Required Quantity', 'Department', 'Status', 'Raised By',
    'Local/Import', 'Estimated Unit Price', 'Updated At', 'Deleted',
]

RAW_MATERIAL_SKU_HEADERS = [
    'SKU', 'Material', 'Purchase Sheet Size', 'SKU Type', 'UOM',
    'Sheet Packing/Pcs', 'Unit Cost', 'Safety Stock', 'Max Stock Level',
    'Lead Time Days', 'Active', 'Updated At', 'Deleted',
]

MACHINE_HEADERS = [
    'Machine Name', 'Machine Type', 'Std Impressions/Hour',
    'Std Setup Min/Color', 'Plate Life Impressions', 'Default Colors',
    'Min Print Length (mm)', 'Min Print Width (mm)', 'Max Print Length (mm)',
    'Max Print Width (mm)', 'Active', 'Updated At', 'Deleted',
]

MATERIAL_HEADERS = [
    'Name', 'Updated At', 'Deleted',
]

DEPARTMENT_HEADERS = [
    'Name', 'Updated At', 'Deleted',
]

PRODUCT_TYPE_HEADERS = [
    'Name', 'Updated At', 'Deleted',
]

SKU_RECIPE_HEADERS = [
    'SKU', 'Job Name', 'Material', 'Color Spec', 'Application',
    'Product Type', 'Machine', 'Job Process Type', 'Print Passes',
    'Size W (mm)', 'Size H (mm)', 'UPS', 'Print Sheet Size', 'Active',
    'Notes', 'Updated At', 'Deleted',
]

SYNCED_MODELS = [
    {
        'dotted_path': 'core.JobCard',
        'tab_name': 'Job Cards',
        'headers': JOB_CARD_HEADERS,
        'serializer': 'serialize_job_card',
        'key_field': 'job_card_no',
    },
    {
        'dotted_path': 'planning.PlanningJob',
        'tab_name': 'Planning',
        'headers': PLANNING_HEADERS,
        'serializer': 'serialize_planning_job',
        'key_field': 'jc_number',
    },
    {
        'dotted_path': 'core.Production',
        'tab_name': 'Production',
        'headers': PRODUCTION_HEADERS,
        'serializer': 'serialize_production',
        'key_field': 'id',
    },
    {
        'dotted_path': 'core.ProductionDowntime',
        'tab_name': 'Production Downtime',
        'headers': PRODUCTION_DOWNTIME_HEADERS,
        'serializer': 'serialize_production_downtime',
        'key_field': 'id',
    },
    {
        'dotted_path': 'core.Dispatch',
        'tab_name': 'Dispatch',
        'headers': DISPATCH_HEADERS,
        'serializer': 'serialize_dispatch',
        'key_field': 'id',
    },
    {
        'dotted_path': 'maintenance.MaintenanceRecord',
        'tab_name': 'Maintenance Records',
        'headers': MAINTENANCE_RECORD_HEADERS,
        'serializer': 'serialize_maintenance_record',
        'key_field': 'id',
    },
    {
        'dotted_path': 'maintenance.MaintenanceSparePart',
        'tab_name': 'Maintenance Spare Parts',
        'headers': MAINTENANCE_SPARE_PART_HEADERS,
        'serializer': 'serialize_maintenance_spare_part',
        'key_field': 'id',
    },
    {
        'dotted_path': 'maintenance.MaintenanceServiceJob',
        'tab_name': 'Maintenance Service Jobs',
        'headers': MAINTENANCE_SERVICE_JOB_HEADERS,
        'serializer': 'serialize_maintenance_service_job',
        'key_field': 'id',
    },
    {
        'dotted_path': 'maintenance.MachineDowntime',
        'tab_name': 'Machine Downtime',
        'headers': MACHINE_DOWNTIME_HEADERS,
        'serializer': 'serialize_machine_downtime',
        'key_field': 'id',
    },
    {
        'dotted_path': 'supply_chain.SupplyChainItem',
        'tab_name': 'Supply Chain Items',
        'headers': SUPPLY_CHAIN_ITEM_HEADERS,
        'serializer': 'serialize_supply_chain_item',
        'key_field': 'item_id',
    },
    {
        'dotted_path': 'supply_chain.StockTransaction',
        'tab_name': 'Stock Transactions',
        'headers': STOCK_TRANSACTION_HEADERS,
        'serializer': 'serialize_stock_transaction',
        'key_field': 'id',
    },
    {
        'dotted_path': 'supply_chain.StockDemand',
        'tab_name': 'Stock Demand',
        'headers': STOCK_DEMAND_HEADERS,
        'serializer': 'serialize_stock_demand',
        'key_field': 'id',
    },
    {
        'dotted_path': 'supply_chain.ItemRequest',
        'tab_name': 'Item Requests',
        'headers': ITEM_REQUEST_HEADERS,
        'serializer': 'serialize_item_request',
        'key_field': 'request_no',
    },
    {
        'dotted_path': 'supply_chain.RawMaterialSku',
        'tab_name': 'Raw Material SKUs',
        'headers': RAW_MATERIAL_SKU_HEADERS,
        'serializer': 'serialize_raw_material_sku',
        'key_field': 'sku',
    },
    {
        'dotted_path': 'core.Machine',
        'tab_name': 'Machines',
        'headers': MACHINE_HEADERS,
        'serializer': 'serialize_machine',
        'key_field': 'name',
    },
    {
        'dotted_path': 'core.Material',
        'tab_name': 'Materials',
        'headers': MATERIAL_HEADERS,
        'serializer': 'serialize_material',
        'key_field': 'name',
    },
    {
        'dotted_path': 'core.Department',
        'tab_name': 'Departments',
        'headers': DEPARTMENT_HEADERS,
        'serializer': 'serialize_department',
        'key_field': 'name',
    },
    {
        'dotted_path': 'core.ProductType',
        'tab_name': 'Product Types',
        'headers': PRODUCT_TYPE_HEADERS,
        'serializer': 'serialize_product_type',
        'key_field': 'name',
    },
    {
        'dotted_path': 'planning.SkuRecipe',
        'tab_name': 'SKU Master',
        'headers': SKU_RECIPE_HEADERS,
        'serializer': 'serialize_sku_recipe',
        'key_field': 'sku',
    },
]


def get_registry_entry(dotted_path):
    for entry in SYNCED_MODELS:
        if entry['dotted_path'] == dotted_path:
            return entry
    return None
