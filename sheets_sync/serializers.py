"""One row-builder function per synced tab.

Each function takes a model instance and returns an ordered dict matching
that tab's header row (see sheets_sync.registry). Kept explicit (not a
generic field dump) so FK fields render as human labels and internal-only
fields never leak into the DR spreadsheet.
"""


def _s(value):
    if value is None:
        return ''
    return str(value)


def _safe_attr(instance, attr_name):
    """Read a property/method off `instance`, tolerating any exception it
    might raise (e.g. edge-case division or missing related rows) so one
    bad computed field never breaks the whole sync row."""
    try:
        value = getattr(instance, attr_name)
        if callable(value):
            value = value()
        return _s(value)
    except Exception:
        return ''


def serialize_job_card(instance, deleted=False):
    return {
        'Job Card No': _s(instance.job_card_no),
        'PO No': _s(instance.PO_No),
        'SKU': _s(instance.SKU),
        'Material': _s(instance.material),
        'Colour': _s(instance.colour),
        'Application': _s(instance.application),
        'Order Qty': _s(instance.order_qty),
        'Total Impressions Required': _s(instance.total_impressions_required),
        'Machine': _s(instance.machine_name),
        'Department': _s(instance.department),
        'Status': _s(instance.status),
        'Workflow Status': _s(instance.workflow_status_label),
        'Destination': _s(instance.destination),
        'Remarks': _s(instance.remarks),
        'Balance Qty': _safe_attr(instance, 'balance_qty'),
        'Dispatch %': _safe_attr(instance, 'dispatch_completion_percent'),
        'Job Status': _safe_attr(instance, 'job_status'),
        'Updated At': _s(instance.updated_at) if hasattr(instance, 'updated_at') else '',
        'Deleted': 'YES' if deleted else '',
    }


def serialize_planning_job(instance, deleted=False):
    return {
        'JC Number': _s(instance.jc_number),
        'PO No': _s(instance.po_number),
        'SKU': _s(instance.sku),
        'Job Name': _s(instance.job_name),
        'Material': _s(instance.material),
        'Application': _s(instance.application),
        'Order Qty': _s(instance.order_qty),
        'Delivery Date': _s(instance.delivery_date),
        'Machine': _s(instance.machine_name),
        'Department': _s(instance.department),
        'Status': _s(instance.get_status_display()) if instance.status else '',
        'Priority': _s(instance.get_priority_display()),
        'Remarks': _s(instance.remarks),
        'Updated At': _s(instance.updated_at) if hasattr(instance, 'updated_at') else '',
        'Deleted': 'YES' if deleted else '',
    }


def _safe_percent(instance, attr_name):
    """Read a 0..1 OEE-style ratio property and render it as a rounded
    percentage, tolerating any exception the property might raise."""
    try:
        ratio = getattr(instance, attr_name)
        return _s(round(float(ratio or 0) * 100, 2))
    except Exception:
        return ''


def serialize_production(instance, deleted=False):
    return {
        'Job Card No': _s(instance.job_card.job_card_no) if instance.job_card_id else '',
        'Date': _s(instance.date),
        'Shift': _s(instance.get_shift_display()) if instance.shift else '',
        'Entry Type': _s(instance.get_entry_type_display()) if instance.entry_type else '',
        'Machine': _s(instance.machine),
        'Output Sheets': _s(instance.output_sheets),
        'Waste Sheets': _s(instance.waste_sheets),
        'Impressions': _s(instance.impressions),
        'Packing Qty': _s(instance.packing_qty),
        'Downtime Minutes': _s(instance.downtime_minutes),
        'Status': _s(instance.get_status_display()) if instance.status else '',
        'Remarks': _s(instance.remark_notes),
        'OEE %': _safe_percent(instance, 'oee'),
        'Availability %': _safe_percent(instance, 'availability'),
        'Performance %': _safe_percent(instance, 'performance'),
        'Quality %': _safe_percent(instance, 'quality'),
        'Updated At': _s(instance.updated_at) if hasattr(instance, 'updated_at') else '',
        'Deleted': 'YES' if deleted else '',
    }


def serialize_production_downtime(instance, deleted=False):
    production = instance.production
    return {
        'Job Card No': _s(production.job_card.job_card_no) if production and production.job_card_id else '',
        'Production Date': _s(production.date) if production else '',
        'Shift': _s(production.get_shift_display()) if production and production.shift else '',
        'Machine': _s(production.machine) if production else '',
        'Category': _s(instance.get_category_display()) if instance.category else '',
        'Minutes': _s(instance.minutes),
        'Note': _s(instance.note),
        'Updated At': _s(instance.updated_at) if hasattr(instance, 'updated_at') else '',
        'Deleted': 'YES' if deleted else '',
    }


def serialize_dispatch(instance, deleted=False):
    job_card = instance.job_card
    return {
        'Job Card No': _s(job_card.job_card_no) if instance.job_card_id else '',
        'DC No': _s(instance.dc_no),
        'Dispatch Date': _s(instance.dispatch_date),
        'Dispatch Qty': _s(instance.dispatch_qty),
        'PO No': _s(job_card.PO_No) if instance.job_card_id else '',
        'SKU': _s(job_card.SKU) if instance.job_card_id else '',
        'Created By': _s(instance.created_by),
        'Active': 'YES' if instance.is_active else 'NO',
        'Updated At': _s(instance.updated_at) if hasattr(instance, 'updated_at') else '',
        'Deleted': 'YES' if deleted else '',
    }


def serialize_maintenance_record(instance, deleted=False):
    return {
        'Record No': _s(instance.record_no),
        'Machine': _s(instance.machine),
        'Reported Date': _s(instance.reported_date),
        'Reported By': _s(instance.reported_by),
        'Maintenance Type': _s(instance.get_maintenance_type_display()) if instance.maintenance_type else '',
        'Priority': _s(instance.get_priority_display()) if instance.priority else '',
        'Fault Description': _s(instance.fault_description),
        'Status': _s(instance.get_status_display()) if instance.status else '',
        'Assigned To': _s(instance.assigned_to),
        'Work Start': _s(instance.work_start_at),
        'Work End': _s(instance.work_end_at),
        'Remarks': _s(instance.remarks),
        'Updated At': _s(instance.updated_at) if hasattr(instance, 'updated_at') else '',
        'Deleted': 'YES' if deleted else '',
    }


def serialize_maintenance_spare_part(instance, deleted=False):
    record = instance.record
    return {
        'Record No': _s(record.record_no) if instance.record_id else '',
        'Machine': _s(record.machine) if instance.record_id else '',
        'Description': _s(instance.description),
        'Quantity': _s(instance.quantity),
        'UOM': _s(instance.uom),
        'Existing SKU': _s(instance.existing_sku),
        'Item Request': _s(instance.item_request),
        'Updated At': _s(instance.updated_at) if hasattr(instance, 'updated_at') else '',
        'Deleted': 'YES' if deleted else '',
    }


def serialize_maintenance_service_job(instance, deleted=False):
    record = instance.record
    return {
        'Record No': _s(record.record_no) if instance.record_id else '',
        'Machine': _s(record.machine) if instance.record_id else '',
        'Vendor': _s(instance.vendor),
        'Scope': _s(instance.scope),
        'Item Request': _s(instance.item_request),
        'Sent Out Date': _s(instance.sent_out_date),
        'Returned Date': _s(instance.returned_date),
        'Updated At': _s(instance.updated_at) if hasattr(instance, 'updated_at') else '',
        'Deleted': 'YES' if deleted else '',
    }


def serialize_machine_downtime(instance, deleted=False):
    return {
        'Machine': _s(instance.machine),
        'Record No': _s(instance.record.record_no) if instance.record_id else '',
        'Start At': _s(instance.start_at),
        'End At': _s(instance.end_at),
        'Reason': _s(instance.get_reason_display()) if instance.reason else '',
        'Scheduled Minutes Lost': _s(instance.scheduled_minutes_lost),
        'Updated At': _s(instance.updated_at) if hasattr(instance, 'updated_at') else '',
        'Deleted': 'YES' if deleted else '',
    }


def serialize_supply_chain_item(instance, deleted=False):
    return {
        'Item ID': _s(instance.item_id),
        'Material': _s(instance.material),
        'UOM': _s(instance.uom),
        'Sheet Packing/Pcs': _s(instance.sheet_packing_pcs),
        'Unit Cost': _s(instance.unit_cost),
        'Safety Stock': _s(instance.safety_stock),
        'Max Stock Level': _s(instance.max_stock_level),
        'Lead Time Days': _s(instance.lead_time_days),
        'Updated At': _s(instance.updated_at) if hasattr(instance, 'updated_at') else '',
        'Deleted': 'YES' if deleted else '',
    }


def serialize_stock_transaction(instance, deleted=False):
    if instance.raw_material_sku_id:
        sku = instance.raw_material_sku.sku
    elif instance.item_id:
        sku = instance.item.item_id
    else:
        sku = ''
    return {
        'Job Card No': _s(instance.job_card.job_card_no) if instance.job_card_id else '',
        'SKU': _s(sku),
        'Transaction Type': _s(instance.get_transaction_type_display()) if instance.transaction_type else '',
        'Source': _s(instance.get_source_display()) if instance.source else '',
        'Date': _s(instance.date),
        'Month': _s(instance.month_str),
        'GIN/JC': _s(instance.gin_jc),
        'Sheet Qty/Pcs': _s(instance.sheet_qty_pcs),
        'Pkt/Rim Qty': _s(instance.pkt_rim_qty),
        'Active': 'YES' if instance.is_active else 'NO',
        'Approved': 'YES' if instance.is_approved else 'NO',
        'Updated At': _s(instance.updated_at) if hasattr(instance, 'updated_at') else '',
        'Deleted': 'YES' if deleted else '',
    }


def serialize_stock_demand(instance, deleted=False):
    if instance.raw_material_sku_id:
        sku = instance.raw_material_sku.sku
    elif instance.item_id:
        sku = instance.item.item_id
    else:
        sku = ''
    return {
        'SKU': _s(sku),
        'Month': _s(instance.month_str),
        'Sheet Qty/Pcs': _s(instance.sheet_qty_pcs),
        'Pkt/Rim Qty': _s(instance.pkt_rim_qty),
        'Active': 'YES' if instance.is_active else 'NO',
        'Updated At': _s(instance.updated_at) if hasattr(instance, 'updated_at') else '',
        'Deleted': 'YES' if deleted else '',
    }


def serialize_item_request(instance, deleted=False):
    return {
        'IR-ID': _s(instance.request_no),
        'Request Type': _s(instance.request_type),
        'Request Date': _s(instance.request_date),
        'Item Title': _s(instance.item_title),
        'Machine': _s(instance.machine_display),
        'UOM': _s(instance.uom),
        'Required Quantity': _s(instance.required_quantity),
        'Department': _s(instance.department),
        'Status': _s(instance.get_status_display()) if instance.status else '',
        'Raised By': _s(instance.raised_by),
        'Local/Import': _s(instance.get_local_import_display()) if instance.local_import else '',
        'Estimated Unit Price': _s(instance.estimated_unit_price),
        'Updated At': _s(instance.updated_at) if hasattr(instance, 'updated_at') else '',
        'Deleted': 'YES' if deleted else '',
    }


def serialize_raw_material_sku(instance, deleted=False):
    return {
        'SKU': _s(instance.sku),
        'Material': _s(instance.material),
        'Purchase Sheet Size': _s(instance.purchase_sheet_size),
        'SKU Type': _s(instance.sku_type) if instance.sku_type_id else '',
        'UOM': _s(instance.uom),
        'Sheet Packing/Pcs': _s(instance.sheet_packing_pcs),
        'Unit Cost': _s(instance.unit_cost),
        'Safety Stock': _s(instance.safety_stock),
        'Max Stock Level': _s(instance.max_stock_level),
        'Lead Time Days': _s(instance.lead_time_days),
        'Active': 'YES' if instance.is_active else 'NO',
        'Updated At': _s(instance.updated_at) if hasattr(instance, 'updated_at') else '',
        'Deleted': 'YES' if deleted else '',
    }


def serialize_machine(instance, deleted=False):
    return {
        'Machine Name': _s(instance.name),
        'Machine Type': _s(instance.get_machine_type_display()) if instance.machine_type else '',
        'Std Impressions/Hour': _s(instance.standard_impressions_per_hour),
        'Std Setup Min/Color': _s(instance.standard_setup_minutes_per_color),
        'Plate Life Impressions': _s(instance.plate_life_impressions),
        'Default Colors': _s(instance.default_colors),
        'Min Print Length (mm)': _s(instance.min_print_length_mm),
        'Min Print Width (mm)': _s(instance.min_print_width_mm),
        'Max Print Length (mm)': _s(instance.max_print_length_mm),
        'Max Print Width (mm)': _s(instance.max_print_width_mm),
        'Active': 'YES' if instance.is_active else 'NO',
        'Updated At': _s(instance.updated_at) if hasattr(instance, 'updated_at') else '',
        'Deleted': 'YES' if deleted else '',
    }


def serialize_material(instance, deleted=False):
    return {
        'Name': _s(instance.name),
        'Updated At': _s(instance.updated_at) if hasattr(instance, 'updated_at') else '',
        'Deleted': 'YES' if deleted else '',
    }


def serialize_department(instance, deleted=False):
    return {
        'Name': _s(instance.name),
        'Updated At': _s(instance.updated_at) if hasattr(instance, 'updated_at') else '',
        'Deleted': 'YES' if deleted else '',
    }


def serialize_product_type(instance, deleted=False):
    return {
        'Name': _s(instance.name),
        'Updated At': _s(instance.updated_at) if hasattr(instance, 'updated_at') else '',
        'Deleted': 'YES' if deleted else '',
    }


def serialize_sku_recipe(instance, deleted=False):
    return {
        'SKU': _s(instance.sku),
        'Job Name': _s(instance.job_name),
        'Material': _s(instance.material),
        'Color Spec': _s(instance.color_spec),
        'Application': _s(instance.application),
        'Product Type': _s(instance.product_type),
        'Machine': _s(instance.machine_name),
        'Job Process Type': _s(instance.get_job_process_type_display()) if instance.job_process_type else '',
        'Print Passes': _s(instance.print_passes),
        'Size W (mm)': _s(instance.size_w_mm),
        'Size H (mm)': _s(instance.size_h_mm),
        'UPS': _s(instance.ups),
        'Print Sheet Size': _s(instance.print_sheet_size),
        'Active': 'YES' if instance.is_active else 'NO',
        'Notes': _s(instance.notes),
        'Updated At': _s(instance.updated_at) if hasattr(instance, 'updated_at') else '',
        'Deleted': 'YES' if deleted else '',
    }


SERIALIZERS = {
    'serialize_job_card': serialize_job_card,
    'serialize_planning_job': serialize_planning_job,
    'serialize_production': serialize_production,
    'serialize_production_downtime': serialize_production_downtime,
    'serialize_dispatch': serialize_dispatch,
    'serialize_maintenance_record': serialize_maintenance_record,
    'serialize_maintenance_spare_part': serialize_maintenance_spare_part,
    'serialize_maintenance_service_job': serialize_maintenance_service_job,
    'serialize_machine_downtime': serialize_machine_downtime,
    'serialize_supply_chain_item': serialize_supply_chain_item,
    'serialize_stock_transaction': serialize_stock_transaction,
    'serialize_stock_demand': serialize_stock_demand,
    'serialize_item_request': serialize_item_request,
    'serialize_raw_material_sku': serialize_raw_material_sku,
    'serialize_machine': serialize_machine,
    'serialize_material': serialize_material,
    'serialize_department': serialize_department,
    'serialize_product_type': serialize_product_type,
    'serialize_sku_recipe': serialize_sku_recipe,
}
