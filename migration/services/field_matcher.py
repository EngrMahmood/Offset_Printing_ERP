import difflib
import re

FIELD_ALIASES = {
    'jc': 'jc_number',
    'jc_no': 'jc_number',
    'job_card_no': 'jc_number',
    'jobcard_no': 'jc_number',
    'job_card': 'jc_number',
    'jobcard': 'jc_number',
    'job no': 'jc_number',
    'job number': 'jc_number',
    'month': 'plan_month',
    'date': 'plan_date',
    'po': 'po_number',
    'po_number': 'po_number',
    'po no': 'po_number',
    'po_no': 'po_number',
    'indent': 'pr_reference',
    'pr': 'pr_reference',
    'repeat': 'repeat_flag',
    'repeat_new': 'repeat_flag',
    'material': 'material',
    'color': 'color_spec',
    'color_spec': 'color_spec',
    'color spec': 'color_spec',
    'application': 'application',
    'die_cutting': 'die_cutting',
    'die cutting': 'die_cutting',
    'awc_no': 'awc_no',
    'awc no': 'awc_no',
    'plate_set_no': 'plate_set_no',
    'plate set no': 'plate_set_no',
    'p set no': 'plate_set_no',
    'purchase_material': 'purchase_material_origin',
    'purchase material origin': 'purchase_material_origin',
    'purchase material origin': 'purchase_material_origin',
    'po_received_date': 'plan_date',
    'po approval date': 'plan_date',
    'po_date': 'plan_date',
    'purchase_sheet_size': 'purchase_sheet_size',
    'purchase sheet size': 'purchase_sheet_size',
    'purchase_sheet_ups': 'purchase_sheet_ups',
    'purchase sheet ups': 'purchase_sheet_ups',
    'purchase_sheet_required': 'purchase_sheet_required',
    'purchase sheet require': 'purchase_sheet_required',
    'actual_sheet_required': 'actual_sheet_required',
    'actual sheet required': 'actual_sheet_required',
    'actual_sheet_require': 'actual_sheet_required',
    'actual sheet require': 'actual_sheet_required',
    'wastage_sheets': 'wastage_sheets',
    'wastage': 'wastage_sheets',
    'purchase_sheet_required': 'purchase_sheet_required',
    'purchase sheet require': 'purchase_sheet_required',
    'print_sheet_size': 'print_sheet_size',
    'print sheet size': 'print_sheet_size',
    'sheet': 'actual_sheet_required',
    'order_qty': 'order_qty',
    'order qty': 'order_qty',
    'qty': 'quantity',
    'ups': 'ups',
    'no_of_ups': 'ups',
    'print_pcs': 'print_pcs',
    'print pcs': 'print_pcs',
    'pkt': 'pkt_value',
    'pkt_value': 'pkt_value',
    'remarks': 'remarks',
    'requirement': 'requirement',
    "requirement'": 'requirement',
    'requirement_special_instructions': 'requirement',
    'destination': 'destination',
    'delivery_location': 'destination',
    'department': 'department',
    'cost': 'unit_cost',
    'status': 'status',
    'balance': 'balance_qty',
    'stock_qty': 'stock_qty',
    'stock': 'stock_qty',
    'machine_name': 'machine_name',
    'daily_demand': 'daily_demand',
    'purchase_material': 'purchase_material_origin',
}

FUZZY_THRESHOLD = 0.65


def normalize_name(value):
    if value is None:
        return ''

    normalized = re.sub(r'[^0-9a-zA-Z]+', '_', str(value).strip().lower())
    normalized = re.sub(r'_+', '_', normalized).strip('_')
    return normalized


def _alias_match(normalized_value):
    if normalized_value in FIELD_ALIASES:
        return normalize_name(FIELD_ALIASES[normalized_value])
    if normalized_value.replace('_', ' ') in FIELD_ALIASES:
        return normalize_name(FIELD_ALIASES[normalized_value.replace('_', ' ')])
    return None


def map_row_fields(raw_row):
    mapped = {}
    for key, value in raw_row.items():
        normalized_key = normalize_name(key)
        alias = _alias_match(normalized_key)
        mapped_key = alias if alias else normalized_key
        mapped[mapped_key] = value
    return mapped


def _build_field_candidates(erp_schema):
    candidates = []
    for schema in erp_schema:
        for field in schema['fields']:
            normalized = normalize_name(field['name'])
            candidates.append(
                {
                    'model_label': schema['model_label'],
                    'model_name': schema['model_name'],
                    'field_name': field['name'],
                    'normalized_name': normalized,
                    'field_type': field['type'],
                    'required': field['required'],
                }
            )
    return candidates


def _find_best_match(normalized_column, field_candidates):
    exact_matches = [candidate for candidate in field_candidates if candidate['normalized_name'] == normalized_column]
    if exact_matches:
        return exact_matches[0], 'EXACT', 1.0

    alias_normalized = _alias_match(normalized_column)
    if alias_normalized:
        alias_matches = [candidate for candidate in field_candidates if candidate['normalized_name'] == alias_normalized]
        if alias_matches:
            return alias_matches[0], 'ALIAS', 0.9

    field_names = [candidate['normalized_name'] for candidate in field_candidates]
    close_matches = difflib.get_close_matches(normalized_column, field_names, n=3, cutoff=FUZZY_THRESHOLD)
    if close_matches:
        best_match_name = close_matches[0]
        best_candidate = next(candidate for candidate in field_candidates if candidate['normalized_name'] == best_match_name)
        return best_candidate, 'FUZZY', 0.7

    return None, 'NONE', 0.0


def match_sheet_columns(sheet_columns, erp_schema):
    field_candidates = _build_field_candidates(erp_schema)
    matches = []
    seen_fields = set()
    normalized_columns = [normalize_name(column) for column in sheet_columns]

    for sheet_column, normalized_column in zip(sheet_columns, normalized_columns):
        candidate, match_type, confidence = _find_best_match(normalized_column, field_candidates)

        if candidate and (candidate['model_label'], candidate['field_name']) in seen_fields:
            match_type = 'DUPLICATE'
            confidence = 0.0
            status = 'DUPLICATE'
            result = {
                'sheet_column': sheet_column,
                'erp_model': candidate['model_label'],
                'erp_field': candidate['field_name'],
                'match_type': match_type,
                'status': status,
                'confidence': confidence,
            }
        elif candidate:
            status = 'MATCHED' if match_type in ('EXACT', 'ALIAS', 'FUZZY') else 'REVIEW'
            if match_type == 'NONE':
                status = 'NOT_FOUND'

            result = {
                'sheet_column': sheet_column,
                'erp_model': candidate['model_label'],
                'erp_field': candidate['field_name'],
                'match_type': match_type,
                'status': status,
                'confidence': confidence,
            }
            if status == 'MATCHED':
                seen_fields.add((candidate['model_label'], candidate['field_name']))
        else:
            result = {
                'sheet_column': sheet_column,
                'erp_model': '',
                'erp_field': '',
                'match_type': match_type,
                'status': 'NOT_FOUND',
                'confidence': confidence,
            }

        matches.append(result)

    return matches
