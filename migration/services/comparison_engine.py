from django.db import transaction

from .google_reader import read_google_sheet_metadata
from .erp_schema_reader import get_erp_schema_for_module
from .field_matcher import match_sheet_columns, normalize_name
from migration.models import ComparisonJob, ComparisonResult, ComparisonStatus


def _collect_required_fields(erp_schema):
    required = []
    for schema in erp_schema:
        for field in schema['fields']:
            if field['required']:
                required.append((schema['model_label'], field['name'], normalize_name(field['name'])))
    return required


def _build_column_samples(columns, sample_rows):
    samples = {}
    for column in columns:
        values = []
        for row in sample_rows:
            if column in row:
                values.append(row.get(column))
            else:
                normalized = normalize_name(column)
                for key in row.keys():
                    if normalize_name(key) == normalized:
                        values.append(row.get(key))
                        break
        samples[column] = values[:3]
    return samples


def compare_sheet_to_erp(sheet_url, module_name, oauth_token=None, user=None):
    metadata = read_google_sheet_metadata(sheet_url, oauth_token=oauth_token)
    sheet_columns = metadata.get('columns', [])
    sample_rows = metadata.get('sample_rows', [])
    if not sheet_columns:
        raise RuntimeError('Could not identify any sheet columns from the provided Google Sheet URL.')
    erp_schema = get_erp_schema_for_module(module_name)
    match_results = match_sheet_columns(sheet_columns, erp_schema)

    required_fields = _collect_required_fields(erp_schema)
    matched_field_keys = {
        (result['erp_model'], result['erp_field'])
        for result in match_results
        if result['erp_model'] and result['erp_field'] and result['status'] == 'MATCHED'
    }

    missing_fields = [
        {'erp_model': model_label, 'erp_field': field_name}
        for model_label, field_name, normalized in required_fields
        if (model_label, field_name) not in matched_field_keys
    ]

    total_columns = len(sheet_columns)
    matched_columns = sum(1 for result in match_results if result['status'] == 'MATCHED')
    extra_columns = sum(1 for result in match_results if result['status'] == 'NOT_FOUND')

    status = ComparisonStatus.COMPLETED if not missing_fields else ComparisonStatus.REVIEW

    with transaction.atomic():
        comparison_job = ComparisonJob.objects.create(
            module=module_name,
            sheet_url=sheet_url,
            status=status,
            total_columns=total_columns,
            matched_columns=matched_columns,
            missing_columns=len(missing_fields),
            extra_columns=extra_columns,
            created_by=user,
        )

        sample_values_by_column = _build_column_samples(sheet_columns, sample_rows)

        result_objs = []
        for result in match_results:
            result_objs.append(
                ComparisonResult(
                    comparison_job=comparison_job,
                    sheet_column=result['sheet_column'],
                    erp_model=result['erp_model'],
                    erp_field=result['erp_field'],
                    match_type=result['match_type'],
                    status=result['status'],
                    confidence=result['confidence'],
                    details='',
                    sample_values=sample_values_by_column.get(result['sheet_column'], []),
                )
            )
        ComparisonResult.objects.bulk_create(result_objs)

    return comparison_job, missing_fields, match_results, erp_schema, metadata
