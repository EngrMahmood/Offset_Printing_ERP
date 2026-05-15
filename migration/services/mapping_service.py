from migration.models import ColumnMapping


def save_mapping_profile(comparison_job, profile_name, mappings, user=None):
    if not profile_name:
        profile_name = f'profile_{comparison_job.pk}'

    ColumnMapping.objects.filter(comparison_job=comparison_job, profile_name=profile_name).delete()

    mapping_objects = []
    for sheet_column, mapping in mappings.items():
        mapping_objects.append(
            ColumnMapping(
                comparison_job=comparison_job,
                profile_name=profile_name,
                sheet_column=sheet_column,
                erp_model=mapping.get('erp_model', ''),
                erp_field=mapping.get('erp_field', ''),
                match_confidence=mapping.get('confidence', 0.0) or 0.0,
                is_confirmed=mapping.get('confirmed', False),
                created_by=user,
            )
        )

    ColumnMapping.objects.bulk_create(mapping_objects)
    return mapping_objects


def get_mapping_profile(comparison_job, profile_name):
    return ColumnMapping.objects.filter(comparison_job=comparison_job, profile_name=profile_name)
