from django.apps import apps

from migration.models import ComparisonModule

ERP_MODULE_MODEL_MAP = {
    ComparisonModule.PO_INTAKE: [
        'planning.PoDocument',
        'planning.PlanningJob',
    ],
    ComparisonModule.SKU_MASTER: [
        'planning.Sku',
        'planning.PlanningJob',
    ],
    ComparisonModule.JOB_FINALIZE: [
        'planning.PlanningJob',
        'planning.SkuRecipe',
    ],
    ComparisonModule.PLANNING: [
        'planning.PlanningJob',
        'planning.SkuRecipe',
    ],
    ComparisonModule.PRODUCTION: [
        'planning.PlanningJob',
        'planning.SkuRecipe',
    ],
    ComparisonModule.DISPATCH: [
        'planning.PlanningJob',
        'planning.SkuRecipe',
    ],
}


def _resolve_model(model_path):
    app_label, model_name = model_path.split('.')
    return apps.get_model(app_label, model_name)


def _build_field_schema(field):
    return {
        'name': field.name,
        'type': field.get_internal_type() if hasattr(field, 'get_internal_type') else getattr(field, 'internal_type', 'Unknown'),
        'required': not getattr(field, 'blank', False) and not getattr(field, 'null', False) and not getattr(field, 'primary_key', False),
        'unique': getattr(field, 'unique', False),
        'help_text': getattr(field, 'help_text', ''),
        'is_relation': getattr(field, 'is_relation', False),
    }


def get_erp_schema_for_module(module_name):
    model_paths = ERP_MODULE_MODEL_MAP.get(module_name, [])
    schemas = []
    for model_path in model_paths:
        try:
            model = _resolve_model(model_path)
        except LookupError:
            continue

        fields = [
            _build_field_schema(field)
            for field in model._meta.get_fields()
            if getattr(field, 'concrete', False) and not getattr(field, 'auto_created', False)
        ]

        schemas.append(
            {
                'model_label': model._meta.label,
                'model_name': model._meta.object_name,
                'fields': fields,
            }
        )

    if not schemas:
        for model in apps.get_app_config('planning').get_models():
            fields = [
                _build_field_schema(field)
                for field in model._meta.get_fields()
                if getattr(field, 'concrete', False) and not getattr(field, 'auto_created', False)
            ]
            schemas.append(
                {
                    'model_label': model._meta.label,
                    'model_name': model._meta.object_name,
                    'fields': fields,
                }
            )

    return schemas
