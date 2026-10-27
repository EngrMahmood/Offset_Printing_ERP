"""Tell the shared Master Data screen where a Machine/Operator/Department is used by
Flexo, so deleting a shared master that Flexo relies on is refused (the offset checks
only know about offset records)."""
from .models import (
    FlexoJobCard, FlexoPlanningJob, FlexoPrintingLogEntry, FlexoSkuRecipe,
    FlexoSlittingCuttingLogEntry,
)


def linked_flexo_records(entity_type, record):
    """'Planning Jobs: 3, SKU Masters: 1' for a shared master record, '' when unused."""
    if entity_type == 'machine':
        usage = [
            ('SKU Masters', FlexoSkuRecipe.objects.filter(machine=record)),
            ('Planning Jobs', FlexoPlanningJob.objects.filter(machine_name=record)),
            ('Job Cards', FlexoJobCard.objects.filter(machine=record)),
            ('Printing Logs', FlexoPrintingLogEntry.objects.filter(machine=record)),
            ('Slitting/Cutting Logs', FlexoSlittingCuttingLogEntry.objects.filter(machine=record)),
        ]
    elif entity_type == 'operator':
        usage = [
            ('Printing Logs', FlexoPrintingLogEntry.objects.filter(operator=record)),
            ('Slitting/Cutting Logs', FlexoSlittingCuttingLogEntry.objects.filter(operator=record)),
        ]
    elif entity_type == 'department':
        usage = [
            ('SKU Masters', FlexoSkuRecipe.objects.filter(department=record)),
            ('Planning Jobs', FlexoPlanningJob.objects.filter(department=record)),
            ('Job Cards', FlexoJobCard.objects.filter(department=record)),
        ]
    else:
        return ''
    return ', '.join(f'{label}: {count}' for label, qs in usage if (count := qs.count()))
