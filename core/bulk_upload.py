import csv
from django.db import transaction
from .models import JobCard, Material,Machine, Department
from .validators import validate_jobcard_row


def process_jobcard_upload(file):
    decoded_file = file.read().decode('utf-8').splitlines()
    reader = csv.DictReader(decoded_file)

    models = {
        "JobCard": JobCard,
        "Material": Material,
        "Machine": Machine,
        "Department": Department
    }

    success = []
    errors = []

    # OPTIONAL: Full rollback mode (ERP SAFE MODE)
    with transaction.atomic():

        for i, row in enumerate(reader, start=1):

            row_errors = validate_jobcard_row(row, models)

            if row_errors:
                errors.append({
                    "row": i,
                    "job_card": row.get("job_card_no"),
                    "errors": row_errors
                })
                continue

            # SAFE INSERT
            JobCard.objects.create(
                job_card_no=row['job_card_no'],
                SKU=row['SKU'],
                order_qty=int(row['order_qty']),

                material=Material.objects.filter(name=row.get('material')).first() if row.get('material') else None,

                colour=int(row['colour']) if row.get('colour') else None,

                machine_name=Machine.objects.filter(name=row.get('machine_name')).first() if row.get('machine_name') else None,

                department=Department.objects.filter(name=row.get('department')).first() if row.get('department') else None,
            )

            success.append(row['job_card_no'])

    return {
        "success_count": len(success),
        "error_count": len(errors),
        "success": success,
        "errors": errors
    }