def validate_jobcard_row(row, models):
    errors = []

    # Required fields
    if not row.get('job_card_no'):
        errors.append("Job Card No missing")

    if not row.get('SKU'):
        errors.append("SKU missing")

    # Order qty
    try:
        if int(row.get('order_qty', 0)) <= 0:
            errors.append("Order Qty must be > 0")
    except:
        errors.append("Order Qty must be numeric")

    # Master checks
    Material = models['Material']
    Machine = models['Machine']
    Department = models['Department']

    if row.get('material') and not Material.objects.filter(name=row['material']).exists():
        errors.append(f"Material not found: {row['material']}")

    if row.get('machine_name') and not Machine.objects.filter(name=row['machine_name']).exists():
        errors.append(f"Machine not found: {row['machine_name']}")

    if row.get('department') and not Department.objects.filter(name=row['department']).exists():
        errors.append(f"Department not found: {row['department']}")

    # Duplicate job card check
    JobCard = models['JobCard']
    if JobCard.objects.filter(job_card_no=row.get('job_card_no')).exists():
        errors.append(f"Duplicate Job Card: {row['job_card_no']}")

    return errors