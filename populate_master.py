import os
import django
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'Offset_ERP.settings')
django.setup()

from core.models import JobCard, Material
from supply_chain.models import RawMaterialSku, normalize_purchase_sheet_size


def populate_master():
    print("Populating master material data...")
    job_cards = JobCard.objects.filter(is_active=True).select_related('planning_job')

    unique_pairs = {}
    for jc in job_cards:
        if not jc.planning_job or not jc.planning_job.material:
            continue
        mat_name = jc.planning_job.material.strip()
        if not mat_name:
            continue
        purchase_size = normalize_purchase_sheet_size(
            jc.purchase_sheet_size
            or getattr(jc.planning_job, 'purchase_sheet_size_display', None)
            or jc.planning_job.purchase_sheet_size
            or ''
        )
        if not purchase_size:
            purchase_size = 'UNSPECIFIED'
        unique_pairs[(mat_name, purchase_size)] = True

    created_count = 0
    sku_created = 0
    updated_jcs = 0
    sku_index = 1

    for mat_name, purchase_size in sorted(unique_pairs.keys()):
        material, created = Material.objects.get_or_create(name=mat_name)
        if created:
            created_count += 1

        sku_code = f'MAT-{sku_index:04d}'
        sku_index += 1
        raw_sku, raw_created = RawMaterialSku.objects.get_or_create(
            material=material,
            purchase_sheet_size=purchase_size,
            defaults={
                'sku': sku_code,
                'uom': 'Sheets',
            },
        )
        if raw_created:
            sku_created += 1

        jcs_to_update = [
            jc for jc in job_cards
            if jc.planning_job
            and jc.planning_job.material
            and jc.planning_job.material.strip() == mat_name
        ]
        for jc in jcs_to_update:
            if jc.material != material:
                jc.material = material
                jc.save(update_fields=['material'])
                updated_jcs += 1

    print(f"Created {created_count} new Material masters.")
    print(f"Created {sku_created} new Raw Material SKUs.")
    print(f"Updated {updated_jcs} JobCards to link to the new Material masters.")


if __name__ == '__main__':
    populate_master()
