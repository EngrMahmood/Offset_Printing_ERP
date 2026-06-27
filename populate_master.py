import os
import django
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'Offset_ERP.settings')
django.setup()

from core.models import JobCard, Material
from supply_chain.models import SupplyChainItem

def populate_master():
    print("Populating master material data...")
    # Get distinct materials from planning_jobs
    job_cards = JobCard.objects.filter(is_active=True).select_related('planning_job')
    
    unique_materials = set()
    for jc in job_cards:
        if jc.planning_job and jc.planning_job.material:
            mat_name = jc.planning_job.material.strip()
            if mat_name:
                unique_materials.add(mat_name)
                
    created_count = 0
    sc_created = 0
    updated_jcs = 0
    
    for idx, mat_name in enumerate(sorted(unique_materials)):
        material, created = Material.objects.get_or_create(name=mat_name)
        if created:
            created_count += 1
            
        # Create a supply chain item for this material if it doesn't exist
        sc_item, sc_item_created = SupplyChainItem.objects.get_or_create(
            material=material,
            defaults={
                'item_id': f"MAT-{idx+1:04d}",
                'uom': 'Sheets',
            }
        )
        if sc_item_created:
            sc_created += 1
            
        # Link JobCards that have this material text to the actual Material master
        jcs_to_update = [jc for jc in job_cards if jc.planning_job and jc.planning_job.material and jc.planning_job.material.strip() == mat_name]
        for jc in jcs_to_update:
            if jc.material != material:
                jc.material = material
                jc.save(update_fields=['material'])
                updated_jcs += 1

    print(f"Created {created_count} new Material masters.")
    print(f"Created {sc_created} new Supply Chain items.")
    print(f"Updated {updated_jcs} JobCards to link to the new Material masters.")

if __name__ == '__main__':
    populate_master()
