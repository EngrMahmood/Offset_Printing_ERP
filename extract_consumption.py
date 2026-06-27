import os
import django
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'Offset_ERP.settings')
django.setup()

from django.db.models import Sum
from core.models import JobCard, Production
from collections import defaultdict

def extract_consumption():
    print("--- Material Consumption Summary ---")
    
    # Store aggregated data per material string
    material_data = defaultdict(lambda: {'planned': 0, 'produced': 0, 'waste': 0, 'jobs': 0})
    
    for jc in JobCard.objects.filter(is_active=True):
        if jc.planning_job and jc.planning_job.material:
            mat = jc.planning_job.material.strip()
        else:
            continue
            
        planned = jc.total_sheets_planned or 0
        
        prod_agg = Production.objects.filter(job_card=jc, is_active=True).aggregate(
            prod=Sum('output_sheets'),
            waste=Sum('waste_sheets')
        )
        
        produced = prod_agg['prod'] or 0
        waste = prod_agg['waste'] or 0
        
        material_data[mat]['jobs'] += 1
        material_data[mat]['planned'] += planned
        material_data[mat]['produced'] += produced
        material_data[mat]['waste'] += waste

    for mat, data in sorted(material_data.items()):
        total_consumed = data['produced'] + data['waste']
        print(f"Material: {mat}")
        print(f"  Total Jobs: {data['jobs']}")
        print(f"  Planned Sheets: {data['planned']}")
        print(f"  Actual Sheets Produced: {data['produced']}")
        print(f"  Actual Waste Sheets: {data['waste']}")
        print(f"  Total Consumed: {total_consumed}")
        print("-" * 30)

if __name__ == '__main__':
    extract_consumption()
