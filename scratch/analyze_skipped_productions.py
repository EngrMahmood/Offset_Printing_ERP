import os
import sys
import django

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'Offset_ERP.settings')
django.setup()

from core.models import Production
from supply_chain.jc_sync import get_raw_material_sku_for_job_card

def analyze_skips():
    print("Analyzing active production records to find skipped ones...")
    productions = Production.objects.filter(is_active=True).select_related(
        'job_card', 'job_card__material', 'job_card__planning_job'
    ).order_by('id')
    
    total = productions.count()
    no_jc = 0
    no_sku = []
    zero_sheets = 0
    synced = 0
    
    for prod in productions:
        if not prod.job_card:
            no_jc += 1
            continue
            
        raw_sku = get_raw_material_sku_for_job_card(prod.job_card)
        if not raw_sku:
            no_sku.append(prod.job_card)
            continue
            
        consumed_sheets = int(prod.output_sheets or 0) + int(prod.waste_sheets or 0)
        if consumed_sheets <= 0:
            zero_sheets += 1
            continue
            
        synced += 1
        
    print(f"\n--- Analysis Summary ---")
    print(f"Total Active Productions: {total}")
    print(f"Synced Successfully:      {synced}")
    print(f"Skipped - No Job Card:    {no_jc}")
    print(f"Skipped - Zero Sheets:     {zero_sheets} (e.g., Packing/Sorting/Make-ready)")
    print(f"Skipped - No SKU Mapping: {len(no_sku)} (Job Cards missing material or size)")
    
    if no_sku:
        print(f"\n--- Sample Job Cards Missing SKU Mappings (Total {len(set(no_sku))}) ---")
        seen_jcs = set()
        count = 0
        for jc in no_sku:
            if jc.job_card_no in seen_jcs:
                continue
            seen_jcs.add(jc.job_card_no)
            
            material_name = jc.material.name if jc.material else "None"
            size = jc.purchase_sheet_size or (jc.planning_job.purchase_sheet_size_display if jc.planning_job else "None")
            print(f"- Job Card: {jc.job_card_no} | Material: {material_name} | Size: {size}")
            count += 1
            if count >= 15:
                print("... and more")
                break

if __name__ == '__main__':
    analyze_skips()
