import os
import sys
import django

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'Offset_ERP.settings')
django.setup()

from core.models import JobCard, Production
from supply_chain.models import StockTransaction
from supply_chain.jc_sync import get_raw_material_sku_for_job_card

def analyze_missing():
    print("Analyzing Job Cards with no stock transactions...")
    all_jcs = JobCard.objects.filter(is_active=True).select_related('material', 'planning_job')
    
    total_jcs = all_jcs.count()
    no_sku = 0
    has_issuances = 0
    no_issuances = []
    
    for jc in all_jcs:
        # Check if it has any synced stock transactions
        txns_count = StockTransaction.objects.filter(job_card=jc, is_active=True).count()
        if txns_count > 0:
            has_issuances += 1
            continue
            
        raw_sku = get_raw_material_sku_for_job_card(jc)
        if not raw_sku:
            no_sku += 1
            continue
            
        # Get production logs
        prods = jc.productions.filter(is_active=True)
        prod_count = prods.count()
        total_prod_sheets = sum(int(p.output_sheets or 0) + int(p.waste_sheets or 0) for p in prods)
        
        no_issuances.append({
            'jc': jc,
            'prod_count': prod_count,
            'total_prod_sheets': total_prod_sheets,
            'planned_sheets': jc.total_sheet_quantity or 0,
        })
        
    print(f"\n--- Job Card Summary ---")
    print(f"Total Active Job Cards:       {total_jcs}")
    print(f"Job Cards with Issuances:     {has_issuances}")
    print(f"Job Cards without SKU mapping: {no_sku}")
    print(f"Job Cards with No Issuances:   {len(no_issuances)}")
    
    if no_issuances:
        print(f"\n--- Sample Job Cards with No Issuances ---")
        count = 0
        for item in no_issuances:
            jc = item['jc']
            print(f"- {jc.job_card_no} | Planned Sheets: {item['planned_sheets']} | Prod Logs: {item['prod_count']} | Prod Sheets: {item['total_prod_sheets']}")
            count += 1
            if count >= 20:
                print("... and more")
                break

if __name__ == '__main__':
    analyze_missing()
