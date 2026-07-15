import os
import sys
import django

# Set up Django environment
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'Offset_ERP.settings')
django.setup()

from django.core import serializers
from supply_chain.models import RawMaterialSku, SupplyChainItem, StockTransaction, StockDemand, PhysicalStockCount, ChangeRequest
from supply_chain.jc_sync import sync_all_job_card_issuances
from populate_master import populate_master

BACKUP_FILE = 'supply_chain_backup.json'

def backup_data():
    print(f"Backing up supply chain models to {BACKUP_FILE}...")
    models_to_backup = [RawMaterialSku, SupplyChainItem, StockTransaction, StockDemand, PhysicalStockCount, ChangeRequest]
    
    all_objects = []
    for model in models_to_backup:
        all_objects.extend(list(model.objects.all()))
        
    data = serializers.serialize('json', all_objects, indent=2)
    with open(BACKUP_FILE, 'w', encoding='utf-8') as f:
        f.write(data)
    print("Backup complete.")

def clear_data():
    print("Clearing all supply chain data...")
    # Delete dependent models first to avoid potential constraints (though CASCADE is defined)
    ChangeRequest.objects.filter(model_name__in=['RawMaterialSku', 'StockDemand', 'StockTransaction', 'PhysicalStockCount']).delete()
    PhysicalStockCount.objects.all().delete()
    StockDemand.objects.all().delete()
    StockTransaction.objects.all().delete()
    RawMaterialSku.objects.all().delete()
    SupplyChainItem.objects.all().delete()
    print("Supply chain data cleared.")

def resync_data():
    print("Resyncing data from planning and production...")
    # 1. Rebuild RawMaterialSku objects from Job Cards and Planning Jobs
    populate_master()
    
    # 2. Re-sync all issuances from Production logs (will create them as unapproved pending items in queue)
    synced, skipped = sync_all_job_card_issuances()
    print(f"Synced {synced} issuance transaction(s) in the queue. Skipped {skipped}.")

def main():
    backup_data()
    clear_data()
    resync_data()
    print("All steps completed successfully. Supply chain data is clean and freshly synced!")

if __name__ == '__main__':
    main()
