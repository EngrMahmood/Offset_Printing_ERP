import os
import sys
import django

# Set up Django environment
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'Offset_ERP.settings')
django.setup()

from django.core import serializers
from django.db import transaction
from supply_chain.models import RawMaterialSku, SupplyChainItem, StockTransaction, StockDemand, PhysicalStockCount, ChangeRequest

BACKUP_FILE = 'supply_chain_backup.json'

@transaction.atomic
def restore_data():
    if not os.path.exists(BACKUP_FILE):
        print(f"Error: Backup file {BACKUP_FILE} does not exist.")
        return
        
    print("Clearing current supply chain data before restore...")
    ChangeRequest.objects.filter(model_name__in=['RawMaterialSku', 'StockDemand', 'StockTransaction', 'PhysicalStockCount']).delete()
    PhysicalStockCount.objects.all().delete()
    StockDemand.objects.all().delete()
    StockTransaction.objects.all().delete()
    RawMaterialSku.objects.all().delete()
    SupplyChainItem.objects.all().delete()
    
    print(f"Restoring supply chain data from {BACKUP_FILE}...")
    with open(BACKUP_FILE, 'r', encoding='utf-8') as f:
        data = f.read()
        
    for obj in serializers.deserialize('json', data):
        obj.save()
        
    print("Restore complete and transaction committed successfully.")

if __name__ == '__main__':
    restore_data()
