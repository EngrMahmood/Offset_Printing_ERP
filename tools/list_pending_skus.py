import os
import sys
sys.path.insert(0, os.getcwd())
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'Offset_ERP.settings')
import django
django.setup()
from planning.views import _collect_pending_sku_rows
from planning.models import PoDocument
po_docs = PoDocument.objects.exclude(extracted_payload__isnull=True).order_by('-created_at')[:400]
rows = _collect_pending_sku_rows(po_docs)
print('Pending rows:', len(rows))
for r in rows[:20]:
    print(r)
