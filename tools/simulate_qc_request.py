import os
import sys
sys.path.insert(0, os.getcwd())
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'Offset_ERP.settings')
import django
django.setup()
from django.test import Client
from django.contrib.auth import get_user_model
User = get_user_model()
user = User.objects.filter(username__iexact='QC').first()
if not user:
    print('QC user not found')
    sys.exit(1)
client = Client()
client.force_login(user)
resp = client.get('/planning/pending-skus/')
print('status_code=', resp.status_code)
print('is_redirect=', resp.status_code in (301,302))
# print small part of content for debugging
content = resp.content.decode('utf-8')
print('contains Pending SKUs header?', 'Pending SKUs' in content)
# Check pending_count displayed
import re
m = re.search(r'Pending: (\d+)', content)
print('pending_count matched:', m.group(1) if m else 'no-match')
# Check if "No pending SKUs" text present
print('no pending message present?', 'No pending SKUs for the selected filter' in content)

for path in ['/planning/', '/planning/pending-skus/', '/planning/sku-recipes/?status=pending_review', '/planning/approval-queue/']:
    resp = client.get(path)
    print(f'\nGET {path} =>', resp.status_code, 'redirect' if resp.status_code in (301,302) else 'ok')
    if resp.status_code == 200:
        snippet = resp.content.decode('utf-8')[:400]
        print('snippet:', snippet.replace('\n',' ')[:200])
