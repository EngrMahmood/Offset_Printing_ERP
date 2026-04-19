import os
import django

# Updated project settings path
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'Offset_ERP.settings')
django.setup()

import openpyxl
from django.core.files.uploadedfile import SimpleUploadedFile
from django.test import RequestFactory
from django.contrib.messages.storage.fallback import FallbackStorage
from django.contrib.auth.models import User
from planning.views import sku_recipe_bulk_upload

file_path = 'planning/docs/sku_recipe_upload_sample_2.xlsx'

with open(file_path, 'rb') as f:
    upload_file = SimpleUploadedFile('sku_recipe_upload_sample_2.xlsx', f.read(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

factory = RequestFactory()
request = factory.post('/planning/sku-recipes/bulk-upload/', {'upload_file': upload_file})

# Mock user and messages
user = User.objects.first()
if not user:
    user = User.objects.create_superuser('admin', 'admin@example.com', 'password')
request.user = user

# Required for messages to work in the view
setattr(request, '_messages', FallbackStorage(request))

print('Calling sku_recipe_bulk_upload...')
try:
    response = sku_recipe_bulk_upload(request)
    print(f'Response Status Code: {response.status_code}')
    
    from django.contrib import messages
    storage = messages.get_messages(request)
    msgs = [str(m) for m in storage]
    print(f'Messages generated: {msgs}')
    
    if response.status_code == 302:
        print(f'Redirected to: {response.url}')
except Exception as e:
    import traceback
    print(f'Error occurred: {e}')
    traceback.print_exc()
