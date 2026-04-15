import os
import sys
import django
from urllib.request import urlopen

# Setup Django
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'Offset_ERP.settings')
sys.path.insert(0, os.path.dirname(__file__))
django.setup()

# Test the download by making HTTP request
try:
    with urlopen('http://127.0.0.1:8000/download-template/?format=excel') as response:
        print('Status:', response.status)
        print('Content-Type:', response.headers.get('content-type'))
        print('Content-Disposition:', response.headers.get('content-disposition'))
        content = response.read()
        print('Content length:', len(content))
        
        if content.startswith(b'PK'):
            print('Content starts with PK - this is likely a ZIP/XLSX file')
            # Save to file to check
            with open('test_download.xlsx', 'wb') as f:
                f.write(content)
            print('Saved as test_download.xlsx')
        else:
            print('Content preview:', content[:200].decode('utf-8', errors='ignore'))
except Exception as e:
    print('Error:', e)