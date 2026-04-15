#!/usr/bin/env python3
import urllib.request
import time

# Wait a bit for server to be ready
time.sleep(2)

try:
    with urllib.request.urlopen('http://127.0.0.1:8000/download-template/?format=excel') as response:
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