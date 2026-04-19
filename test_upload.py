import requests
import os

session = requests.Session()
url = 'http://127.0.0.1:8000/planning/sku-recipes/bulk-upload/'
file_path = 'planning/docs/sku_recipe_upload_sample_2.xlsx'

# Get CSRF token
login_url = 'http://127.0.0.1:8000/' # Or any page
r = session.get(login_url)
csrf_token = session.cookies.get('csrftoken')

if not csrf_token:
    print('Error: Could not obtain CSRF token')
    # Maybe try a more specific URL if the above doesn't work
    # exit(1)

files = {'upload_file': open(file_path, 'rb')}
data = {'csrfmiddlewaretoken': csrf_token}
headers = {'Referer': login_url}

try:
    response = session.post(url, files=files, data=data, headers=headers)
    print(f'Status Code: {response.status_code}')
    print(f'Response snippet: {response.text[:500]}')
except Exception as e:
    print(f'Request failed: {e}')
