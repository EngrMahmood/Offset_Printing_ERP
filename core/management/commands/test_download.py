from django.core.management.base import BaseCommand
from django.test import RequestFactory
from core.views import download_template

class Command(BaseCommand):
    help = 'Test the download template view'

    def handle(self, *args, **options):
        rf = RequestFactory()
        request = rf.get('/download-template/?format=excel')
        response = download_template(request)
        
        self.stdout.write(f'Status: {response.status_code}')
        self.stdout.write(f'Content-Type: {response.get("Content-Type")}')
        self.stdout.write(f'Content-Disposition: {response.get("Content-Disposition")}')
        self.stdout.write(f'Content length: {len(response.content)}')
        
        if response.content.startswith(b'PK'):
            self.stdout.write('Content starts with PK - this is XLSX')
            # Save to file
            with open('test_download.xlsx', 'wb') as f:
                f.write(response.content)
            self.stdout.write('Saved as test_download.xlsx')
        else:
            self.stdout.write('Content preview:')
            self.stdout.write(response.content[:200].decode('utf-8', errors='ignore'))