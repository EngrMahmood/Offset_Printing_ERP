"""Export review status CSV for user update file."""
import os, sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'Offset_ERP.settings')
import django
django.setup()

import csv
import importlib.util
from collections import defaultdict
from core.models import ProductType

spec = importlib.util.spec_from_file_location('rev', ROOT / 'scripts' / 'review_master_data_update.py')
rev = importlib.util.module_from_spec(spec)
spec.loader.exec_module(rev)

erp_pt = set(ProductType.objects.values_list('name', flat=True))
_, rows = rev.load_rows()
by_sku = defaultdict(list)
for r in rows:
    by_sku[r['sku'].casefold()].append(r)

out = ROOT / 'migration_reports' / 'update_file_review_status.csv'
fields = [
    'review_status', 'sheet', 'sku', 'duplicate', 'resolved_product_type',
    'resolved_job_process', 'resolved_print_passes', 'issues',
]

with open(out, 'w', newline='', encoding='utf-8-sig') as f:
    w = csv.DictWriter(f, fieldnames=fields)
    w.writeheader()
    for r in rows:
        issues = []
        pt = r['resolved_product_type']
        jp = r['resolved_job_process']
        pp = rev.parse_passes(r['resolved_print_passes'])
        dup = 'Y' if len(by_sku[r['sku'].casefold()]) > 1 else 'N'
        if dup == 'Y':
            issues.append('duplicate SKU in multiple sheets')
        if not pt:
            issues.append('missing product_type')
        elif pt not in erp_pt:
            issues.append(f'invalid product_type: {pt}')
        if not jp:
            issues.append('missing job_process')
        elif jp not in {'print_and_pack', 'cut_and_pack'}:
            issues.append('invalid job_process')
        if jp == 'print_and_pack' and pp in (None, ''):
            issues.append('missing print_passes')
        elif jp == 'print_and_pack' and (pp == 'INVALID' or pp not in rev.VALID_PRINT_PASSES):
            issues.append(f'invalid print_passes: {r["resolved_print_passes"]}')
        if jp == 'cut_and_pack' and pp not in (None, ''):
            issues.append('cut_and_pack must not have passes')
        status = 'READY' if not issues else 'FIX REQUIRED'
        w.writerow({
            'review_status': status,
            'sheet': r['sheet'],
            'sku': r['sku'],
            'duplicate': dup,
            'resolved_product_type': pt,
            'resolved_job_process': jp,
            'resolved_print_passes': r['resolved_print_passes'],
            'issues': '; '.join(issues),
        })

print('Wrote', out)
