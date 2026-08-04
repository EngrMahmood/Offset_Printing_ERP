"""Detailed breakdown of user update file review."""
import os, sys
from pathlib import Path
from collections import Counter, defaultdict

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'Offset_ERP.settings')
import django
django.setup()

from core.models import ProductType
import importlib.util

spec = importlib.util.spec_from_file_location('rev', ROOT / 'scripts' / 'data_migration_archive' / 'review_master_data_update.py')
rev = importlib.util.module_from_spec(spec)
spec.loader.exec_module(rev)

erp_pt = set(ProductType.objects.values_list('name', flat=True))
_, rows = rev.load_rows()

by_sku = defaultdict(list)
for r in rows:
    by_sku[r['sku'].casefold()].append(r)

print('=== DUPLICATES ===')
for key, items in sorted(by_sku.items(), key=lambda x: -len(x[1])):
    if len(items) > 1:
        print(items[0]['sku'], '|', [i['sheet'] for i in items])

err_types = Counter()
ready = 0
for r in rows:
    issues = []
    pt = r['resolved_product_type']
    jp = r['resolved_job_process']
    pp = rev.parse_passes(r['resolved_print_passes'])
    if not pt:
        issues.append('missing product_type')
    elif pt not in erp_pt:
        issues.append('invalid product_type')
    if not jp:
        issues.append('missing job_process')
    elif jp not in {'print_and_pack', 'cut_and_pack'}:
        issues.append('invalid job_process')
    if jp == 'print_and_pack' and pp in (None, ''):
        issues.append('missing print_passes')
    if jp == 'cut_and_pack' and pp not in (None, ''):
        issues.append('cut_and_pack has passes')
    if issues:
        for i in issues:
            err_types[i] += 1
    else:
        ready += 1

print('\n=== ERROR BREAKDOWN ===')
for k, c in err_types.most_common():
    print(c, k)

print('\nReady:', ready, 'Errors:', len(rows) - ready)

# per sheet ready count
for sheet in sorted(set(r['sheet'] for r in rows)):
    subset = [r for r in rows if r['sheet'] == sheet]
    ok = 0
    for r in subset:
        pt = r['resolved_product_type']
        jp = r['resolved_job_process']
        pp = rev.parse_passes(r['resolved_print_passes'])
        if pt and pt in erp_pt and jp in {'print_and_pack', 'cut_and_pack'}:
            if jp == 'cut_and_pack' and pp in (None, ''):
                ok += 1
            elif jp == 'print_and_pack' and pp in (1, 2, 3):
                ok += 1
    print(f'{sheet}: {len(subset)} rows, ready {ok}, errors {len(subset)-ok}')
