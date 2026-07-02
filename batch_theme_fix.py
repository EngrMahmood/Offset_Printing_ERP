"""Batch replace theme.css links with erp-global-theme.css (via theme_colors include)."""
import os
import re

BASE = 'd:/Offset_Printing_ERP'

files = [
    'core/templates/master_data.html',
    'core/templates/shift_config.html',
    'core/templates/archived_records.html',
    'core/templates/change_history.html',
    'core/templates/confirm_delete.html',
    'core/templates/confirm_restore.html',
    'core/templates/dispatch_entry.html',
    'core/templates/erp_readme.html',
    'core/templates/manage_user_roles.html',
    'core/templates/override_requests.html',
    'core/templates/request_edit_override.html',
    'core/templates/review_override_request.html',
    'core/templates/upload.html',
    'planning/templates/planning/approval_queue.html',
    'planning/templates/planning/manual_po_entry.html',
    'planning/templates/planning/pending_sku_master_entry.html',
    'planning/templates/planning/pending_skus.html',
    'planning/templates/planning/pending_skus_ignored.html',
    'planning/templates/planning/planning_archived_jobs.html',
    'planning/templates/planning/planning_import.html',
    'planning/templates/planning/planning_job_edit.html',
    'planning/templates/planning/planning_readme.html',
    'planning/templates/planning/planning_report.html',
    'planning/templates/planning/planning_scan.html',
    'planning/templates/planning/planning_welcome.html',
    'planning/templates/planning/po_new_skus.html',
    'planning/templates/planning/po_review.html',
    'planning/templates/planning/po_upload.html',
    'planning/templates/planning/sku_recipe_bulk_upload.html',
    'planning/templates/planning/sku_recipe_edit.html',
    'planning/templates/planning/sku_recipes.html',
    'planning/templates/planning/sku_recipes_archived.html',
]

# Pattern matches <link rel="stylesheet" href="...theme.css...">
theme_css_pattern = re.compile(
    r'<link\s+rel=["\']stylesheet["\']\s+href=["\'][^"\']*theme\.css[^"\']*["\']\s*[^>]*>',
    re.IGNORECASE
)
replace_with = "{% include 'includes/ui/theme_colors.html' %}"

# Also fix body style - replace old body background with ERP variable
old_body_styles = [
    "background-color: #f5f5f5;",
    "background-color: #f3f6f9;",
    "background: #f3f6f9;",
    "background: #edf1f5;",
    "background: #f1f3f6;",
    "background: #f3f5f8;",
]

# Fix old nav/button classes in HTML (not in CSS definitions)
btn_replacements = [
    ('class="nav-btn"', 'class="erp-btn erp-btn-secondary"'),
    ("class='nav-btn'", "class='erp-btn erp-btn-secondary'"),
    ('class="btn"', 'class="erp-btn erp-btn-secondary"'),
    ("class='btn'", "class='erp-btn erp-btn-secondary'"),
    ('class="btn btn-green"', 'class="erp-btn erp-btn-success"'),
    ('class="btn btn-danger"', 'class="erp-btn erp-btn-danger"'),
    ('class="btn btn-primary"', 'class="erp-btn erp-btn-primary"'),
    ('class="btn btn-amber"', 'class="erp-btn erp-btn-warning"'),
    ('class="btn btn-warning"', 'class="erp-btn erp-btn-warning"'),
]

changed = 0
skipped = 0

for f in files:
    full = os.path.join(BASE, f).replace('/', os.sep)
    if not os.path.exists(full):
        print(f'MISSING: {f}')
        continue
    c = open(full, encoding='utf-8', errors='ignore').read()

    # Skip if already fully themed
    if 'theme_colors' in c:
        print(f'SKIP (already): {f}')
        skipped += 1
        continue

    modified = False

    # 1. Replace theme.css link
    if theme_css_pattern.search(c):
        c = theme_css_pattern.sub(replace_with, c)
        modified = True
    elif 'theme.css' not in c and '</head>' in c:
        # No CSS at all - inject
        c = c.replace('</head>', replace_with + '\n</head>', 1)
        modified = True

    # 2. Fix body background in inline styles (replace old backgrounds)
    for old_bg in old_body_styles:
        if old_bg in c:
            c = c.replace(old_bg, 'background: var(--erp-bg-soft, #f3f5f8);')
            modified = True

    # 3. Fix font-family in body style
    if "font-family: Arial, sans-serif;" in c:
        c = c.replace(
            "font-family: Arial, sans-serif;",
            "font-family: 'IBM Plex Sans', Arial, sans-serif;"
        )
        modified = True

    if modified:
        open(full, 'w', encoding='utf-8').write(c)
        print(f'FIXED: {f}')
        changed += 1
    else:
        print(f'NO_CHANGE: {f}')

print(f'\nDone. Changed: {changed}, Skipped: {skipped}')
