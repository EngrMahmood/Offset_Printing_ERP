from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def patch(path, replacements):
    p = ROOT / path
    text = p.read_text(encoding='utf-8')
    orig = text
    for old, new in replacements:
        if old not in text:
            print(f'  skip missing pattern in {p.name}: {old[:60]!r}...')
            continue
        text = text.replace(old, new, 1)
    if text != orig:
        p.write_text(text, encoding='utf-8')
        print(f'patched {path}')
    else:
        print(f'no change {path}')


def patch_dispatch():
    path = 'core/templates/dispatch_records.html'
    p = ROOT / path
    text = p.read_text(encoding='utf-8')
    text = text.replace('<div class="erp-page-stack">', '<div class="erp-page-stack erp-records-shell">', 1)
    old_aside = '<aside class="erp-filter-sidebar">'
    new_aside = '''<aside class="erp-filter-sidebar erp-records-filter-sidebar is-collapsible" data-collapse-key="dispatchRecordsFilterCollapsed">
                <div class="erp-records-filter-head">
                    <h3 class="erp-filter-sidebar-title">Refine Search</h3>
                    <button type="button" class="erp-records-filter-toggle" aria-expanded="true">Collapse</button>
                </div>
                <div class="erp-records-filter-body">'''
    text = text.replace(old_aside, new_aside, 1)
    text = text.replace(
        '<h3 class="erp-filter-sidebar-title">Refine Search</h3>\n\n                    <form method="GET"',
        '<form method="GET"',
        1,
    )
    text = text.replace(
        '                    </form>\n\n                </aside>',
        '                    </form>\n                </div>\n                </aside>',
        1,
    )
    text = text.replace(
        '                <div>\n\n                    {% if has_active_filters %}',
        '                <div class="erp-records-data-pane">\n                    <div class="erp-records-pane-head">\n                    {% if has_active_filters %}',
        1,
    )
    text = text.replace(
        '                    </div>\n\n                    {% endif %}\n\n\n\n                    <div class="erp-table-wrap">',
        '                    </div>\n                    </div>\n                    <div class="erp-table-scroll erp-records-table-wrap">',
        1,
    )
    p.write_text(text, encoding='utf-8')
    print('patched dispatch_records.html')


def patch_production_records():
    path = ROOT / 'production/templates/production/production_records.html'
    text = path.read_text(encoding='utf-8')
    repls = {
        'production-records-shell': 'erp-records-shell',
        'production-filter-sidebar': 'erp-records-filter-sidebar is-collapsible',
        'production-filter-sidebar-head': 'erp-records-filter-head',
        'production-filter-toggle': 'erp-records-filter-toggle',
        'production-filter-body': 'erp-records-filter-body',
        'production-records-data-pane': 'erp-records-data-pane',
        'production-records-pane-head': 'erp-records-pane-head',
        'production-records-table-wrap': 'erp-records-table-wrap',
        'production-records-layout': 'erp-records-layout',
        'id="production_filter_sidebar"': 'data-collapse-key="productionRecordsFilterCollapsed"',
        ' id="production_filter_toggle"': '',
        ' id="production_filter_body"': '',
        ' id="production_records_table_wrap"': '',
        '--production-records-sticky-top': '--erp-records-sticky-top',
        '--production-records-pane-offset': '--erp-records-pane-offset',
        '.erp-content:has(.production-records-shell)': '.erp-content:has(.erp-records-shell)',
    }
    for old, new in repls.items():
        text = text.replace(old, new)
    path.write_text(text, encoding='utf-8')
    print('patched production_records.html')


def patch_production_wip():
    patch('production/templates/production/production_wip.html', [
        (
            '<div class="erp-page-stack production-wip">',
            '<div class="erp-page-stack erp-records-shell production-wip">',
        ),
    ])
    p = ROOT / 'production/templates/production/production_wip.html'
    text = p.read_text(encoding='utf-8')
    text = text.replace(
        '<aside class="erp-filter-sidebar">',
        '''<aside class="erp-filter-sidebar erp-records-filter-sidebar is-collapsible" data-collapse-key="productionWipFilterCollapsed">
            <div class="erp-records-filter-head">
                <h3 class="erp-filter-sidebar-title">Filter WIP Jobs</h3>
                <button type="button" class="erp-records-filter-toggle" aria-expanded="true">Collapse</button>
            </div>
            <div class="erp-records-filter-body">''',
        1,
    )
    text = text.replace(
        '<h3 class="erp-filter-sidebar-title">Filter WIP Jobs</h3>\n                <form',
        '<form',
        1,
    )
    text = text.replace(
        '                </form>\n            </aside>',
        '                </form>\n            </div>\n            </aside>',
        1,
    )
    text = text.replace(
        '            <div>\n                <div class="erp-table-meta">',
        '            <div class="erp-records-data-pane">\n                <div class="erp-records-pane-head">\n                <div class="erp-table-meta">',
        1,
    )
    text = text.replace(
        '                </div>\n\n                <div class="erp-table-scroll">\n                    <table class="erp-table erp-table-dense production-records-table">',
        '                </div>\n                </div>\n                <div class="erp-table-scroll erp-records-table-wrap">\n                    <table class="erp-table erp-table-dense production-records-table">',
        1,
    )
    text = text.replace(
        '<div class="erp-table-scroll">\n                <table class="erp-table erp-table-dense">',
        '<div class="erp-table-scroll erp-records-table-wrap">\n                <table class="erp-table erp-table-dense">',
        1,
    )
    p.write_text(text, encoding='utf-8')
    print('patched production_wip.html')


def patch_plates():
    for rel in [
        'printing_plates/templates/printing_plates/plate_request_list.html',
        'printing_plates/templates/printing_plates/plate_sent_list.html',
        'printing_plates/templates/printing_plates/plate_received_list.html',
        'printing_plates/templates/printing_plates/plate_queue.html',
    ]:
        patch(rel, [
            ('<div class="erp-page-stack">', '<div class="erp-page-stack erp-records-shell">'),
            ('<aside class="erp-filter-sidebar"', '<aside class="erp-filter-sidebar erp-records-filter-sidebar"'),
            ('<div class="erp-table-scroll">', '<div class="erp-table-scroll erp-records-table-wrap">'),
        ])


def patch_other_tables():
    patches = [
        ('planning/templates/planning/pending_skus.html', [
            ('<div class="erp-table-wrap erp-table-scroll">', '<div class="erp-table-wrap erp-table-scroll erp-records-table-wrap">'),
        ]),
        ('qc/templates/qc/master_sku_review_queue.html', [
            ('<div class="erp-table-wrap erp-table-scroll" style="overflow-x:auto;">', '<div class="erp-table-wrap erp-table-scroll erp-records-table-wrap">'),
        ]),
        ('migration/templates/migration/preview.html', [
            ('<div class="erp-table-wrap erp-table-scroll" style="overflow-x:auto;">', '<div class="erp-table-wrap erp-table-scroll erp-records-table-wrap">'),
        ]),
        ('manual_working/templates/manual_working/manual_working_list.html', [
            ('<div class="erp-table-wrap" style="overflow-x:auto;">', '<div class="erp-table-wrap erp-table-scroll erp-records-table-wrap">'),
        ]),
        ('core/templates/archived_records.html', []),
    ]
    for path, reps in patches:
        if reps:
            patch(path, reps)


if __name__ == '__main__':
    patch_dispatch()
    patch_production_records()
    patch_production_wip()
    patch_plates()
    patch_other_tables()
