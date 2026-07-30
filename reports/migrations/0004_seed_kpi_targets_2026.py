from django.db import migrations


KPI_ROWS_2026 = [
    {
        'kpi_slug': 'order_fulfillment',
        'year': 2026,
        'position': 'Sr. Manager',
        'objective': 'Operational Excellence',
        'title': 'Order Fulfillment Efficiency (%)',
        'description': (
            'Achieve 80% Order Fulfillment efficiency by maximizing actual output against '
            'planned capacity through optimized machine utilization and reduced downtime.'
        ),
        'uom': '%',
        'weightage_pct': 20,
        'monitoring_frequency': 'Quarterly',
        'min_value': 80,
        'target_value': 85,
        'max_value': 100,
        'higher_is_better': True,
    },
    {
        'kpi_slug': 'wastage_reduction',
        'year': 2026,
        'position': 'Sr. Manager',
        'objective': 'Quality',
        'title': 'Wastage Reduction Efficiency (%)',
        'description': (
            'Maintain overall production wastage within defined limits by controlling '
            'setup, running, and rejection waste.'
        ),
        'uom': '%',
        'weightage_pct': 20,
        'monitoring_frequency': 'Quarterly',
        'min_value': 0,
        'target_value': 5,
        'max_value': 8,
        'higher_is_better': False,
    },
    {
        'kpi_slug': 'dispatch_alignment',
        'year': 2026,
        'position': 'Sr. Manager',
        'objective': 'Operational Excellence',
        'title': 'Dispatch vs Production Alignment (%)',
        'description': (
            'Ensure that cumulative quantity dispatched is at least 95% of quantity produced '
            'on a quarterly basis, to prevent floor choking, ensure dispatch alignment, and '
            'avoid material stagnation.'
        ),
        'uom': '%',
        'weightage_pct': 15,
        'monitoring_frequency': 'Quarterly',
        'min_value': 80,
        'target_value': 95,
        'max_value': 130,
        'higher_is_better': True,
    },
]


def seed_kpi_targets(apps, schema_editor):
    KPITarget = apps.get_model('reports', 'KPITarget')
    for row in KPI_ROWS_2026:
        KPITarget.objects.update_or_create(
            kpi_slug=row['kpi_slug'], year=row['year'], defaults=row,
        )


def unseed_kpi_targets(apps, schema_editor):
    KPITarget = apps.get_model('reports', 'KPITarget')
    KPITarget.objects.filter(year=2026, kpi_slug__in=[row['kpi_slug'] for row in KPI_ROWS_2026]).delete()


class Migration(migrations.Migration):

    dependencies = [
        ('reports', '0003_kpitarget_kpiactionnote'),
    ]

    operations = [
        migrations.RunPython(seed_kpi_targets, unseed_kpi_targets),
    ]
