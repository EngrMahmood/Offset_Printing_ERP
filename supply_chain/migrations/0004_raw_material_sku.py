from django.db import migrations, models
import django.db.models.deletion


def copy_supply_chain_items_to_raw_skus(apps, schema_editor):
    SupplyChainItem = apps.get_model('supply_chain', 'SupplyChainItem')
    RawMaterialSku = apps.get_model('supply_chain', 'RawMaterialSku')
    StockTransaction = apps.get_model('supply_chain', 'StockTransaction')
    StockDemand = apps.get_model('supply_chain', 'StockDemand')
    PhysicalStockCount = apps.get_model('supply_chain', 'PhysicalStockCount')

    mapping = {}
    for legacy in SupplyChainItem.objects.select_related('material').all():
        sku_code = legacy.item_id or f'LEGACY-{legacy.pk}'
        purchase_size = 'UNSPECIFIED'
        raw, _created = RawMaterialSku.objects.get_or_create(
            material_id=legacy.material_id,
            purchase_sheet_size=purchase_size,
            defaults={
                'sku': sku_code,
                'uom': legacy.uom,
                'sheet_packing_pcs': legacy.sheet_packing_pcs,
                'unit_cost': legacy.unit_cost,
                'safety_stock': legacy.safety_stock,
                'max_stock_level': legacy.max_stock_level,
                'lead_time_days': legacy.lead_time_days,
                'is_active': True,
            },
        )
        if raw.sku != sku_code and not RawMaterialSku.objects.filter(sku=sku_code).exists():
            raw.sku = sku_code
            raw.save(update_fields=['sku'])
        mapping[legacy.pk] = raw.pk

    for txn in StockTransaction.objects.all().iterator():
        if txn.item_id and txn.item_id in mapping:
            txn.raw_material_sku_id = mapping[txn.item_id]
            txn.save(update_fields=['raw_material_sku_id'])

    for demand in StockDemand.objects.all().iterator():
        if demand.item_id and demand.item_id in mapping:
            demand.raw_material_sku_id = mapping[demand.item_id]
            demand.save(update_fields=['raw_material_sku_id'])

    for count in PhysicalStockCount.objects.all().iterator():
        if count.item_id and count.item_id in mapping:
            count.raw_material_sku_id = mapping[count.item_id]
            count.save(update_fields=['raw_material_sku_id'])


class Migration(migrations.Migration):

    dependencies = [
        ('core', '0063_auto_sync_and_backfill_materials'),
        ('supply_chain', '0003_physical_stock_count'),
    ]

    operations = [
        migrations.CreateModel(
            name='RawMaterialSku',
            fields=[
                ('id', models.BigAutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('sku', models.CharField(max_length=50, unique=True, verbose_name='Raw Material SKU')),
                ('purchase_sheet_size', models.CharField(max_length=80, verbose_name='Purchase Sheet Size')),
                ('uom', models.CharField(default='Sheets', max_length=20, verbose_name='UOM')),
                ('sheet_packing_pcs', models.IntegerField(default=1, verbose_name='Sheet Packing/Pcs')),
                ('unit_cost', models.DecimalField(decimal_places=2, default=0.0, max_digits=12, verbose_name='Unit Cost')),
                ('safety_stock', models.IntegerField(default=0, verbose_name='Safety Stock')),
                ('max_stock_level', models.IntegerField(default=10000, verbose_name='Maximum Stock Level')),
                ('lead_time_days', models.IntegerField(default=1, verbose_name='Lead Time (Days)')),
                ('is_active', models.BooleanField(db_index=True, default=True)),
                ('created_at', models.DateTimeField(auto_now_add=True)),
                ('updated_at', models.DateTimeField(auto_now=True)),
                ('material', models.ForeignKey(on_delete=django.db.models.deletion.PROTECT, related_name='raw_material_skus', to='core.material')),
            ],
            options={
                'verbose_name': 'Raw Material SKU',
                'verbose_name_plural': 'Raw Material SKUs',
                'ordering': ['material__name', 'purchase_sheet_size', 'sku'],
                'unique_together': {('material', 'purchase_sheet_size')},
            },
        ),
        migrations.AddField(
            model_name='stocktransaction',
            name='raw_material_sku',
            field=models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.CASCADE, related_name='transactions', to='supply_chain.rawmaterialsku'),
        ),
        migrations.AddField(
            model_name='stockdemand',
            name='raw_material_sku',
            field=models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.CASCADE, related_name='demands', to='supply_chain.rawmaterialsku'),
        ),
        migrations.AddField(
            model_name='physicalstockcount',
            name='raw_material_sku',
            field=models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.CASCADE, related_name='physical_counts', to='supply_chain.rawmaterialsku'),
        ),
        migrations.AlterField(
            model_name='stocktransaction',
            name='item',
            field=models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.CASCADE, related_name='legacy_transactions', to='supply_chain.supplychainitem'),
        ),
        migrations.AlterField(
            model_name='stockdemand',
            name='item',
            field=models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.CASCADE, related_name='legacy_demands', to='supply_chain.supplychainitem'),
        ),
        migrations.AlterField(
            model_name='physicalstockcount',
            name='item',
            field=models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.CASCADE, related_name='legacy_physical_counts', to='supply_chain.supplychainitem'),
        ),
        migrations.RunPython(copy_supply_chain_items_to_raw_skus, migrations.RunPython.noop),
    ]
