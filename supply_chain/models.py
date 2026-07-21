import re

from django.db import models
from django.utils import timezone
from django.contrib.auth.models import User



# Width × height: "10.5 x 15", "10.5X15", "10.5*15", "10.5 × 15" → "10.5*15"
_PURCHASE_SIZE_PAIR_RE = re.compile(
    r'^\s*(\d+(?:\.\d+)?)\s*[xX×*]\s*(\d+(?:\.\d+)?)\s*$'
)


def normalize_purchase_sheet_size(value):
    """Canonical purchase sheet size for matching and rollup.

    Dimension pairs become ``width*height`` (no spaces). Other labels keep
    collapsed whitespace only (e.g. ``A4``).
    """
    text = ' '.join(str(value or '').strip().split())
    if not text:
        return ''
    match = _PURCHASE_SIZE_PAIR_RE.match(text)
    if match:
        return f'{match.group(1)}*{match.group(2)}'
    return text


def normalize_material_name(value):
    """Case-insensitive material key: collapse whitespace and lowercase."""
    return ' '.join(str(value or '').strip().split()).lower()


def display_material_name(value):
    """Stable display label for free-text material names."""
    text = ' '.join(str(value or '').strip().split())
    if not text:
        return ''
    if text.isupper() or text.islower():
        return text.title()
    return text


class RawMaterialSku(models.Model):
    """Procurement + inventory identity: material type and purchase sheet size."""

    sku = models.CharField(max_length=50, unique=True, verbose_name='Raw Material SKU')
    material = models.ForeignKey(
        'core.Material',
        on_delete=models.PROTECT,
        related_name='raw_material_skus',
    )
    purchase_sheet_size = models.CharField(max_length=80, verbose_name='Purchase Sheet Size')
    sku_type = models.ForeignKey(
        'supply_chain.ItemRequestType',
        on_delete=models.PROTECT,
        null=True,
        blank=True,
        related_name='raw_material_skus',
        verbose_name='SKU Type',
        help_text='Raw Material, Consumable, Service, Maintenance, …',
    )
    uom = models.CharField(max_length=20, verbose_name='UOM', default='Sheets')
    sheet_packing_pcs = models.IntegerField(verbose_name='Sheet Packing/Pcs', default=1)
    unit_cost = models.DecimalField(max_digits=12, decimal_places=2, default=0.0, verbose_name='Unit Cost')
    safety_stock = models.IntegerField(default=0, verbose_name='Safety Stock')
    max_stock_level = models.IntegerField(default=10000, verbose_name='Maximum Stock Level')
    lead_time_days = models.IntegerField(default=1, verbose_name='Lead Time (Days)')
    is_active = models.BooleanField(default=True, db_index=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['material__name', 'purchase_sheet_size', 'sku']
        unique_together = ('material', 'purchase_sheet_size')
        verbose_name = 'Raw Material SKU'
        verbose_name_plural = 'Raw Material SKUs'

    def __str__(self):
        return self.display_label

    @property
    def display_label(self):
        return f'{self.material.name} — {self.purchase_sheet_size} ({self.sku})'

    @property
    def item_id(self):
        return self.sku


class SupplyChainItem(models.Model):
    """Legacy inventory row — superseded by RawMaterialSku."""

    material = models.OneToOneField(
        'core.Material',
        on_delete=models.CASCADE,
        related_name='supply_chain_details',
    )
    item_id = models.CharField(max_length=50, unique=True, verbose_name='Item ID', blank=True, null=True)
    uom = models.CharField(max_length=20, verbose_name='UOM', default='Sheets')
    sheet_packing_pcs = models.IntegerField(verbose_name='Sheet Packing/Pcs', default=1)
    unit_cost = models.DecimalField(max_digits=12, decimal_places=2, default=0.0, verbose_name='Unit Cost')
    safety_stock = models.IntegerField(default=0, verbose_name='Safety Stock')
    max_stock_level = models.IntegerField(default=10000, verbose_name='Maximum Stock Level')
    lead_time_days = models.IntegerField(default=1, verbose_name='Lead Time (Days)')
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return f'{self.item_id or "N/A"} - {self.material.name}'


class StockTransaction(models.Model):
    TRANSACTION_TYPES = [
        ('OPENING', 'Opening'),
        ('RECEIVING', 'Receiving'),
        ('ISSUANCE', 'Issuance'),
        ('ADJUSTMENT', 'Adjustment'),
    ]
    SOURCE_CHOICES = [
        ('MANUAL', 'Manual'),
        ('JOB_CARD', 'Job Card'),
    ]

    raw_material_sku = models.ForeignKey(
        RawMaterialSku,
        on_delete=models.CASCADE,
        related_name='transactions',
        null=True,
        blank=True,
    )
    item = models.ForeignKey(
        SupplyChainItem,
        on_delete=models.CASCADE,
        related_name='legacy_transactions',
        null=True,
        blank=True,
    )
    transaction_type = models.CharField(max_length=20, choices=TRANSACTION_TYPES)
    source = models.CharField(max_length=20, choices=SOURCE_CHOICES, default='MANUAL')
    job_card = models.ForeignKey(
        'core.JobCard',
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='stock_transactions',
    )
    production = models.OneToOneField(
        'core.Production',
        on_delete=models.CASCADE,
        null=True,
        blank=True,
        related_name='stock_issuance',
    )
    month_str = models.CharField(max_length=20, blank=True, null=True, verbose_name='Month')
    date = models.DateField(default=timezone.now, verbose_name='Date')
    gin_jc = models.CharField(max_length=50, blank=True, null=True, verbose_name='GIN / JC')
    sheet_qty_pcs = models.IntegerField(default=0, verbose_name='Sheet Qty/Pcs')
    pkt_rim_qty = models.IntegerField(default=0, verbose_name='Pkt/Rim Qty')
    is_active = models.BooleanField(default=True, db_index=True)
    is_approved = models.BooleanField(default=True, db_index=True, verbose_name='Is Approved')
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        sku = self.raw_material_sku.sku if self.raw_material_sku_id else 'N/A'
        return f'{self.get_transaction_type_display()} - {sku} - {self.date}'


class StockDemand(models.Model):
    raw_material_sku = models.ForeignKey(
        RawMaterialSku,
        on_delete=models.CASCADE,
        related_name='demands',
        null=True,
        blank=True,
    )
    item = models.ForeignKey(
        SupplyChainItem,
        on_delete=models.CASCADE,
        related_name='legacy_demands',
        null=True,
        blank=True,
    )
    month_str = models.CharField(max_length=20, blank=True, null=True, verbose_name='Month')
    sheet_qty_pcs = models.IntegerField(default=0, verbose_name='Sheet Qty/Pcs')
    pkt_rim_qty = models.IntegerField(default=0, verbose_name='Pkt/Rim Qty')
    is_active = models.BooleanField(default=True, db_index=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        sku = self.raw_material_sku.sku if self.raw_material_sku_id else 'N/A'
        return f'Demand - {sku} - {self.month_str}'


class PhysicalStockCount(models.Model):
    raw_material_sku = models.ForeignKey(
        RawMaterialSku,
        on_delete=models.CASCADE,
        related_name='physical_counts',
        null=True,
        blank=True,
    )
    item = models.ForeignKey(
        SupplyChainItem,
        on_delete=models.CASCADE,
        related_name='legacy_physical_counts',
        null=True,
        blank=True,
    )
    count_date = models.DateField(default=timezone.now, verbose_name='Count Date')
    physical_sheet_qty = models.IntegerField(default=0, verbose_name='Physical Sheet Qty/Pcs')
    physical_pkt_rim_qty = models.IntegerField(default=0, verbose_name='Physical Pkt/Rim Qty')
    system_sheet_qty = models.IntegerField(default=0, verbose_name='System Sheet Qty/Pcs')
    system_pkt_rim_qty = models.IntegerField(default=0, verbose_name='System Pkt/Rim Qty')
    accuracy_percent = models.DecimalField(
        max_digits=6,
        decimal_places=2,
        default=0,
        verbose_name='Inventory Accuracy %',
    )
    notes = models.CharField(max_length=255, blank=True, default='')
    is_active = models.BooleanField(default=True, db_index=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['-count_date', '-id']

    def __str__(self):
        sku = self.raw_material_sku.sku if self.raw_material_sku_id else 'N/A'
        return f'Count {sku} @ {self.count_date} — {self.accuracy_percent}%'

    @property
    def variance(self):
        return self.physical_sheet_qty - self.system_sheet_qty


class ChangeRequest(models.Model):
    MODEL_CHOICES = [
        ('RawMaterialSku', 'Raw Material SKU'),
        ('StockDemand', 'Monthly Demand'),
        ('StockTransaction', 'Stock Transaction'),
        ('PhysicalStockCount', 'Physical Stock Count'),
        ('ItemRequest', 'Item Request'),
    ]

    ACTION_CHOICES = [
        ('CREATE', 'Create'),
        ('UPDATE', 'Update'),
        ('DELETE', 'Delete'),
    ]

    STATUS_CHOICES = [
        ('PENDING', 'Pending Review'),
        ('APPROVED', 'Approved'),
        ('REJECTED', 'Rejected'),
    ]

    model_name = models.CharField(max_length=50, choices=MODEL_CHOICES)
    action = models.CharField(max_length=10, choices=ACTION_CHOICES)
    target_id = models.PositiveIntegerField(null=True, blank=True)
    proposed_data = models.JSONField(help_text="Proposed field values serialized to JSON", default=dict, blank=True)

    requested_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='supply_chain_change_requests')
    requested_at = models.DateTimeField(auto_now_add=True)

    status = models.CharField(max_length=10, choices=STATUS_CHOICES, default='PENDING')
    reviewed_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='reviewed_supply_chain_change_requests')
    reviewed_at = models.DateTimeField(null=True, blank=True)
    rejection_reason = models.TextField(blank=True, default='')

    class Meta:
        ordering = ['-requested_at']

    def __str__(self):
        return f"{self.get_action_display()} {self.get_model_name_display()} - {self.get_status_display()}"

    @property
    def target_object(self):
        if not self.target_id:
            return None
        from django.apps import apps
        try:
            model_class = apps.get_model('supply_chain', self.model_name)
            return model_class.objects.get(pk=self.target_id)
        except Exception:
            return None

    def get_old_and_new_values(self):
        if self.action != 'UPDATE':
            return []
        target = self.target_object
        if not target:
            return []
        changes = []
        from django.apps import apps
        model_class = apps.get_model('supply_chain', self.model_name)
        for field_name, new_val in self.proposed_data.items():
            is_fk = False
            actual_field_name = field_name
            if field_name.endswith('_id'):
                is_fk = True
                actual_field_name = field_name[:-3]
            try:
                field = model_class._meta.get_field(actual_field_name)
                label = field.verbose_name or actual_field_name.replace('_', ' ').title()
            except Exception:
                label = actual_field_name.replace('_', ' ').title()

            if is_fk:
                old_id = getattr(target, field_name)
                old_val = '-'
                if old_id:
                    try:
                        related_model = field.related_model
                        old_val = str(related_model.objects.get(pk=old_id))
                    except Exception:
                        old_val = f"ID: {old_id}"
                new_val_str = '-'
                if new_val:
                    try:
                        related_model = field.related_model
                        new_val_str = str(related_model.objects.get(pk=new_val))
                    except Exception:
                        new_val_str = f"ID: {new_val}"
                changes.append({
                    'field': actual_field_name,
                    'label': label,
                    'old': old_val,
                    'new': new_val_str
                })
            else:
                old_val = getattr(target, actual_field_name)
                if isinstance(old_val, bool):
                    old_val = 'Yes' if old_val else 'No'
                if isinstance(new_val, bool):
                    new_val = 'Yes' if new_val else 'No'
                changes.append({
                    'field': actual_field_name,
                    'label': label,
                    'old': old_val if old_val is not None else '-',
                    'new': new_val if new_val is not None else '-'
                })
        return changes

    def get_proposed_fields(self):
        from django.apps import apps
        try:
            model_class = apps.get_model('supply_chain', self.model_name)
        except Exception:
            return []
        fields = []
        for field_name, val in self.proposed_data.items():
            is_fk = False
            actual_field_name = field_name
            if field_name.endswith('_id'):
                is_fk = True
                actual_field_name = field_name[:-3]
            try:
                field = model_class._meta.get_field(actual_field_name)
                label = field.verbose_name or actual_field_name.replace('_', ' ').title()
            except Exception:
                label = actual_field_name.replace('_', ' ').title()

            if is_fk:
                val_str = '-'
                if val:
                    try:
                        related_model = field.related_model
                        val_str = str(related_model.objects.get(pk=val))
                    except Exception:
                        val_str = f"ID: {val}"
                fields.append({'label': label, 'value': val_str})
            else:
                if isinstance(val, bool):
                    val = 'Yes' if val else 'No'
                fields.append({'label': label, 'value': val if val is not None else '-'})
        return fields

    def apply(self, user):
        from django.apps import apps
        model_class = apps.get_model('supply_chain', self.model_name)
        if self.action == 'CREATE':
            if self.model_name == 'RawMaterialSku' and 'material_name' in self.proposed_data:
                from .raw_material_sku import upsert_raw_material_sku_row
                obj, errors, _ = upsert_raw_material_sku_row(self.proposed_data)
                if errors:
                    raise Exception('; '.join(errors))
                self.target_id = obj.pk
            else:
                instance = model_class(**self.proposed_data)
                instance.save()
                self.target_id = instance.pk
        elif self.action == 'UPDATE':
            instance = model_class.objects.get(pk=self.target_id)
            for k, v in self.proposed_data.items():
                setattr(instance, k, v)
            instance.save()
        elif self.action == 'DELETE':
            instance = model_class.objects.get(pk=self.target_id)
            # Soft delete!
            instance.is_active = False
            instance.save(update_fields=['is_active'])

        self.status = 'APPROVED'
        self.reviewed_by = user
        self.reviewed_at = timezone.now()
        self.save()


class ItemRequestType(models.Model):
    name = models.CharField(max_length=50, unique=True)
    code = models.CharField(max_length=4, help_text='Short prefix used in the IR-ID, e.g. RM, CON, MNT')
    is_active = models.BooleanField(default=True, db_index=True)

    class Meta:
        ordering = ['name']
        verbose_name = 'Item Request Type'
        verbose_name_plural = 'Item Request Types'

    def __str__(self):
        return self.name


class ItemRequestDepartment(models.Model):
    """Department lookup dedicated to the Item Request module (separate from core.Department)."""

    name = models.CharField(max_length=100, unique=True)
    is_active = models.BooleanField(default=True, db_index=True)

    class Meta:
        ordering = ['name']
        verbose_name = 'Item Request Department'
        verbose_name_plural = 'Item Request Departments'

    def __str__(self):
        return self.name


class ItemRequest(models.Model):
    LOCAL_IMPORT_CHOICES = [
        ('LOCAL', 'Local'),
        ('IMPORT', 'Import'),
    ]

    STATUS_CHOICES = [
        ('SUBMITTED', 'Submitted'),
        ('MGR_REVIEW', 'Manager Review'),
        ('SC_REVIEW', 'Supply Chain Review'),
        ('APPROVED', 'Approved'),
        ('IN_PROCUREMENT', 'In Procurement'),
        ('RECEIVED', 'Received'),
        ('CLOSED', 'Closed'),
        ('REJECTED', 'Rejected'),
        ('NEEDS_REVISION', 'Needs Revision'),
    ]

    OPEN_STATUSES = {'SUBMITTED', 'MGR_REVIEW', 'SC_REVIEW', 'APPROVED', 'IN_PROCUREMENT', 'NEEDS_REVISION'}

    request_no = models.CharField(max_length=30, unique=True, blank=True, null=True, verbose_name='IR-ID')
    request_type = models.ForeignKey(ItemRequestType, on_delete=models.PROTECT, related_name='requests')
    request_date = models.DateField(default=timezone.now)

    item_title = models.CharField(max_length=255, verbose_name='Item Title')
    machine = models.ForeignKey('core.Machine', on_delete=models.SET_NULL, null=True, blank=True, related_name='item_requests')
    machine_other = models.CharField(max_length=150, blank=True, default='', verbose_name='Machine (if not in list)')

    uom = models.CharField(max_length=30, verbose_name='Unit of Measure')
    specifications = models.TextField(verbose_name='Specifications / Technical Details')
    description = models.TextField(blank=True, default='', verbose_name='Description / Purpose of Use')
    dimensions = models.CharField(max_length=100, blank=True, default='', verbose_name='Dimensions (if applicable)')
    local_import = models.CharField(max_length=10, choices=LOCAL_IMPORT_CHOICES, blank=True, default='')
    part_number = models.CharField(max_length=100, blank=True, default='', verbose_name='Part Number')
    required_quantity = models.DecimalField(max_digits=12, decimal_places=2, default=0, verbose_name='Required Quantity')

    department = models.ForeignKey(ItemRequestDepartment, on_delete=models.PROTECT, related_name='item_requests')
    existing_sku = models.ForeignKey(RawMaterialSku, on_delete=models.SET_NULL, null=True, blank=True, related_name='item_requests')
    estimated_unit_price = models.DecimalField(max_digits=12, decimal_places=2, null=True, blank=True)
    cost_centre = models.CharField(max_length=50, blank=True, default='', verbose_name='Budget / Cost Centre')
    attachment = models.FileField(upload_to='item_requests/%Y/%m/', null=True, blank=True)

    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='SUBMITTED', db_index=True)
    raised_by = models.ForeignKey(User, on_delete=models.CASCADE, related_name='item_requests')
    is_active = models.BooleanField(default=True, db_index=True)
    deleted_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='item_requests_deleted')
    deleted_at = models.DateTimeField(null=True, blank=True)

    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['-created_at']
        verbose_name = 'Item Request'
        verbose_name_plural = 'Item Requests'

    def __str__(self):
        return self.request_no or f'IR-DRAFT-{self.pk}'

    @property
    def machine_display(self):
        if self.machine_id:
            return self.machine.name
        return self.machine_other or '-'

    @property
    def is_open(self):
        return self.status in self.OPEN_STATUSES

    @classmethod
    def find_open_duplicates(cls, item_title, department_id, exclude_pk=None):
        """Fuzzy match: open requests for the same department with a similar item title."""
        key = normalize_material_name(item_title)
        if not key:
            return cls.objects.none()
        qs = cls.objects.filter(
            is_active=True,
            department_id=department_id,
            status__in=cls.OPEN_STATUSES,
        )
        if exclude_pk:
            qs = qs.exclude(pk=exclude_pk)
        matches = [r.pk for r in qs.only('pk', 'item_title') if normalize_material_name(r.item_title) == key]
        return cls.objects.filter(pk__in=matches)


class ItemRequestApproval(models.Model):
    ACTION_CHOICES = [
        ('SUBMIT', 'Submit'),
        ('APPROVE', 'Approve'),
        ('REJECT', 'Reject'),
        ('REVISE', 'Needs Revision'),
        ('RESUBMIT', 'Resubmit'),
        ('EDIT', 'Edited (post-approval)'),
        ('EDIT_REQUESTED', 'Edit Requested (pending change review)'),
        ('DELETE', 'Deleted'),
    ]
    STAGE_CHOICES = [
        ('REQUESTER', 'Requester'),
        ('MANAGER', 'Manager'),
        ('SUPPLY_CHAIN', 'Supply Chain'),
    ]

    request = models.ForeignKey(ItemRequest, on_delete=models.CASCADE, related_name='approvals')
    actor = models.ForeignKey(User, on_delete=models.CASCADE, related_name='item_request_approvals')
    action = models.CharField(max_length=20, choices=ACTION_CHOICES)
    stage = models.CharField(max_length=20, choices=STAGE_CHOICES)
    comment = models.TextField(blank=True, default='')
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['created_at']
        verbose_name = 'Item Request Approval'
        verbose_name_plural = 'Item Request Approvals'

    def __str__(self):
        return f'{self.get_action_display()} — {self.request} by {self.actor}'


class ItemProcurementTimeline(models.Model):
    request = models.OneToOneField(ItemRequest, on_delete=models.CASCADE, related_name='procurement')

    sku = models.ForeignKey(
        'supply_chain.RawMaterialSku',
        on_delete=models.PROTECT,
        null=True,
        blank=True,
        related_name='procurement_timelines',
        verbose_name='Linked SKU',
        help_text='Link to an existing SKU. Leave blank while a new code is being opened.',
    )
    item_code = models.CharField(max_length=50, blank=True, default='')
    code_opened_date = models.DateField(null=True, blank=True)
    sku_pre_existing = models.BooleanField(
        default=False,
        verbose_name='SKU already existed',
        help_text='True when the request was linked to an existing SKU, so code opening was not required.',
    )
    indent_pr_no = models.CharField(max_length=50, blank=True, default='')
    pr_date = models.DateField(null=True, blank=True)
    po_no = models.CharField(max_length=50, blank=True, default='')
    po_date = models.DateField(null=True, blank=True)
    supplier = models.CharField(max_length=150, blank=True, default='')
    received_date = models.DateField(null=True, blank=True)
    unit_price = models.DecimalField(max_digits=12, decimal_places=2, null=True, blank=True)
    received_qty = models.DecimalField(max_digits=12, decimal_places=2, null=True, blank=True)
    remarks = models.TextField(blank=True, default='')

    updated_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='item_procurement_updates')
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = 'Item Procurement Timeline'
        verbose_name_plural = 'Item Procurement Timelines'

    def __str__(self):
        return f'Procurement — {self.request}'

    @property
    def code_opening_required(self):
        """Code opening only applies when the item was not already an SKU."""
        return not self.sku_pre_existing

    @property
    def code_opening_days(self):
        """Days taken to open the code, or days elapsed so far if still open.

        Returns None when the SKU already existed (nothing to track).
        """
        if self.sku_pre_existing:
            return None
        start = self.request.request_date
        if not start:
            return None
        end = self.code_opened_date or timezone.localdate()
        return (end - start).days


class ItemRequestQuote(models.Model):
    """Supplier quote attached to a procurement timeline to justify PO price."""

    procurement = models.ForeignKey(ItemProcurementTimeline, on_delete=models.CASCADE, related_name='quotes')
    supplier = models.CharField(max_length=150)
    quoted_price = models.DecimalField(max_digits=12, decimal_places=2, null=True, blank=True)
    file = models.FileField(upload_to='item_requests/quotes/%Y/%m/')
    notes = models.CharField(max_length=255, blank=True, default='')
    uploaded_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True, related_name='item_request_quotes')
    uploaded_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['-uploaded_at']
        verbose_name = 'Item Request Quote'
        verbose_name_plural = 'Item Request Quotes'

    def __str__(self):
        return f'{self.supplier} quote — {self.procurement.request}'

