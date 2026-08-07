from django.conf import settings
from django.db import models


class FaultCategory(models.Model):
    name = models.CharField(max_length=50, unique=True)
    is_active = models.BooleanField(default=True, db_index=True)

    class Meta:
        ordering = ['name']

    def __str__(self):
        return self.name


class MaintenanceRecord(models.Model):
    MAINTENANCE_TYPE_CHOICES = [
        ('PREVENTIVE', 'Preventive'),
        ('CORRECTIVE', 'Corrective'),
        ('BREAKDOWN', 'Breakdown'),
        ('PREDICTIVE', 'Predictive'),
    ]
    EXECUTION_TYPE_CHOICES = [
        ('IN_HOUSE', 'In-house'),
        ('OUTSOURCE', 'Outsource'),
    ]
    PRIORITY_CHOICES = [
        ('LOW', 'Low'),
        ('MEDIUM', 'Medium'),
        ('MAJOR', 'Major'),
        ('CRITICAL', 'Critical'),
    ]
    STATUS_CHOICES = [
        ('PENDING_APPROVAL', 'Pending Approval'),
        ('REPORTED', 'Reported'),
        ('DIAGNOSED', 'Diagnosed'),
        ('AWAITING_PARTS', 'Awaiting Parts'),
        ('AWAITING_VENDOR', 'Awaiting Vendor'),
        ('IN_PROGRESS', 'In Progress'),
        ('COMPLETED', 'Completed'),
        ('VERIFIED', 'Verified'),
        ('CLOSED', 'Closed'),
        ('REJECTED', 'Rejected'),
        ('CANCELLED', 'Cancelled'),
    ]
    OPEN_STATUSES = {
        'PENDING_APPROVAL', 'REPORTED', 'DIAGNOSED', 'AWAITING_PARTS', 'AWAITING_VENDOR', 'IN_PROGRESS', 'COMPLETED',
    }

    record_no = models.CharField(max_length=30, unique=True, blank=True, null=True, verbose_name='Record No.')
    machine = models.ForeignKey('core.Machine', on_delete=models.PROTECT, related_name='maintenance_records')
    reported_date = models.DateField()
    reported_by = models.ForeignKey(
        settings.AUTH_USER_MODEL, on_delete=models.PROTECT, related_name='maintenance_records_reported',
    )
    fault_types = models.ManyToManyField(FaultCategory, related_name='maintenance_records', blank=True)
    maintenance_type = models.CharField(max_length=20, choices=MAINTENANCE_TYPE_CHOICES, blank=True, default='')
    execution_type = models.CharField(max_length=20, choices=EXECUTION_TYPE_CHOICES, default='IN_HOUSE')
    priority = models.CharField(max_length=20, choices=PRIORITY_CHOICES, default='MEDIUM')
    fault_description = models.TextField()
    proposed_solution = models.TextField(blank=True)
    spare_parts_needed = models.BooleanField(default=False)
    repair_needed = models.BooleanField(default=False)
    repair_details = models.TextField(blank=True)
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='PENDING_APPROVAL', db_index=True)
    assigned_to = models.ForeignKey(
        settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, blank=True,
        related_name='maintenance_records_assigned',
    )
    work_start_at = models.DateTimeField(null=True, blank=True)
    work_end_at = models.DateTimeField(null=True, blank=True)
    labour_hours = models.DecimalField(max_digits=8, decimal_places=2, default=0)
    remarks = models.TextField(blank=True)
    is_active = models.BooleanField(default=True, db_index=True)
    deleted_by = models.ForeignKey(
        settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True, blank=True,
        related_name='maintenance_records_deleted',
    )
    deleted_at = models.DateTimeField(null=True, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ['-reported_date', '-id']

    def __str__(self):
        return self.record_no or f'Maintenance #{self.pk}'

    @property
    def spares_cost(self):
        total = 0
        for line in self.spare_parts.select_related('item_request__procurement').all():
            procurement = getattr(line.item_request, 'procurement', None) if line.item_request_id else None
            if procurement and procurement.unit_price and procurement.received_qty:
                total += procurement.unit_price * procurement.received_qty
        return total

    @property
    def service_cost(self):
        total = 0
        for job in self.service_jobs.select_related('item_request__procurement').all():
            procurement = getattr(job.item_request, 'procurement', None) if job.item_request_id else None
            if procurement and procurement.unit_price and procurement.received_qty:
                total += procurement.unit_price * procurement.received_qty
        return total

    @property
    def labour_cost(self):
        from django.conf import settings as django_settings

        rate = getattr(django_settings, 'MAINTENANCE_LABOUR_HOURLY_RATE', 500)
        return (self.labour_hours or 0) * rate

    @property
    def total_cost(self):
        return self.spares_cost + self.service_cost + self.labour_cost


class MaintenanceSparePart(models.Model):
    record = models.ForeignKey(MaintenanceRecord, on_delete=models.CASCADE, related_name='spare_parts')
    description = models.CharField(max_length=255)
    quantity = models.DecimalField(max_digits=12, decimal_places=2)
    uom = models.CharField(max_length=30, blank=True)
    existing_sku = models.ForeignKey(
        'supply_chain.RawMaterialSku', on_delete=models.SET_NULL, null=True, blank=True,
        related_name='maintenance_spare_lines',
    )
    item_request = models.ForeignKey(
        'supply_chain.ItemRequest', on_delete=models.SET_NULL, null=True, blank=True,
        related_name='maintenance_spare_lines',
    )
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f'{self.description} x{self.quantity}'


class MaintenanceServiceJob(models.Model):
    record = models.ForeignKey(MaintenanceRecord, on_delete=models.CASCADE, related_name='service_jobs')
    vendor = models.CharField(max_length=150, blank=True)
    scope = models.TextField()
    item_request = models.ForeignKey(
        'supply_chain.ItemRequest', on_delete=models.SET_NULL, null=True, blank=True,
        related_name='maintenance_service_jobs',
    )
    sent_out_date = models.DateField(null=True, blank=True)
    returned_date = models.DateField(null=True, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f'Service job for {self.record}'


class MachineDowntime(models.Model):
    REASON_CHOICES = [
        ('BREAKDOWN', 'Breakdown'),
        ('PLANNED_PM', 'Planned PM'),
        ('AWAITING_PART', 'Awaiting Part'),
        ('AWAITING_VENDOR', 'Awaiting Vendor'),
        ('OTHER', 'Other'),
    ]

    machine = models.ForeignKey('core.Machine', on_delete=models.PROTECT, related_name='downtime_intervals')
    record = models.ForeignKey(
        MaintenanceRecord, on_delete=models.SET_NULL, null=True, blank=True, related_name='downtime_intervals',
    )
    start_at = models.DateTimeField()
    end_at = models.DateTimeField(null=True, blank=True)
    reason = models.CharField(max_length=20, choices=REASON_CHOICES, default='BREAKDOWN')
    scheduled_minutes_lost = models.PositiveIntegerField(default=0)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['-start_at']

    def __str__(self):
        return f'{self.machine} down from {self.start_at}'


class MaintenanceApproval(models.Model):
    ACTION_CHOICES = [
        ('SUBMIT', 'Submitted'),
        ('APPROVE', 'Approved'),
        ('REJECT', 'Rejected'),
        ('DELETE', 'Deleted'),
    ]

    record = models.ForeignKey(MaintenanceRecord, on_delete=models.CASCADE, related_name='approvals')
    actor = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True)
    action = models.CharField(max_length=20, choices=ACTION_CHOICES)
    comment = models.TextField(blank=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['-created_at']

    def __str__(self):
        return f'{self.record} — {self.get_action_display()}'


class PreventiveMaintenancePlan(models.Model):
    INTERVAL_TYPE_CHOICES = [
        ('DAYS', 'Days'),
        ('IMPRESSIONS', 'Impressions'),
    ]

    machine = models.ForeignKey('core.Machine', on_delete=models.CASCADE, related_name='pm_plans')
    title = models.CharField(max_length=150)
    interval_type = models.CharField(max_length=20, choices=INTERVAL_TYPE_CHOICES, default='DAYS')
    interval_value = models.PositiveIntegerField(help_text='Days between services, or impressions between services.')
    last_done_at = models.DateField(null=True, blank=True)
    next_due_at = models.DateField(null=True, blank=True)
    next_due_impressions = models.PositiveIntegerField(null=True, blank=True)
    is_active = models.BooleanField(default=True, db_index=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['machine', 'title']

    def __str__(self):
        return f'{self.title} ({self.machine})'


class MaintenanceAttachment(models.Model):
    record = models.ForeignKey(MaintenanceRecord, on_delete=models.CASCADE, related_name='attachments')
    file = models.FileField(upload_to='maintenance/attachments/%Y/%m/')
    caption = models.CharField(max_length=255, blank=True)
    uploaded_by = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True)
    uploaded_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return self.caption or self.file.name


class MaintenanceActivityLog(models.Model):
    record = models.ForeignKey(MaintenanceRecord, on_delete=models.CASCADE, related_name='activity_log')
    actor = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.SET_NULL, null=True)
    action = models.CharField(max_length=255)
    from_status = models.CharField(max_length=20, blank=True)
    to_status = models.CharField(max_length=20, blank=True)
    note = models.TextField(blank=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ['-created_at']

    def __str__(self):
        return f'{self.record} — {self.action}'
