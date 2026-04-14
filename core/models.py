from django.db import models, transaction, IntegrityError
from django.core.exceptions import ValidationError
from django.db.models import Sum
from django.contrib.auth import get_user_model


# =========================
# MASTER TABLES
# =========================

class Machine(models.Model):
    name = models.CharField(max_length=100, unique=True)
    standard_speed = models.FloatField(default=4000)
    is_active = models.BooleanField(default=True)

    def __str__(self):
        return self.name


class Department(models.Model):
    name = models.CharField(max_length=100, unique=True)

    def __str__(self):
        return self.name


class Material(models.Model):
    name = models.CharField(max_length=100, unique=True)

    def __str__(self):
        return self.name


class Operator(models.Model):
    name = models.CharField(max_length=100)
    employee_code = models.CharField(max_length=50, null=True, blank=True)
    is_active = models.BooleanField(default=True)

    def __str__(self):
        return self.name

# =========================
# JOB CARD
# =========================

class JobCard(models.Model):
    job_card_no = models.CharField(max_length=50)

    month = models.CharField(max_length=20, null=True, blank=True)
    po_date = models.DateField(null=True, blank=True)
    PO_No = models.CharField(max_length=50, null=True, blank=True)

    SKU = models.CharField(max_length=100)

    material = models.ForeignKey(Material, on_delete=models.SET_NULL, null=True, blank=True)

    colour = models.IntegerField(null=True, blank=True)
    application = models.CharField(max_length=100, null=True, blank=True)

    order_qty = models.IntegerField()

    ups = models.IntegerField(null=True, blank=True)
    print_sheet_size = models.CharField(max_length=50, null=True, blank=True)

    wastage = models.IntegerField(default=0)

    actual_sheet_required = models.IntegerField(null=True, blank=True)

    purchase_sheet_size = models.CharField(max_length=50, null=True, blank=True)
    purchase_sheet_ups = models.IntegerField(null=True, blank=True)

    remarks = models.TextField(null=True, blank=True)

    destination = models.CharField(max_length=100, null=True, blank=True)

    machine_name = models.ForeignKey('Machine', on_delete=models.SET_NULL, null=True, blank=True)
    department = models.ForeignKey(Department, on_delete=models.SET_NULL, null=True, blank=True)

    die_cutting = models.CharField(max_length=100, null=True, blank=True)

    is_active = models.BooleanField(default=True)
    created_at = models.DateTimeField(auto_now_add=True)

    status = models.CharField(max_length=20, default='Open')

    def __str__(self):
        return self.job_card_no

    # ===== ERP PROPERTIES =====

    @property
    def total_production(self):
        return self.productions.aggregate(total=Sum('output_qty'))['total'] or 0

    @property
    def total_dispatch(self):
        return self.dispatch_set.aggregate(total=Sum('dispatch_qty'))['total'] or 0

    @property
    def total_waste(self):
        return self.productions.aggregate(total=Sum('waste_qty'))['total'] or 0

    @property
    def balance_qty(self):
        return self.order_qty - self.total_dispatch

    @property
    def waste_percentage(self):
        if self.total_production == 0:
            return 0
        return round((self.total_waste / self.total_production) * 100, 2)

    @property
    def job_status(self):
        if self.total_dispatch >= self.order_qty:
            return "Completed"
        elif self.total_production > 0:
            return "In Progress"
        return "Open"


# =========================
# PRODUCTION
# =========================

User = get_user_model()


class Production(models.Model):

    SHIFT_CHOICES = [
        ('A', 'Shift A'),
        ('B', 'Shift B'),
        
            ]

    job_card = models.ForeignKey('JobCard', on_delete=models.CASCADE, related_name='productions')

    date = models.DateField()
    shift = models.CharField(max_length=10, choices=SHIFT_CHOICES)

    machine = models.ForeignKey('Machine', on_delete=models.PROTECT)

    output_qty = models.PositiveIntegerField()
    waste_qty = models.PositiveIntegerField(default=0)

    planned_time = models.FloatField()
    run_time = models.FloatField()
    downtime = models.FloatField(default=0)
    setup_time = models.FloatField(default=0)

    ideal_run_rate = models.FloatField(null=True, blank=True)

    operator = models.ForeignKey(
        'Operator',
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        limit_choices_to={'is_active': True}
        
    )

    created_at = models.DateTimeField(auto_now_add=True)

    def clean(self):
        errors = {}

        existing = Production.objects.filter(job_card=self.job_card)\
            .exclude(id=self.id)\
            .aggregate(total=Sum('output_qty'))['total'] or 0

        if existing + self.output_qty > self.job_card.order_qty:
            errors['output_qty'] = "Production exceeds order quantity!"

        if self.run_time and self.planned_time:
            if self.run_time > self.planned_time:
                errors['run_time'] = "Run time cannot exceed planned time."

        total_time = (self.run_time or 0) + (self.downtime or 0) + (self.setup_time or 0)
        if self.planned_time and total_time > self.planned_time:
            errors['planned_time'] = "Total time exceeds planned time."

        if self.ideal_run_rate is not None and self.ideal_run_rate <= 0:
            errors['ideal_run_rate'] = "Must be > 0"

        if errors:
            raise ValidationError(errors)

    def save(self, *args, **kwargs):
        if not self.ideal_run_rate and self.machine:
            self.ideal_run_rate = self.machine.standard_speed

        self.full_clean()

        try:
            with transaction.atomic():
                super().save(*args, **kwargs)
        except IntegrityError:
            raise ValidationError("DB error while saving Production")

    # ===== OEE =====

    def availability(self):
        return self.run_time / self.planned_time if self.planned_time else 0

    def performance(self):
        if self.run_time and self.ideal_run_rate:
            return self.output_qty / (self.ideal_run_rate * self.run_time)
        return 0

    def quality(self):
        total = self.output_qty + self.waste_qty
        return self.output_qty / total if total else 0

    def oee(self):
        return self.availability() * self.performance() * self.quality()

    def operator_efficiency(self):
        if self.run_time and self.ideal_run_rate:
            return (self.output_qty / (self.ideal_run_rate * self.run_time)) * 100
        return 0

    def __str__(self):
        return f"{self.job_card.job_card_no} - {self.date}"


# =========================
# DISPATCH
# =========================

class Dispatch(models.Model):

    job_card = models.ForeignKey(JobCard, on_delete=models.CASCADE)
    dispatch_date = models.DateField()
    dispatch_qty = models.IntegerField(default=0)

    def clean(self):
        if not self.job_card:
            return

        existing_dispatch = Dispatch.objects.filter(job_card=self.job_card)\
            .exclude(id=self.id)\
            .aggregate(total=Sum('dispatch_qty'))['total'] or 0

        total_after = existing_dispatch + self.dispatch_qty

        total_production = self.job_card.productions.aggregate(
            total=Sum('output_qty')
        )['total'] or 0

        if total_after > total_production:
            raise ValidationError("Dispatch exceeds production!")

    def __str__(self):
        return f"{self.job_card.job_card_no} - {self.dispatch_date}"