from django.db import models
from django.core.exceptions import ValidationError
from django.db.models import Sum


# =========================
# MASTER TABLES
# =========================

class Machine(models.Model):
    name = models.CharField(max_length=100, unique=True)

    class Meta:
        ordering = ['name']

    def __str__(self):
        return self.name


class Department(models.Model):
    name = models.CharField(max_length=100, unique=True)

    class Meta:
        ordering = ['name']

    def __str__(self):
        return self.name


class Material(models.Model):
    name = models.CharField(max_length=100, unique=True)

    class Meta:
        ordering = ['name']

    def __str__(self):
        return self.name


# =========================
# JOB CARD
# =========================

class JobCard(models.Model):
    job_card_no = models.CharField("JC", max_length=50)

    month = models.CharField(max_length=20, null=True, blank=True)

    po_date = models.DateField(null=True, blank=True)
    PO_No = models.IntegerField(null=True, blank=True)

    SKU = models.CharField(max_length=100)

    material = models.ForeignKey(Material, on_delete=models.SET_NULL, null=True, blank=True)

    # You said colour is integer (kept as is)
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

    machine_name = models.ForeignKey(Machine, on_delete=models.SET_NULL, null=True, blank=True)

    department = models.ForeignKey(Department, on_delete=models.SET_NULL, null=True, blank=True)

    die_cutting = models.CharField(max_length=100, null=True, blank=True)

    is_active = models.BooleanField(default=True)

    created_at = models.DateTimeField(auto_now_add=True)

    status = models.CharField(max_length=20, default='Open')

    def __str__(self):
        return self.job_card_no

    # =========================
    # PROPERTIES (ERP LOGIC)
    # =========================

    @property
    def total_production(self):
        return sum(p.output_qty for p in self.production_set.all())

    @property
    def total_dispatch(self):
        return sum(d.dispatch_qty for d in self.dispatch_set.all())

    @property
    def balance_qty(self):
        return self.order_qty - self.total_dispatch

    @property
    def total_waste(self):
        return sum(p.waste_qty for p in self.production_set.all())

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

class Production(models.Model):
    job_card = models.ForeignKey(JobCard, on_delete=models.CASCADE)

    date = models.DateField()
    shift = models.CharField(max_length=20)

    machine = models.CharField(max_length=50)

    output_qty = models.IntegerField(default=0)
    waste_qty = models.IntegerField(default=0)

    def clean(self):
        if not self.job_card:
         return

        existing = Production.objects.filter(
            job_card=self.job_card
    ).exclude(id=self.id).aggregate(
        total=Sum('output_qty')
    )['total'] or 0

        if existing + (self.output_qty or 0) > self.job_card.order_qty:
         raise ValidationError("Production exceeds order quantity!")


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

        existing_dispatch = Dispatch.objects.filter(
          job_card=self.job_card
         ).exclude(id=self.id).aggregate(
         total=Sum('dispatch_qty')
              )['total'] or 0

        total_after_save = existing_dispatch + (self.dispatch_qty or 0)

        total_production = self.job_card.production_set.aggregate(
            total=Sum('output_qty')
                 )['total'] or 0

        if total_after_save > total_production:
         raise ValidationError("Dispatch exceeds production!")

    def __str__(self):
        return f"{self.job_card.job_card_no} - {self.dispatch_date}"