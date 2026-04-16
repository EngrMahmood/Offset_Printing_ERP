from django.db import models, transaction, IntegrityError
from django.core.exceptions import ValidationError
from django.db.models import Sum
from django.contrib.auth import get_user_model


# =========================
# MASTER TABLES
# =========================

class Machine(models.Model):
    name = models.CharField(max_length=100, unique=True)
    standard_impressions_per_hour = models.FloatField(default=4000, help_text="Standard printing speed in impressions per hour")

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
    job_card_no = models.CharField(max_length=50, unique=True)

    month = models.CharField(max_length=20, null=True, blank=True)
    po_date = models.DateField(null=True, blank=True)
    PO_No = models.CharField(max_length=50, null=True, blank=True)

    SKU = models.CharField(max_length=100)

    material = models.ForeignKey(Material, on_delete=models.SET_NULL, null=True, blank=True)

    colour = models.IntegerField(null=True, blank=True)
    application = models.CharField(max_length=100, null=True, blank=True)

    order_qty = models.IntegerField()

    total_impressions_required = models.IntegerField(
        null=True, 
        blank=True,
        help_text="Total impressions required for this job (manually entered based on machine config - 1/2/5 color, front-back, etc.)"
    )

    ups = models.IntegerField(null=True, blank=True)
    print_sheet_size = models.CharField(max_length=50, null=True, blank=True)

    wastage = models.IntegerField(default=0,help_text="in Sheets")

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
    def required_sheets(self):
        if self.ups:
            return self.order_qty / self.ups
        return 0
    
    @property
    def total_sheets_planned(self):
        return int (self.required_sheets + self.wastage)
    
    @property
    def total_production(self):
        return self.productions.aggregate(total=Sum('output_sheets'))['total'] or 0

    @property
    def total_dispatch(self):
        return self.dispatch_set.aggregate(total=Sum('dispatch_qty'))['total'] or 0

    @property
    def total_waste(self):
        return self.productions.aggregate(total=Sum('waste_sheets'))['total'] or 0

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
        if self.order_qty == 0:
            return "Open"

        dispatch_ratio = (self.total_dispatch / self.order_qty) * 100

        if dispatch_ratio >= 95:
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
    shift = models.CharField(max_length=1, choices=SHIFT_CHOICES)

    machine = models.ForeignKey('Machine', on_delete=models.PROTECT)

    output_sheets = models.PositiveIntegerField()
    waste_sheets = models.PositiveIntegerField(default=0)

    WASTE_CHOICES = [
        ('paper_jam', 'Paper Jam (Affects OEE Quality)'),
        ('color_issue', 'Color/Registration Issue (Affects OEE Quality)'),
        ('material_defect', 'Material Defect (External - Excluded from OEE)'),
        ('operator_error', 'Operator Error (Affects OEE Quality)'),
        ('machine_issue', 'Machine Issue (Affects OEE Quality)'),
        ('other', 'Other (Affects OEE Quality)'),
    ]

    waste_reason = models.CharField(
        max_length=20,
        choices=WASTE_CHOICES,
        null=True,
        blank=True,
        help_text="Primary reason for waste"
    )

    impressions = models.PositiveIntegerField(help_text="Total impressions produced (sheets × passes)")

    planned_time = models.FloatField(help_text="in minutes")
    run_time = models.FloatField(help_text="in minutes")
    downtime = models.FloatField(default=0,help_text="in minutes")
    setup_time = models.FloatField(default=0,help_text="in minutes")

    DOWNTIME_CHOICES = [
        ('setup', 'Setup Time (Planned - Excluded from OEE)'),
        ('maintenance', 'Maintenance (Planned - Excluded from OEE)'),
        ('breakdown', 'Machine Breakdown (Unplanned - Affects OEE)'),
        ('material', 'Material Issue (External - Excluded from OEE)'),
        ('operator', 'Operator Issue (Affects OEE)'),
        ('other', 'Other (Affects OEE)'),
    ]

    downtime_category = models.CharField(
        max_length=20,
        choices=DOWNTIME_CHOICES,
        null=True,
        blank=True,
        help_text="Category of downtime"
    )

    ideal_run_rate = models.FloatField(null=True, blank=True)

    operator = models.ForeignKey('Operator',on_delete=models.SET_NULL,null=True,blank=True,limit_choices_to={'is_active': True})

    created_at = models.DateTimeField(auto_now_add=True)

    #Calculated Fields

    @property
    def pcs_produced(self):
        if self.job_card.ups:
            return self.output_sheets * self.job_card.ups
        return 0
    
    @property
    def good_sheets(self):
        return self.output_sheets
    
    @property
    def total_sheets(self):
        return self.output_sheets + self.waste_sheets

    def clean(self):
        errors = {}

        existing = Production.objects.filter(job_card=self.job_card)\
            .exclude(id=self.id)\
            .aggregate(
                total_output=Sum('output_sheets'),
                total_waste=Sum('waste_sheets')
        )

        existing_output = existing['total_output'] or 0
        existing_waste = existing['total_waste'] or 0

        total_existing_consumption = existing_output + existing_waste
        current_consumption = (self.output_sheets or 0) + (self.waste_sheets or 0)

    # 🔴 MAIN VALIDATION (FIXED)
        if total_existing_consumption + current_consumption > self.job_card.total_sheets_planned:
         errors['output_sheets'] = "Total sheets (production + waste) exceed planned sheets!"

        # Impressions validation
        if self.impressions <= 0:
            errors['impressions'] = "Impressions must be greater than 0"
        if self.impressions < self.output_sheets:
            errors['impressions'] = "Impressions should be at least equal to output sheets"

    # ⏱ TIME VALIDATIONS
        if self.run_time and self.planned_time:
         if self.run_time > self.planned_time:
            errors['run_time'] = "Run time cannot exceed planned time."

        total_time = (self.run_time or 0) + (self.downtime or 0) + (self.setup_time or 0)
        if self.planned_time and total_time > self.planned_time:
         errors['planned_time'] = "Total time exceeds planned time."

    # ⚙️ RATE VALIDATION
        if self.ideal_run_rate is not None and self.ideal_run_rate <= 0:
            errors['ideal_run_rate'] = "Must be > 0"

    # 🔥 VERY IMPORTANT (OUTSIDE ALL BLOCKS)
        if errors:
         raise ValidationError(errors)

    def save(self, *args, **kwargs):
        if not self.ideal_run_rate and self.machine:
            self.ideal_run_rate = self.machine.standard_impressions_per_hour

        self.full_clean()

        try:
            with transaction.atomic():
                super().save(*args, **kwargs)
        except IntegrityError:
            raise ValidationError("DB error while saving Production")

    # ===== OEE =====

    @property
    def expected_impressions(self):
        """Expected impressions based on machine capacity and run time"""
        if not self.machine or not self.machine.standard_impressions_per_hour:
            return 0
        run_time_hours = self.run_time / 60
        return self.machine.standard_impressions_per_hour * run_time_hours

    @property
    def availability(self):
        if self.planned_time == 0:
            return 0
        # OEE Availability excludes planned downtime
        # Include only unplanned downtime categories
        unplanned_categories = ['breakdown', 'operator', 'other']
        unplanned_downtime = self.downtime if self.downtime_category in unplanned_categories else 0
        return (self.planned_time - unplanned_downtime) / self.planned_time

    @property
    def performance(self):
        if not self.machine or not self.machine.standard_impressions_per_hour:
            return 0
        
        run_time_hours = self.run_time/60

        expected_impressions = self.machine.standard_impressions_per_hour * run_time_hours
        if expected_impressions == 0:
            return 0

        return self.impressions / expected_impressions
    


    @property
    def quality(self):
        if self.total_sheets == 0:
            return 0
        # OEE Quality excludes wastes not caused by production process
        # Only count wastes that affect machine/process quality
        quality_affecting_wastes = ['paper_jam', 'color_issue', 'operator_error', 'machine_issue', 'other']
        quality_waste = self.waste_sheets if self.waste_reason in quality_affecting_wastes else 0
        good_sheets = self.output_sheets  # assuming output_sheets are good
        total_quality_sheets = good_sheets + quality_waste
        if total_quality_sheets == 0:
            return 0
        return good_sheets / total_quality_sheets
    

    @property
    def oee(self):
        return round((self.availability * self.performance * self.quality),2)
    


    def operator_efficiency(self):
        if self.run_time and self.ideal_run_rate:
            run_time_hours = self.run_time/60
            expected_impressions = self.ideal_run_rate * run_time_hours
            if expected_impressions == 0:
                return 0
            return (self.impressions / expected_impressions) * 100
        return 0

    def __str__(self):
        return f"{self.job_card.job_card_no} - {self.date}"

# ========================= 
# DISPATCH 
# =========================


class Dispatch(models.Model):

    job_card = models.ForeignKey(JobCard, on_delete=models.CASCADE)

    dc_no = models.CharField(
        max_length=50,
        null=True,
        blank=True,
        help_text="Dispatch Challan Number (can be shared across multiple Job Cards)"

    )

    dispatch_date = models.DateField()

    dispatch_qty = models.IntegerField(default=0)


    # =========================
    # VALIDATION ONLY
    # =========================
    def clean(self):
        errors = {}

        if self.dispatch_qty <= 0:
            errors['dispatch_qty'] = "Dispatch must be greater than 0"

        existing_dispatch = Dispatch.objects.filter(job_card=self.job_card)\
            .exclude(id=self.id)\
            .aggregate(total=Sum('dispatch_qty'))['total'] or 0

        total_after = existing_dispatch + (self.dispatch_qty or 0)

        total_production = sum(
            p.pcs_produced for p in self.job_card.productions.all()
        )

        if total_after > total_production:
            errors['dispatch_qty'] = "Dispatch cannot exceed total produced quantity!"

        if errors:
            raise ValidationError(errors)

    # =========================
    # SAVE LOGIC
    # =========================
    def save(self, *args, **kwargs):
        self.full_clean()   # runs clean() + field validation
        super().save(*args, **kwargs)

    def __str__(self):
        return f"{self.job_card.job_card_no} - {self.dispatch_date}"