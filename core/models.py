from django.db import models
from django.core.exceptions import ValidationError

class JobCard(models.Model):
    job_card_no = models.CharField(max_length=50)
    SKU = models.CharField(max_length=100)
    PO_No = models.IntegerField(null=True,blank=True)
    order_qty = models.IntegerField()
    created_at = models.DateTimeField(auto_now_add=True)
    status = models.CharField(max_length=20, default='Open')

    def __str__(self):
        return self.job_card_no
    
    @property
    def total_waste(self):
        return sum(p.waste_qty for p in self.production_set.all())

    @property
    def waste_percentage(self):
        if self.total_production == 0:
            return 0
        return round ((self.total_waste / self.total_production) * 100, 2)



    @property
    def job_status(self):
        if self.total_dispatch >= self.order_qty:
         return "Completed"
        elif self.total_production > 0:
         return "In Progress"
        return "Open"

    @property
    def total_production(self):
        return sum(p.output_qty for p in self.production_set.all())

    @property
    def total_dispatch(self):
        return sum(d.dispatch_qty for d in self.dispatch_set.all())

    @property
    def balance_qty(self):
        return self.order_qty - self.total_dispatch


class Production(models.Model):
    job_card = models.ForeignKey(JobCard, on_delete=models.CASCADE)
    date = models.DateField()
    shift = models.CharField(max_length=20)
    machine = models.CharField(max_length=50)
    output_qty = models.IntegerField()
    waste_qty = models.IntegerField(default=0)

    def clean(self):
        total_production = sum(p.output_qty for p in self.job_card.production_set.all())
        if total_production + self.output_qty > self.job_card.order_qty:
            raise ValidationError("Production exceeds order quantity!")

    def __str__(self):
        return f"{self.job_card} - {self.date}"


class Dispatch(models.Model):
    job_card = models.ForeignKey(JobCard, on_delete=models.CASCADE)
    dispatch_date = models.DateField()
    dispatch_qty = models.IntegerField()


    def clean(self):
        total_dispatch = sum(d.dispatch_qty for d in self.job_card.dispatch_set.all())
        if total_dispatch + self.dispatch_qty > self.job_card.total_production:
            raise ValidationError("Dispatch exceeds production!")


    def __str__(self):
        return f"{self.job_card} - {self.dispatch_date}"
    
class Meta:
    ordering = ['-date']