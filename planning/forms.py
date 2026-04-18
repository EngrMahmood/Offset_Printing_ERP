from django import forms

from .models import PlanningJob, SkuRecipe


class PlanningJobEditForm(forms.ModelForm):
    class Meta:
        model = PlanningJob
        fields = [
            'plan_date',
            'po_number',
            'sku',
            'job_name',
            'material',
            'color_spec',
            'application',
            'order_qty',
            'print_sheets',
            'machine_name',
            'department',
            'destination',
            'unit_cost',
            'daily_demand',
            'remarks',
            'requirement',
            'status',
        ]
        widgets = {
            'plan_date': forms.DateInput(attrs={'type': 'date'}),
            'remarks': forms.Textarea(attrs={'rows': 3}),
            'requirement': forms.Textarea(attrs={'rows': 3}),
        }


class SkuRecipeForm(forms.ModelForm):
    class Meta:
        model = SkuRecipe
        fields = [
            'sku',
            'job_name',
            'material',
            'color_spec',
            'application',
            'size_w_mm',
            'size_h_mm',
            'ups',
            'print_sheet_size',
            'purchase_sheet_size',
            'purchase_material',
            'machine_name',
            'department',
            'default_unit_cost',
            'notes',
        ]
        widgets = {
            'notes': forms.Textarea(attrs={'rows': 3}),
        }
