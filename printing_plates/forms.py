from django import forms

from .models import PlateRequest


class PlateRequestForm(forms.ModelForm):
    class Meta:
        model = PlateRequest
        fields = [
            'planning_job',
            'job_card',
            'sku_recipe',
            'machine',
            'department',
            'set_no',
            'new_set_no',
            'awc_no',
            'vendor',
            'plate_quantity',
            'status',
            'plate_color',
            'impression',
            'remarks',
            'requested_at',
            'progress',
            'image',
            'link',
        ]
        labels = {
            'planning_job': 'Planning Job',
            'job_card': 'JC#',
            'sku_recipe': 'SKU',
            'machine': 'Machine Name',
            'department': 'Department',
            'set_no': 'Set #',
            'new_set_no': 'New Set #',
            'awc_no': 'AWC #',
            'vendor': 'Vendor',
            'plate_quantity': 'Plate Quantity',
            'status': 'Status',
            'plate_color': 'Plate Inks',
            'impression': 'Impression',
            'remarks': 'Remarks',
            'requested_at': 'Request Date',
            'progress': 'Progress',
            'link': 'Link',
            'image': 'Image',
        }
        widgets = {
            'awc_no': forms.TextInput(attrs={'inputmode': 'text', 'autocomplete': 'off'}),
            'remarks': forms.Textarea(attrs={'rows': 4}),
            'requested_at': forms.DateTimeInput(attrs={'type': 'datetime-local'}),
            'link': forms.URLInput(attrs={'placeholder': 'https://'}),
        }

    def clean_awc_no(self):
        from planning.services import normalize_awc_no

        return normalize_awc_no(self.cleaned_data.get('awc_no'))
