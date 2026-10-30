from django import forms

from production_plan.models import PlanSettings


class PlanSettingsForm(forms.ModelForm):
    class Meta:
        model = PlanSettings
        exclude = ['updated_at']
        widgets = {
            'draft_time': forms.TimeInput(attrs={'type': 'time'}),
            'rebalance_time': forms.TimeInput(attrs={'type': 'time'}),
            'email_time': forms.TimeInput(attrs={'type': 'time'}),
            'off_dates': forms.Textarea(attrs={'rows': 3}),
            'email_to_extra': forms.Textarea(attrs={'rows': 2}),
            'email_cc': forms.Textarea(attrs={'rows': 2}),
        }
