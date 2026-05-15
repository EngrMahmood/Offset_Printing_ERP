from django import forms

from .models import ComparisonModule, ImportModule


class GoogleSheetUploadForm(forms.Form):
    module = forms.ChoiceField(choices=ImportModule.choices)
    sheet_url = forms.URLField(
        max_length=1024,
        label='Google Sheet URL',
        widget=forms.URLInput(
            attrs={
                'placeholder': 'https://docs.google.com/spreadsheets/d/...',
                'class': 'form-control',
            }
        ),
    )

    def clean_sheet_url(self):
        value = (self.cleaned_data.get('sheet_url') or '').strip()
        if 'docs.google.com/spreadsheets' not in value:
            raise forms.ValidationError('Please provide a valid Google Sheet URL.')
        return value


class GoogleSheetComparisonForm(forms.Form):
    module = forms.ChoiceField(choices=ComparisonModule.choices, label='Target ERP module')
    sheet_url = forms.URLField(
        max_length=1024,
        label='Google Sheet URL',
        widget=forms.URLInput(
            attrs={
                'placeholder': 'https://docs.google.com/spreadsheets/d/...',
                'class': 'form-control',
            }
        ),
    )

    def clean_sheet_url(self):
        value = (self.cleaned_data.get('sheet_url') or '').strip()
        if 'docs.google.com/spreadsheets' not in value:
            raise forms.ValidationError('Please provide a valid Google Sheet URL.')
        return value
