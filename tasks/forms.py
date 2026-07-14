from django import forms
from django.contrib.auth.models import User
from .models import Task, Team

class TaskForm(forms.ModelForm):
    due_date = forms.DateField(widget=forms.DateInput(attrs={'type': 'date', 'class': 'erp-input'}))

    class Meta:
        model = Task
        fields = ['title', 'description', 'assignee', 'assigned_team', 'priority', 'due_date']
        widgets = {
            'title': forms.TextInput(attrs={'class': 'erp-input', 'placeholder': 'Task Title'}),
            'description': forms.Textarea(attrs={'class': 'erp-input', 'rows': 4, 'placeholder': 'Task Description'}),
            'assignee': forms.Select(attrs={'class': 'erp-select'}),
            'assigned_team': forms.Select(attrs={'class': 'erp-select'}),
            'priority': forms.Select(attrs={'class': 'erp-select'}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.fields['assignee'].empty_label = "Select Individual (Optional)"
        self.fields['assigned_team'].empty_label = "Select Team (Optional)"
        # Fetch active users for assignment
        self.fields['assignee'].queryset = User.objects.filter(is_active=True).order_by('username')

    def clean(self):
        cleaned_data = super().clean()
        assignee = cleaned_data.get('assignee')
        assigned_team = cleaned_data.get('assigned_team')

        if not assignee and not assigned_team:
            raise forms.ValidationError("You must assign the task to either an individual employee or a team.")
        return cleaned_data


class TeamForm(forms.ModelForm):
    members = forms.ModelMultipleChoiceField(
        queryset=User.objects.filter(is_active=True).order_by('username'),
        widget=forms.CheckboxSelectMultiple,
        required=False
    )

    class Meta:
        model = Team
        fields = ['name', 'description', 'members']
        widgets = {
            'name': forms.TextInput(attrs={'class': 'erp-input', 'placeholder': 'Team Name'}),
            'description': forms.Textarea(attrs={'class': 'erp-input', 'rows': 3, 'placeholder': 'Team Description'}),
        }
