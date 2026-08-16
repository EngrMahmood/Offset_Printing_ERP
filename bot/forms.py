import json

from django import forms

from bot.models import PERIOD_CUSTOM, WEEKDAY_CHOICES, BotAutomation
from bot.report_adapter import available_report_choices


class BotAutomationForm(forms.ModelForm):
    # Drives the layout of bot_form.html. Fields listed in WIDE_FIELDS span the
    # full grid width; everything else sits in the auto-fit columns.
    FIELD_GROUPS = [
        ('Identity', ['code', 'name', 'is_active', 'description']),
        ('Report', [
            'report_slug', 'run_as',
            'report_period', 'report_date_from', 'report_date_to',
            'report_filters',
        ]),
        ('Schedule', ['frequency', 'send_time', 'day_of_month', 'start_date', 'end_date', 'weekdays']),
        ('Recipients', ['email_to', 'email_cc', 'email_bcc', 'recipient_roles']),
        ('Email Draft', ['subject_template', 'body_template']),
        ('Delivery Options', [
            'attach_report', 'attachment_format', 'send_when_empty', 'use_ai_summary',
            'max_rows_in_body', 'retry_count', 'retry_interval_minutes',
        ]),
    ]
    WIDE_FIELDS = {
        'name', 'description', 'report_filters', 'weekdays', 'recipient_roles',
        'subject_template', 'body_template',
    }
    CHECKBOX_LIST_FIELDS = {'weekdays', 'recipient_roles'}

    def grouped_fields(self):
        """[(title, [{'field': BoundField, 'wide': bool, 'checklist': bool}, ...]), ...]"""
        groups = []
        for title, names in self.FIELD_GROUPS:
            items = [
                {
                    'field': self[name],
                    'wide': name in self.WIDE_FIELDS,
                    'checklist': name in self.CHECKBOX_LIST_FIELDS,
                }
                for name in names
                if name in self.fields
            ]
            if items:
                groups.append({'title': title, 'items': items})
        return groups

    # Rendered as checkboxes but stored as a CSV string on the model, so a bot
    # row stays readable in the admin and in exports.
    weekdays = forms.MultipleChoiceField(
        choices=WEEKDAY_CHOICES,
        required=False,
        widget=forms.CheckboxSelectMultiple,
        label='Run on days',
        help_text='Leave all unchecked to run every day (Daily) or on Monday (Weekly).',
    )
    recipient_roles = forms.MultipleChoiceField(
        choices=(),
        required=False,
        widget=forms.CheckboxSelectMultiple,
        label='Also send to everyone with these roles',
        help_text='Resolved to live user emails at send time, so joiners/leavers need no bot edit.',
    )
    report_filters = forms.CharField(
        required=False,
        widget=forms.Textarea(attrs={'rows': 3, 'spellcheck': 'false'}),
        label='Report filters (JSON)',
        help_text='Escape hatch for any other filter the report accepts, '
                  'e.g. {"stage": "not_released"}. Leave blank for the report defaults. '
                  'The Period above wins over a "period" key typed in here.',
    )

    class Meta:
        model = BotAutomation
        fields = [
            'code', 'name', 'description', 'is_active',
            'report_slug', 'report_period', 'report_date_from', 'report_date_to', 'report_filters',
            'frequency', 'send_time', 'weekdays', 'day_of_month', 'start_date', 'end_date',
            'email_to', 'email_cc', 'email_bcc', 'recipient_roles',
            'subject_template', 'body_template', 'max_rows_in_body',
            'attach_report', 'attachment_format', 'send_when_empty', 'use_ai_summary',
            'retry_count', 'retry_interval_minutes', 'run_as',
        ]
        widgets = {
            'description': forms.Textarea(attrs={'rows': 2}),
            'send_time': forms.TimeInput(attrs={'type': 'time'}, format='%H:%M'),
            'start_date': forms.DateInput(attrs={'type': 'date'}, format='%Y-%m-%d'),
            'end_date': forms.DateInput(attrs={'type': 'date'}, format='%Y-%m-%d'),
            'report_date_from': forms.DateInput(attrs={'type': 'date'}, format='%Y-%m-%d'),
            'report_date_to': forms.DateInput(attrs={'type': 'date'}, format='%Y-%m-%d'),
            'email_to': forms.Textarea(attrs={'rows': 2}),
            'email_cc': forms.Textarea(attrs={'rows': 2}),
            'email_bcc': forms.Textarea(attrs={'rows': 2}),
            'body_template': forms.Textarea(attrs={'rows': 14, 'spellcheck': 'false'}),
        }

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        # Report list comes from the live registry — never hard-coded.
        self.fields['report_slug'] = forms.ChoiceField(
            choices=available_report_choices(),
            label='Report',
            help_text='Which registered report this bot emails.',
        )

        from core.models import UserProfile
        self.fields['recipient_roles'].choices = UserProfile.ROLE_CHOICES

        instance = kwargs.get('instance') or self.instance
        if instance and instance.pk:
            # These must land in self.initial, not field.initial. ModelForm has
            # already filled self.initial from the instance via model_to_dict,
            # and BoundField.value() reads self.initial first — so field.initial
            # was never reaching the widgets. The stored values are the wrong
            # shape for these three widgets: report_filters rendered as Python's
            # dict repr (single quotes, rejected by clean_report_filters as
            # "Not valid JSON"), and weekdays/recipient_roles rendered as one CSV
            # string that matched no checkbox, so reopening a bot and saving it
            # silently cleared them.
            self.initial['weekdays'] = [str(n) for n in sorted(instance.weekday_numbers)]
            self.initial['recipient_roles'] = instance.role_keys
            self.initial['report_filters'] = json.dumps(instance.report_filters or {}, indent=2)
            self.initial['report_slug'] = instance.report_slug

        for name, field in self.fields.items():
            widget = field.widget
            if isinstance(widget, (forms.CheckboxInput, forms.CheckboxSelectMultiple)):
                continue
            existing = widget.attrs.get('class', '')
            widget.attrs['class'] = f'{existing} erp-input'.strip()

    def clean_report_filters(self):
        raw = (self.cleaned_data.get('report_filters') or '').strip()
        if not raw:
            return {}
        try:
            value = json.loads(raw)
        except ValueError as exc:
            raise forms.ValidationError(f'Not valid JSON: {exc}') from exc
        if not isinstance(value, dict):
            raise forms.ValidationError('Filters must be a JSON object, e.g. {"stage": "not_released"}.')
        return value

    def clean_weekdays(self):
        return ','.join(self.cleaned_data.get('weekdays') or [])

    def clean_recipient_roles(self):
        return ','.join(self.cleaned_data.get('recipient_roles') or [])

    def clean_day_of_month(self):
        value = self.cleaned_data.get('day_of_month')
        if value is not None and not (1 <= value <= 31):
            raise forms.ValidationError('Day of month must be between 1 and 31.')
        return value

    def clean(self):
        cleaned = super().clean()

        start_date = cleaned.get('start_date')
        end_date = cleaned.get('end_date')
        if start_date and end_date and end_date < start_date:
            self.add_error('end_date', 'End date cannot be before the start date.')

        period = cleaned.get('report_period')
        date_from = cleaned.get('report_date_from')
        date_to = cleaned.get('report_date_to')
        if period == PERIOD_CUSTOM:
            if not date_from:
                self.add_error('report_date_from', 'A custom range needs a start date.')
            if not date_to:
                self.add_error('report_date_to', 'A custom range needs an end date.')
        if date_from and date_to and date_to < date_from:
            self.add_error('report_date_to', 'Range end cannot be before the range start.')

        if cleaned.get('frequency') == 'MONTHLY' and not cleaned.get('day_of_month'):
            self.add_error('day_of_month', 'Monthly bots need a day of month.')

        has_explicit = bool((cleaned.get('email_to') or '').strip())
        has_roles = bool(cleaned.get('recipient_roles'))
        if cleaned.get('is_active') and not (has_explicit or has_roles):
            self.add_error(
                'email_to',
                'An active bot needs at least one recipient — an address here or a recipient role.',
            )

        # Turning on "Use AI Summary" without also editing the template to
        # display {{ai_summary}} produces a bot that silently generates a
        # summary nobody ever sees — this has bitten real bots more than
        # once. Auto-insert the block (right above the report table, same
        # placement used everywhere else) the first time the checkbox is
        # turned on for a template that doesn't already reference it, so the
        # two are never a two-step, easy-to-forget process.
        body_template = cleaned.get('body_template') or ''
        if cleaned.get('use_ai_summary') and body_template and '{{ai_summary}}' not in body_template:
            insert = '{% if ai_summary %}<p>{{ai_summary}}</p>{% endif %}\n\n'
            if '{{report_table}}' in body_template:
                body_template = body_template.replace('{{report_table}}', insert + '{{report_table}}', 1)
            else:
                body_template = insert + body_template
            cleaned['body_template'] = body_template

        # Fail at save time rather than at 08:30 on a Monday morning.
        from bot.template_engine import render_template
        from django.template import TemplateSyntaxError

        probe = {
            'today': '', 'date': '', 'time': '', 'report_title': '', 'report_table': '',
            'total_records': 0, 'bot_name': '', 'user_name': '', 'department': '',
            'filters_summary': '', 'period_label': '', 'period_from': '', 'period_to': '',
            'ai_summary': '',
        }
        for field in ('subject_template', 'body_template'):
            value = cleaned.get(field)
            if not value:
                continue
            try:
                render_template(value, probe)
            except TemplateSyntaxError as exc:
                self.add_error(field, f'Template error: {exc}')

        return cleaned


class TestSendForm(forms.Form):
    email = forms.EmailField(
        label='Send a test copy to',
        widget=forms.EmailInput(attrs={'class': 'erp-input', 'placeholder': 'you@example.com'}),
    )
