"""Field-level errors for entry forms.

A view raises FieldError({'operator': 'Select the operator.', ...}) (or a Django ValidationError with a message dict)
and calls remember_field_errors(request, exc) before it re-renders the form. The page then shows what the user typed
(core.context_processors.form_repost + theme/js/form_repost.js) and marks only the fields that need attention in red.

Keys are form field names. For repeated rows ('downtime_category[]') add '#<row number from 0>' to mark one row:
'downtime_category[]#1'.
"""
from django.core.exceptions import ValidationError


class FieldError(ValueError):
    """Validation failed on specific form fields; .errors maps field name -> message."""

    def __init__(self, errors, message=None):
        self.errors = {str(key): str(value) for key, value in dict(errors).items()}
        super().__init__(message or ' '.join(self.errors.values()))


def field_errors_from(exc):
    if isinstance(exc, FieldError):
        return dict(exc.errors)
    if isinstance(exc, ValidationError) and hasattr(exc, 'message_dict'):
        return {key: ' '.join(str(m) for m in value) for key, value in exc.message_dict.items() if key != '__all__'}
    return {}


def remember_field_errors(request, exc):
    """Hand the errors of this failed save to the page that is about to be rendered."""
    request.field_errors = field_errors_from(exc)
    return request.field_errors
