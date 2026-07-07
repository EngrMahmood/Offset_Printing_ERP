from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal

from django.db.models import QuerySet


def to_json_safe(value):
    if value is None:
        return None
    if isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, Decimal):
        return float(value)
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    if isinstance(value, dict):
        return {str(k): to_json_safe(v) for k, v in value.items()}
    if isinstance(value, (list, tuple, set)):
        return [to_json_safe(item) for item in value]
    if isinstance(value, QuerySet):
        return [to_json_safe(item) for item in list(value)]
    if hasattr(value, 'pk'):
        return {
            'id': getattr(value, 'pk', None),
            'label': str(value),
        }
    return str(value)
