from django import template

register = template.Library()


@register.filter
def get_item(mapping, key):
    """Look up `key` in a dict-like `mapping` from a template (dot lookup can't use a loop variable as the key)."""
    try:
        return mapping.get(key)
    except AttributeError:
        return None


@register.filter
def cap99(value):
    """Display cap for counters that can legitimately run into the thousands
    (e.g. a high-frequency notification rule with no cleanup) — matches the
    "99+" cap already used on the navbar bell badge, so the same count
    doesn't display two different ways on one page."""
    try:
        n = int(value)
    except (TypeError, ValueError):
        return value
    return '99+' if n > 99 else n
