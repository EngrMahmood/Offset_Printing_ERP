from django import template

register = template.Library()


@register.filter
def get_item(value, key):
    if isinstance(value, dict):
        return value.get(key, '')
    try:
        return getattr(value, key, '')
    except Exception:
        return ''
