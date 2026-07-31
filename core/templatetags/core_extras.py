from django import template

register = template.Library()


@register.filter
def get_item(mapping, key):
    """Look up `key` in a dict-like `mapping` from a template (dot lookup can't use a loop variable as the key)."""
    try:
        return mapping.get(key)
    except AttributeError:
        return None
