from django import template

register = template.Library()


@register.simple_tag(takes_context=True)
def is_active_app(context, app_name):
    request = context.get('request')
    if not request or not getattr(request, 'resolver_match', None):
        return ''
    return 'is-active' if request.resolver_match.app_name == app_name else ''


@register.simple_tag(takes_context=True)
def is_active_url(context, url_name):
    request = context.get('request')
    if not request or not getattr(request, 'resolver_match', None):
        return ''
    return 'is-active' if request.resolver_match.url_name == url_name else ''
