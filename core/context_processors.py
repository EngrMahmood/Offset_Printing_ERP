import json

from django.conf import settings
from core.navigation import get_nav_permissions


def software_version(request):
    return {
        'erp_software_version': getattr(settings, 'ERP_SOFTWARE_VERSION', '0.0.0'),
        'erp_software_release_date': getattr(settings, 'ERP_SOFTWARE_RELEASE_DATE', ''),
    }


def navigation_permissions(request):
    return {
        'nav': get_nav_permissions(request),
    }


def ai_features(request):
    """Whether the dashboard-wide Ask AI widget should render for this
    request. Cheap (singleton row, already cached per-request by Django's
    query cache) — avoids every page needing its own view code just to gate
    one floating button."""
    if not getattr(request, 'user', None) or not request.user.is_authenticated:
        return {'ask_ai_widget_enabled': False}
    try:
        from core.models import AISettings
        row = AISettings.objects.first()
        enabled = bool(row and row.ai_enabled and row.chat_assistant_enabled)
    except Exception:
        enabled = False
    return {'ask_ai_widget_enabled': enabled}


_REPOST_SKIP = ('csrfmiddlewaretoken', 'password', 'pass', 'token')


def form_repost(request):
    """When a form POST comes back as a page (validation error), give the page what was typed so the form can be
    refilled instead of showing blank. Successful saves redirect, so this only fires on an error re-render."""
    if request.method != 'POST':
        return {}
    data = {}
    for key in request.POST.keys():
        if any(word in key.lower() for word in _REPOST_SKIP):
            continue
        values = request.POST.getlist(key)
        data[key] = values if (key.endswith('[]') or len(values) > 1) else values[0]
    if not data:
        return {}
    return {'repost_json': json.dumps({'path': request.path, 'data': data})}
