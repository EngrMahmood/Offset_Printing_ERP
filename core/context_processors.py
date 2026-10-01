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
