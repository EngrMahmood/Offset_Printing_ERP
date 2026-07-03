from django.contrib.auth.decorators import login_required
from django.http import JsonResponse
from django.views.decorators.http import require_GET, require_POST

from core.models import Notification
from core.notifications import serialize_notification


@login_required
@require_GET
def notification_list(request):
    limit = min(int(request.GET.get('limit') or 20), 50)
    unread_only = (request.GET.get('unread') or '').strip() in {'1', 'true', 'yes'}
    since_id = request.GET.get('since_id')

    qs = Notification.objects.filter(user=request.user)
    if unread_only:
        qs = qs.filter(is_read=False)
    if since_id:
        try:
            qs = qs.filter(pk__gt=int(since_id))
        except (TypeError, ValueError):
            pass

    items = list(qs[:limit])
    unread_count = Notification.objects.filter(user=request.user, is_read=False).count()
    return JsonResponse({
        'ok': True,
        'unread_count': unread_count,
        'items': [serialize_notification(item) for item in items],
    })


@login_required
@require_POST
def notification_mark_read(request, pk):
    updated = Notification.objects.filter(user=request.user, pk=pk, is_read=False).update(is_read=True)
    unread_count = Notification.objects.filter(user=request.user, is_read=False).count()
    return JsonResponse({'ok': True, 'updated': updated, 'unread_count': unread_count})


@login_required
@require_POST
def notification_mark_all_read(request):
    updated = Notification.objects.filter(user=request.user, is_read=False).update(is_read=True)
    return JsonResponse({'ok': True, 'updated': updated, 'unread_count': 0})
