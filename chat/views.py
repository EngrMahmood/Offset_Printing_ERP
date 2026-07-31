import json

from django.contrib.auth.decorators import login_required
from django.core.exceptions import PermissionDenied
from django.shortcuts import render

from core.permissions import user_has_permission

from .permissions import can_manage_group


@login_required
def shell(request):
    if not (request.user.is_superuser or user_has_permission(request.user, 'nav.chat')):
        raise PermissionDenied('You do not have access to Chat.')

    can_create_group = request.user.is_superuser or user_has_permission(request.user, 'action.chat_create_group')
    can_initiate_call = request.user.is_superuser or user_has_permission(request.user, 'action.chat_initiate_call')

    context = {
        'can_create_group': can_create_group,
        'can_initiate_call': can_initiate_call,
        'can_manage_group': can_manage_group(request.user),
        'is_superuser': request.user.is_superuser,
        'current_user_id': request.user.id,
        'current_username': request.user.username,
    }
    return render(request, 'chat/shell.html', context)
