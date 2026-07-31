from rest_framework.permissions import BasePermission

from core.permissions import user_has_permission


class ChatAccessPermission(BasePermission):
    """Base gate: user must hold nav.chat to use any chat endpoint."""

    message = 'You do not have access to Chat.'

    def has_permission(self, request, view):
        return bool(request.user and request.user.is_authenticated) and (
            request.user.is_superuser or user_has_permission(request.user, 'nav.chat')
        )


class CanCreateGroup(BasePermission):
    message = 'You do not have permission to create group chats.'

    def has_permission(self, request, view):
        return request.user.is_superuser or user_has_permission(request.user, 'action.chat_create_group')


class CanManageGroupMembers(BasePermission):
    message = 'You do not have permission to manage group members.'

    def has_permission(self, request, view):
        return request.user.is_superuser or user_has_permission(request.user, 'action.chat_manage_group_members')


class CanInitiateCall(BasePermission):
    message = 'You do not have permission to start calls.'

    def has_permission(self, request, view):
        return request.user.is_superuser or user_has_permission(request.user, 'action.chat_initiate_call')


def can_manage_group(user):
    return user.is_superuser or user_has_permission(user, 'action.chat_manage_group_members')


def can_delete_any_message(user):
    return user.is_superuser or user_has_permission(user, 'action.chat_delete_any_message')
