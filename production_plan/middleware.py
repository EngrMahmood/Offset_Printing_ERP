from django.contrib import messages
from django.shortcuts import redirect
from django.urls import reverse


class PlanApprovalGateMiddleware:
    """From the 6:30 pm draft until a planner approves it, planner users cannot use the planning pages.

    Only the Planner role is held; admins and everyone else are not. The gate leads to the plan board,
    where the plan is checked, changed and approved, so nothing else about the ERP is locked."""

    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        return self.get_response(request)

    def process_view(self, request, view_func, view_args, view_kwargs):
        user = getattr(request, 'user', None)
        match = getattr(request, 'resolver_match', None)
        if user is None or not user.is_authenticated or match is None or match.namespace != 'planning':
            return None
        if user.is_superuser or user.is_staff:
            return None
        profile = getattr(user, 'profile', None)
        if not profile or profile.role != 'planner':
            return None
        from production_plan import services
        from production_plan.models import PlanSettings

        try:
            if not PlanSettings.get().gate_enabled or services.pending_draft() is None:
                return None
        except Exception:                                   # the plan tables do not exist yet (before migrate)
            return None
        messages.error(request, 'The production plan is waiting for your approval. Check it and approve it first; '
                                'planning opens again as soon as it is approved.')
        return redirect(reverse('production_plan:board'))
