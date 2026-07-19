from django.db import transaction
from django.urls import reverse
from django.utils import timezone

from core.notifications import notify_roles, notify_users

from .models import ItemProcurementTimeline, ItemRequest, ItemRequestApproval

APPROVER_ROLES = ('manager', 'admin')
SUPPLY_CHAIN_ROLES = ('supply_chain', 'admin')


def _detail_link(item_request):
    try:
        return reverse('supply_chain:item_request_detail', args=[item_request.pk])
    except Exception:
        return ''


def generate_request_no(item_request):
    """Assign a per-year, per-type sequential IR-ID, e.g. IR-MNT-2026-0042."""
    year = timezone.now().year
    prefix = f'IR-{item_request.request_type.code}-{year}-'
    with transaction.atomic():
        last = (
            ItemRequest.objects
            .select_for_update()
            .filter(request_no__startswith=prefix)
            .order_by('-request_no')
            .first()
        )
        next_seq = 1
        if last and last.request_no:
            try:
                next_seq = int(last.request_no.rsplit('-', 1)[-1]) + 1
            except ValueError:
                next_seq = 1
        item_request.request_no = f'{prefix}{next_seq:04d}'
        item_request.save(update_fields=['request_no'])
    return item_request.request_no


def submit_request(item_request, user):
    item_request.status = 'MGR_REVIEW'
    item_request.save(update_fields=['status'])
    ItemRequestApproval.objects.create(
        request=item_request, actor=user, action='SUBMIT', stage='REQUESTER',
    )
    notify_roles(
        APPROVER_ROLES,
        event_type='item_request.submitted',
        title=f'New item request: {item_request.item_title}',
        message=f'{user.username} submitted a request for "{item_request.item_title}" awaiting manager review.',
        link=_detail_link(item_request),
        entity_type='itemrequest',
        entity_id=item_request.pk,
        actor=user,
    )


def resubmit_request(item_request, user):
    item_request.status = 'MGR_REVIEW'
    item_request.save(update_fields=['status'])
    ItemRequestApproval.objects.create(
        request=item_request, actor=user, action='RESUBMIT', stage='REQUESTER',
    )
    notify_roles(
        APPROVER_ROLES,
        event_type='item_request.resubmitted',
        title=f'Item request resubmitted: {item_request.item_title}',
        message=f'{user.username} resubmitted "{item_request.item_title}" for manager review.',
        link=_detail_link(item_request),
        entity_type='itemrequest',
        entity_id=item_request.pk,
        actor=user,
    )


def review_request(item_request, user, stage, action, comment=''):
    """Apply a manager/supply-chain review decision and advance the workflow."""
    ItemRequestApproval.objects.create(
        request=item_request, actor=user, action=action, stage=stage, comment=comment,
    )

    notify_requester = False
    notify_sc = False

    if action == 'REJECT':
        item_request.status = 'REJECTED'
        notify_requester = True
    elif action == 'REVISE':
        item_request.status = 'NEEDS_REVISION'
        notify_requester = True
    elif action == 'APPROVE':
        if stage == 'MANAGER':
            item_request.status = 'SC_REVIEW'
            notify_sc = True
        elif stage == 'SUPPLY_CHAIN':
            item_request.status = 'APPROVED'
            if not item_request.request_no:
                generate_request_no(item_request)
            ItemProcurementTimeline.objects.get_or_create(request=item_request)
            item_request.status = 'IN_PROCUREMENT'
            notify_requester = True

    item_request.save(update_fields=['status'])

    if notify_requester:
        notify_users(
            [item_request.raised_by],
            event_type='item_request.decision',
            title=f'Item request update: {item_request.item_title}',
            message=f'{user.username} marked "{item_request.item_title}" as {item_request.get_status_display()}.'
                    + (f' Comment: {comment}' if comment else ''),
            link=_detail_link(item_request),
            entity_type='itemrequest',
            entity_id=item_request.pk,
            actor=user,
            exclude_actor=False,
        )

    if notify_sc:
        notify_roles(
            SUPPLY_CHAIN_ROLES,
            event_type='item_request.sc_review',
            title=f'Item request awaiting supply-chain review: {item_request.item_title}',
            message=f'"{item_request.item_title}" was approved by manager and needs supply-chain review.',
            link=_detail_link(item_request),
            entity_type='itemrequest',
            entity_id=item_request.pk,
            actor=user,
        )

    return item_request
