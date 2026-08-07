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
    if not item_request.request_no:
        generate_request_no(item_request)
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


def _link_existing_sku(timeline, item_request):
    """Carry a pre-existing SKU onto the timeline and skip the code-opening stage.

    When the requester already identified the item as an existing SKU there is
    nothing to open, so the timeline jumps straight to indent/PR. Otherwise the
    SKU fields stay blank — that is exactly the duration the timeline measures.
    """
    sku = item_request.existing_sku
    if not sku:
        return timeline
    timeline.sku = sku
    timeline.item_code = sku.sku
    timeline.sku_pre_existing = True
    timeline.code_opened_date = timeline.code_opened_date or timezone.localdate()
    timeline.save(update_fields=['sku', 'item_code', 'sku_pre_existing', 'code_opened_date'])
    return timeline


def pending_review_count(user):
    """Number of item requests waiting on *this* user's decision.

    Mirrors the stage gating in ``views.item_request_review``: managers own
    MGR_REVIEW, supply chain owns SC_REVIEW, admins own both.
    """
    if not user or not getattr(user, 'is_authenticated', False):
        return 0

    profile = getattr(user, 'profile', None)
    role = (getattr(profile, 'role', '') or '').strip().lower()
    if user.is_superuser:
        role = role or 'admin'

    statuses = []
    if role in APPROVER_ROLES:
        statuses.append('MGR_REVIEW')
    if role in SUPPLY_CHAIN_ROLES:
        statuses.append('SC_REVIEW')
    if not statuses:
        return 0
    return ItemRequest.objects.filter(status__in=statuses, is_active=True).count()


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
            timeline, _ = ItemProcurementTimeline.objects.get_or_create(request=item_request)
            _link_existing_sku(timeline, item_request)
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
