"""Sheet arithmetic for jobs printed on a combined (smart-merge) layout.

A job's own plan is worked out at its standalone ups (JobCard.ups). When it is printed on a
combined sheet it only owns ``merge_allocated_ups`` of that sheet, so the sheets that ran
through the press say nothing about how much of the job's own plan was used. Expressing them
"at the job's standalone ups" (sheets x merge ups / card ups) puts them on the same footing as
the plan, tolerance and sheet budget. Ordinary entries are returned unchanged.
"""
from django.db.models import Case, F, FloatField, Q, Value, When
from django.db.models.functions import Cast


def standalone_sheets(sheets, merge_ups, card_ups):
    sheets = int(sheets or 0)
    if not merge_ups or not card_ups or merge_ups == card_ups:
        return sheets
    return int(round(sheets * merge_ups / card_ups))


def standalone_sheets_expr(field='output_sheets'):
    """Database expression equivalent of ``standalone_sheets`` for a Production queryset."""
    as_float = Cast(F(field), FloatField())
    return Case(
        When(
            Q(merge_allocated_ups__gt=0, job_card__ups__gt=0),
            then=as_float * Cast(F('merge_allocated_ups'), FloatField()) / Cast(F('job_card__ups'), FloatField()),
        ),
        default=as_float,
        output_field=FloatField(),
    )


def piece_ups_expr():
    """Ups used to turn sheets into pieces: the merge share when there is one, else the card's."""
    return Case(
        When(merge_allocated_ups__gt=0, then=F('merge_allocated_ups')),
        When(job_card__ups__gt=0, then=F('job_card__ups')),
        default=Value(1),
        output_field=FloatField(),
    )


def member_share_of_run(output_sheets, waste_sheets, impressions, member_ups, total_ups, card_ups):
    """One member job's share of a combined run, as the values for its derived Production entry.

    Impressions are divided by the member's share of the combined sheet (member_ups / total_ups).
    Good and waste sheets are expressed at the job's own standalone ups (card_ups), so that
    sheets x card ups is exactly the pieces this job got from the run: the rows across the group
    then add up to what was printed, and the job's plan and tolerance compare like for like.
    Without a usable card ups the full run sheets are kept and counted at the member's ups instead.
    """
    member_ups = int(member_ups or 0)
    total_ups = int(total_ups or 0)
    card_ups = int(card_ups or 0)
    share_of_sheet = (member_ups / total_ups) if member_ups and total_ups else 0
    impressions_share = int(round((impressions or 0) * share_of_sheet))
    if card_ups and member_ups:
        return {
            'output_sheets': int(round((output_sheets or 0) * member_ups / card_ups)),
            'waste_sheets': int(round((waste_sheets or 0) * member_ups / card_ups)),
            'impressions': impressions_share,
            'merge_allocated_ups': None,
        }
    return {
        'output_sheets': int(output_sheets or 0),
        'waste_sheets': 0,
        'impressions': impressions_share,
        'merge_allocated_ups': member_ups or None,
    }


def group_for_run(entry):
    """The combined-layout group a lead's printing entry belongs to, or None.

    Prefers the group that is still open; a finished group (no longer 'open') is still found by
    the lead job so older runs can be recomputed."""
    from planning.models import MergeGroup

    planning_job = getattr(entry.job_card, 'planning_job', None)
    if not planning_job:
        return None
    group = planning_job.active_merge_group
    if group and group.lead_job_id == planning_job.id:
        return group
    return (
        MergeGroup.objects.filter(lead_job=planning_job).exclude(status='cancelled')
        .order_by('-created_at').first()
    )


def refresh_member_entries(parent, group=None):
    """Recompute every derived member entry of ``parent`` from the lead's current values.

    Used when the lead's entry is edited (or archived/restored) so the members' shares follow it.
    Returns the number of member entries changed."""
    from core.models import Production

    group = group or group_for_run(parent)
    if group is None:
        return 0
    items = {it.planning_job_id: it for it in group.items.select_related('planning_job')}
    total_ups = sum(it.allocated_ups or 0 for it in items.values()) or (group.total_sheet_ups or 0)
    changed = 0
    for child in Production.objects.filter(merge_parent=parent).select_related('job_card'):
        item = items.get(getattr(child.job_card, 'planning_job_id', None))
        if item is None:
            continue
        share = member_share_of_run(
            parent.output_sheets, parent.waste_sheets, parent.impressions,
            item.allocated_ups, total_ups, child.job_card.ups,
        )
        wanted = dict(
            output_sheets=share['output_sheets'], waste_sheets=share['waste_sheets'],
            impressions=share['impressions'], merge_allocated_ups=share['merge_allocated_ups'],
            is_active=parent.is_active, date=parent.date, shift=parent.shift, machine_id=parent.machine_id,
            status=parent.status,
        )
        if any(getattr(child, key) != value for key, value in wanted.items()):
            for key, value in wanted.items():
                setattr(child, key, value)
            child.save()
            changed += 1
    return changed
