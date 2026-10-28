"""When is a job's printing finished?

One definition, used by the WIP status, the Pending Work printing list, the KPI "not printed" gap and
the check on the printing entry form, so they can never disagree.

A job's printing is complete when its final pass has good sheets and any of these holds:
  1. Order quantity: the pieces printed reach the order quantity.
  2. Sheets used: good + waste sheets have used every planned sheet (nothing left to print with).
  3. Completed within tolerance: a final-pass entry is marked "completed" and the pieces printed are
     within the job's production tolerance of the order quantity.
  4. Approved: the Production Manager approved closing the printing short (PrintingCompletionRequest).

Sheets are counted at the job's own ups, so a job printed on a combined layout is compared with its own
plan fairly (see core.merge_sheets).
"""
from __future__ import annotations

from core.models import PrintingCompletionRequest

RULE_LABELS = {
    'order_qty': 'order quantity reached',
    'sheets_used': 'all planned sheets used',
    'completed_within_tolerance': 'marked completed, within tolerance',
    'pm_approved': 'approved by the Production Manager',
}


def printing_completion_status(job_card, approved_job_ids=None) -> dict:
    """Where the job's printing stands. ``approved_job_ids`` lets a report that checks many jobs pass the
    set of job cards with an approved short-close once, instead of one query per job."""
    from production.printing_pass_helpers import get_job_card_pass_count

    printing = [p for p in job_card.printing_productions]
    total_passes = get_job_card_pass_count(job_card)
    final_entries = [p for p in printing if p.print_pass_number == total_passes and (p.output_sheets or 0) > 0]

    order_pcs = int(job_card.order_qty_pcs or 0)
    printed_pcs = int(job_card.total_printed_pcs or 0)
    shortfall = max(order_pcs - printed_pcs, 0)
    good = sum(p.output_sheets or 0 for p in printing)
    sheets_used = sum(p.standalone_total_sheets for p in printing)
    planned = int(job_card.total_sheets_planned or 0)
    tolerance_pct = float(job_card.production_tolerance_percent or 0)
    within_tolerance = order_pcs > 0 and printed_pcs >= order_pcs * (1 - tolerance_pct / 100.0)
    marked_completed = any(p.status == 'completed' for p in final_entries)
    if approved_job_ids is not None:
        approved = job_card.pk in approved_job_ids
    else:
        approved = PrintingCompletionRequest.objects.filter(job_card=job_card, status='approved').exists()

    rule = None
    if final_entries and order_pcs > 0:
        if printed_pcs >= order_pcs:
            rule = 'order_qty'
        elif planned > 0 and sheets_used >= planned:
            rule = 'sheets_used'
        elif marked_completed and within_tolerance:
            rule = 'completed_within_tolerance'
        elif approved:
            rule = 'pm_approved'

    return {
        'complete': rule is not None,
        'rule': rule,
        'rule_label': RULE_LABELS.get(rule, ''),
        'final_pass_started': bool(final_entries),
        'order_pcs': order_pcs,
        'printed_pcs': printed_pcs,
        'shortfall_pcs': shortfall,
        'shortfall_pct': round(shortfall * 100.0 / order_pcs, 2) if order_pcs else 0.0,
        'good_sheets': good,
        'sheets_used': sheets_used,
        'planned_sheets': planned,
        'tolerance_pct': tolerance_pct,
        'within_tolerance': within_tolerance,
        'marked_completed': marked_completed,
        'approved': approved,
    }


def is_printing_complete(job_card, approved_job_ids=None) -> bool:
    return printing_completion_status(job_card, approved_job_ids)['complete']


def pending_printing_pcs(job_card, approved_job_ids=None) -> int:
    """Pieces the job still owes at printing: nothing once printing is complete, otherwise the shortfall."""
    status = printing_completion_status(job_card, approved_job_ids)
    return 0 if status['complete'] else status['shortfall_pcs']


def approved_completion_job_ids() -> set:
    return set(PrintingCompletionRequest.objects.filter(status='approved').values_list('job_card_id', flat=True))


def shortfall_message(status: dict) -> str:
    return (
        f"Printing is not finished: {status['printed_pcs']:,} of {status['order_pcs']:,} pcs printed "
        f"({status['shortfall_pcs']:,} pcs / {status['shortfall_pct']}% short, tolerance {status['tolerance_pct']:g}%), "
        f"and only {status['sheets_used']:,} of {status['planned_sheets']:,} planned sheets are used."
    )


def guard_completed_entry(job_card, entry, user, reason):
    """Run inside the transaction that saves a printing entry the supervisor marked "completed".

    When the job's printing really is finished by the rules above, nothing happens. When it is not, the
    completion cannot simply be accepted: with no reason the save is refused (ValueError rolls the whole
    entry back), and with a reason the entry is kept as "in progress" and a PrintingCompletionRequest is
    queued for the Production Manager. Returns that request, or None when the completion stood."""
    from production.printing_pass_helpers import get_job_card_pass_count

    if entry.entry_type != 'printing' or entry.status != 'completed' or entry.merge_parent_id:
        return None
    # Only the final pass can finish a job; an intermediate pass marked completed is just that pass done.
    if entry.print_pass_number != get_job_card_pass_count(job_card) or (entry.output_sheets or 0) <= 0:
        return None
    status = printing_completion_status(job_card)
    if status['complete']:
        return None
    reason = (reason or '').strip()
    if not reason:
        raise ValueError(
            shortfall_message(status)
            + ' To close it anyway, write the reason in Remarks and save again: the completion then goes '
              'to the Production Manager for approval.'
        )
    entry.status = 'in_progress'
    entry.save(update_fields=['status'])
    request = PrintingCompletionRequest.objects.filter(job_card=job_card, status='pending').first()
    fields = dict(
        production=entry, requested_by=user, reason=reason, order_pcs=status['order_pcs'],
        printed_pcs=status['printed_pcs'], shortfall_pcs=status['shortfall_pcs'], shortfall_pct=status['shortfall_pct'],
    )
    if request is None:
        request = PrintingCompletionRequest.objects.create(job_card=job_card, **fields)
    else:
        for key, value in fields.items():
            setattr(request, key, value)
        request.save()
    return request
