from django.core.management.base import BaseCommand

from core.models import JobCard, JobCardWipStatus
from production.printing_completion import approved_completion_job_ids, printing_completion_status
from production.wip_service import update_wip_status_for_job


class Command(BaseCommand):
    help = (
        "Re-check jobs shown as 'Partial Printing' or 'Printing' against the shared printing-complete rule "
        "(order qty reached, every planned sheet used, completed within tolerance, or Production Manager "
        "approved) and move the ones that are really finished to 'Printing Completed'. Supervisor overrides are "
        "left alone unless --include-manual is given. Dry run unless --apply."
    )

    def add_arguments(self, parser):
        parser.add_argument('--apply', action='store_true', help='Write the changes (default is a dry run).')
        parser.add_argument('--include-manual', action='store_true',
                            help='Also reset jobs whose status a supervisor set by hand (the status becomes automatic again).')

    def handle(self, *args, **options):
        approved_ids = approved_completion_job_ids()
        wip = JobCardWipStatus.objects.filter(
            status__name__in=('Partial Printing', 'Printing'), job_card__is_active=True,
            job_card__status__in=('released', 'in_production'),
        ).select_related('status', 'job_card')
        changed, skipped_manual = [], []
        for record in wip:
            card = record.job_card
            if not card.is_print_job:
                continue
            completion = printing_completion_status(card, approved_ids)
            if not (completion['final_pass_started'] and completion['complete']):
                continue
            if record.is_manual and not options['include_manual']:
                skipped_manual.append((card.job_card_no, record.status.name, completion))
                continue
            changed.append((card.job_card_no, record.status.name, completion))
            if options['apply']:
                update_wip_status_for_job(card, 'Printing Completed', user=None, is_manual=False, force=True)
        verb = 'Moved' if options['apply'] else 'Would move'
        self.stdout.write(f'{verb} {len(changed)} job(s) to Printing Completed.')
        for jc, old, c in changed:
            self.stdout.write(f"  {jc}: {old} -> Printing Completed ({c['rule_label']}; {c['shortfall_pcs']:,} pcs / {c['shortfall_pct']}% short)")
        if skipped_manual:
            self.stdout.write(f'{len(skipped_manual)} job(s) left alone because a supervisor set their status by hand '
                              f'(use --include-manual to reset them):')
            for jc, old, c in skipped_manual:
                self.stdout.write(f"  {jc}: shows '{old}' but printing is complete ({c['rule_label']})")
