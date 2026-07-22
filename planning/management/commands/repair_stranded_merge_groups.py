"""Reset merge groups that predate the group-approval gate.

Groups created before Round 7 sit at ``artwork_requested`` / ``artwork_ready``
without ever passing the combined-layout production approval — their member cards
were never production-approved as a group, so they cannot cleanly reach printing.
This command moves such groups (no ``layout_approved_at`` and no combined plate yet)
back to ``accepted`` so the planner can click "Approve combined layout" and run the
proper pipeline. It changes nothing else: items, lead and savings are untouched.

Idempotent: a group already at ``accepted`` / ``layout_approved`` / ``cancelled``,
or one whose combined plate is already raised, is skipped.
"""

from django.core.management.base import BaseCommand

from planning.models import MergeGroup
from printing_plates.services import combined_plate_request_for_group


class Command(BaseCommand):
    help = 'Reset pre-approval merge groups back to accepted so the layout gate applies.'

    def add_arguments(self, parser):
        parser.add_argument('--apply', action='store_true', help='Persist changes (otherwise dry-run).')

    def handle(self, *args, **options):
        apply = options['apply']
        candidates = MergeGroup.objects.filter(
            status__in=['artwork_requested', 'artwork_ready', 'layout_done'],
            layout_approved_at__isnull=True,
        )
        touched = 0
        for group in candidates:
            if combined_plate_request_for_group(group):
                self.stdout.write(f'skip {group.code}: combined plate already raised')
                continue
            touched += 1
            if apply:
                group.status = 'accepted'
                group.designer_requested_at = None
                group.designer_notes = ''
                group.save(update_fields=['status', 'designer_requested_at', 'designer_notes'])
            self.stdout.write(f'{"reset" if apply else "would reset"} {group.code} -> accepted')

        self.stdout.write(self.style.SUCCESS(
            f'{"Reset" if apply else "Would reset"} {touched} group(s).'
            + ('' if apply else ' Re-run with --apply to persist.')
        ))
