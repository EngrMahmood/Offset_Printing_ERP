from django.core.management.base import BaseCommand

from core.merge_sheets import group_for_run, member_share_of_run
from core.models import Production


class Command(BaseCommand):
    help = (
        "Convert combined-layout member entries written before runs were divided by share: they carried the "
        "lead's full sheets with no impressions. Each becomes its share of the lead's run (good sheets and waste "
        "at the job's own ups, impressions by share of the combined sheet). Dry run unless --apply."
    )

    def add_arguments(self, parser):
        parser.add_argument('--apply', action='store_true', help='Write the changes (default is a dry run).')

    def handle(self, *args, **options):
        # Old-style member entries are the ones still marked with a merge share; converted ones are not.
        old = Production.objects.filter(
            entry_type='printing', merge_parent__isnull=False, merge_allocated_ups__isnull=False,
        ).select_related('merge_parent', 'merge_parent__job_card', 'job_card')
        total = old.count()
        converted = skipped = 0
        before_sheets = after_sheets = 0
        for child in old:
            parent = child.merge_parent
            group = group_for_run(parent)
            if group is None:
                skipped += 1
                continue
            items = list(group.items.all())
            item = next((it for it in items if it.planning_job_id == child.job_card.planning_job_id), None)
            if item is None:
                skipped += 1
                continue
            total_ups = sum(it.allocated_ups or 0 for it in items) or (group.total_sheet_ups or 0)
            share = member_share_of_run(
                parent.output_sheets, parent.waste_sheets, parent.impressions,
                item.allocated_ups, total_ups, child.job_card.ups,
            )
            before_sheets += child.output_sheets
            after_sheets += share['output_sheets']
            converted += 1
            if options['apply']:
                # update() on purpose: a plain re-write of history, no workflow side effects.
                Production.objects.filter(pk=child.pk).update(
                    output_sheets=share['output_sheets'], waste_sheets=share['waste_sheets'],
                    impressions=share['impressions'], merge_allocated_ups=share['merge_allocated_ups'],
                )
        verb = 'Converted' if options['apply'] else 'Would convert'
        self.stdout.write(
            f'{verb} {converted} of {total} member entr{"y" if total == 1 else "ies"} '
            f'(good sheets {before_sheets:,} -> {after_sheets:,}); {skipped} skipped (group not found).'
        )
