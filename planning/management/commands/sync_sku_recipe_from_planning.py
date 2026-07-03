from django.core.management.base import BaseCommand

from planning.models import PlanningJob
from planning.services import (
    apply_designer_layout_to_sku_recipe,
    ensure_sku_recipe_for_planning_job,
    sync_planning_job_fields_to_sku_recipe,
)


class Command(BaseCommand):
    help = 'Create or backfill SKU master rows from planning job designer/planner fields.'

    def add_arguments(self, parser):
        parser.add_argument('--jc-number', dest='jc_number', help='Single JC number to sync')
        parser.add_argument('--sku', dest='sku', help='Single SKU to sync')
        parser.add_argument('--submit-review', action='store_true', help='Mark synced recipes pending_review when designer data exists')

    def handle(self, *args, **options):
        jc_number = (options.get('jc_number') or '').strip()
        sku = (options.get('sku') or '').strip()
        submit_review = bool(options.get('submit_review'))

        qs = PlanningJob.objects.all().order_by('jc_number')
        if jc_number:
            qs = qs.filter(jc_number__iexact=jc_number)
        if sku:
            qs = qs.filter(sku__iexact=sku)

        created = 0
        updated = 0
        for job in qs:
            if not (job.sku or '').strip():
                continue
            recipe = ensure_sku_recipe_for_planning_job(job, create_if_missing=True)
            if not recipe:
                continue
            was_new = recipe._state.adding
            if sync_planning_job_fields_to_sku_recipe(job, recipe, submit_for_review=submit_review):
                updated += 1
            elif was_new:
                created += 1
            apply_designer_layout_to_sku_recipe(
                job,
                recipe,
                {
                    'size_w_mm': job.size_w_mm,
                    'size_h_mm': job.size_h_mm,
                    'ups': job.ups,
                    'print_sheet_size': job.print_sheet_size,
                    'purchase_sheet_size': job.purchase_sheet_size,
                    'purchase_sheet_ups': job.purchase_sheet_ups,
                    'plate_color': job.color_spec,
                    'set_no': job.plate_set_no,
                },
                submit_for_review=submit_review,
            )
            self.stdout.write(self.style.SUCCESS(f'Synced {job.jc_number} -> {recipe.sku} ({recipe.master_data_status})'))

        self.stdout.write(self.style.SUCCESS(f'Done. Updated {updated} recipe(s).'))
