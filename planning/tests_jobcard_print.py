from django.template.loader import render_to_string
from django.test import TestCase

from .models import MergeGroup, SkuRecipe
from .tests_merge import make_job


class JobCardPrintConsistencyTests(TestCase):
    """The top of the printed job card and its cutting slip must show the same job values."""

    def _render(self, job, **extra):
        job = type(job).objects.get(pk=job.pk)            # fresh instance: the recipe is looked up after it exists
        ctx = {'job': job, 'recipe': job.sku_recipe, 'material_type_clean': job.material_display or '-'}
        ctx.update(extra)
        return render_to_string('Job Card.html', ctx)

    def test_cutting_slip_uses_the_forms_own_material_not_the_sku_master(self):
        job = make_job('JC-11-26-PP-4605.1', 130, sku='SKU-NCR', material='NCR Yellow', print_sheet_size='12*17')
        SkuRecipe.objects.create(sku='SKU-NCR', material='NCR White', print_sheet_size='12*17')
        html = self._render(job)
        self.assertEqual(html.count('NCR Yellow'), 2)           # material type at the top AND the cutting slip "Type"
        self.assertNotIn('NCR White', html)

    def test_cutting_slip_sheet_sizes_and_ups_follow_the_job(self):
        job = make_job('JC-11-26-PP-4606.1', 130, sku='SKU-NCR2', purchase_sheet_size='24*17', purchase_sheet_ups=2, print_sheet_size='12*17')
        SkuRecipe.objects.create(sku='SKU-NCR2', purchase_sheet_size='99*99', purchase_sheet_ups=9, print_sheet_size='88*88', ups=7)
        html = self._render(job)
        self.assertIn('24*17', html)
        self.assertNotIn('99*99', html)
        self.assertNotIn('88*88', html)


class CombinedRunWastageTests(TestCase):
    def test_the_combined_run_counts_the_wastage_sheets_too(self):
        lead = make_job('JC-09-26-PP-2746', 318, ups=6, print_passes=1)
        group = MergeGroup.objects.create(code='MG-TEST-58', status='layout_done', total_sheet_ups=10, run_sheets=53,
                                          combined_wastage_sheets=5, lead_job=lead)
        self.assertEqual(group.total_run_sheets, 58)
        self.assertEqual(group.combined_impressions(), 58)          # 53 + 5 wastage, one pass

    def test_the_printing_form_shows_58_for_the_lead_job(self):
        from production.printing_entry_helpers import build_printing_job_card_maps
        from .models import MergeGroupItem
        from .tests_merge import make_card
        lead = make_job('JC-09-26-PP-2746', 318, ups=6, print_passes=1)
        card = make_card(lead)
        group = MergeGroup.objects.create(code='MG-TEST-58', status='layout_done', total_sheet_ups=10, run_sheets=53,
                                          combined_wastage_sheets=5, lead_job=lead)
        MergeGroupItem.objects.create(merge_group=group, planning_job=lead, allocated_ups=6, is_lead=True)
        _plan, _machine, info = build_printing_job_card_maps([card])
        self.assertEqual(info[str(card.id)]['required_sheets'], '58')
        self.assertEqual(info[str(card.id)]['allowed_impressions'], '58')
