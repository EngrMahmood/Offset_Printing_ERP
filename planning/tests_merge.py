from django.contrib.auth import get_user_model
from django.test import TestCase
from django.urls import reverse

from core.models import UserProfile

from .merge_engine import MergeConfig, allocate_ups, build_suggestions
from .models import MergeGroup, PlanningJob, PlanningPrintRun, SkuRecipe


def set_awc(job, awc_no):
    """AWC lives on the SKU master; create a recipe so awc_no_display resolves."""
    SkuRecipe.objects.update_or_create(sku=job.sku, defaults={'awc_no': awc_no})


def make_job(jc_number, order_qty, **overrides):
    defaults = dict(
        jc_number=jc_number,
        sku=f'SKU-{jc_number}',
        material='Art Card 300gsm',
        size_w_mm=100,
        size_h_mm=200,
        front_colors=4,
        back_colors=0,
        total_colors=4,
        print_passes=1,
        ups=4,
        print_sheet_size='25x36',
        purchase_sheet_size='25x36',
        order_qty=order_qty,
        wastage_sheets=0,
        status='draft',
    )
    defaults.update(overrides)
    return PlanningJob.objects.create(**defaults)


def make_card(job, prefix='JCARD'):
    """A released job card for `job`, with the fields release validation needs."""
    from core.models import JobCard, Machine
    machine, _ = Machine.objects.get_or_create(name='Press-1')
    return JobCard.objects.create(
        job_card_no=f'{prefix}-{job.jc_number}',
        planning_job=job,
        SKU=job.sku,
        order_qty=job.order_qty,
        ups=job.ups,
        total_sheet_quantity=20000,
        total_impressions_required=40000,
        total_colors=job.total_colors,
        wastage=0,
        po_date='2026-07-01',
        machine_name=machine,
        status='released',
    )


class AllocateUpsTests(TestCase):
    def test_ratio_split_across_three_skus(self):
        jobs = [make_job('JC1', 10000), make_job('JC2', 5000), make_job('JC3', 5000)]
        result = allocate_ups(jobs, 4, MergeConfig())
        self.assertIsNotNone(result)
        self.assertEqual([item['allocated_ups'] for item in result['items']], [2, 1, 1])
        self.assertEqual(result['run_sheets'], 5000)
        self.assertEqual(result['worst_overage_pct'], 0.0)

    def test_within_tolerance_is_accepted(self):
        jobs = [make_job('JC1', 10000, ups=2), make_job('JC2', 9800, ups=2)]
        result = allocate_ups(jobs, 2, MergeConfig())
        self.assertIsNotNone(result)
        self.assertLessEqual(result['worst_overage_pct'], 5.0)

    def test_outside_tolerance_is_rejected(self):
        jobs = [make_job('JC1', 10000, ups=2), make_job('JC2', 6000, ups=2)]
        self.assertIsNone(allocate_ups(jobs, 2, MergeConfig()))

    def test_never_under_produces(self):
        jobs = [make_job('JC1', 10000), make_job('JC2', 5000), make_job('JC3', 5000)]
        result = allocate_ups(jobs, 4, MergeConfig())
        for item in result['items']:
            self.assertGreaterEqual(item['planned_produced_qty'], item['net_qty'])


class BuildSuggestionsTests(TestCase):
    def test_rotated_size_still_matches(self):
        make_job('JC1', 10000)
        make_job('JC2', 5000, size_w_mm=200, size_h_mm=100)
        make_job('JC3', 5000)
        suggestions = build_suggestions(list(PlanningJob.objects.all()))
        self.assertEqual(len(suggestions), 1)
        self.assertEqual(len(suggestions[0]['jobs']), 3)

    def test_different_colours_do_not_merge(self):
        make_job('JC1', 10000)
        make_job('JC2', 10000, front_colors=2, total_colors=2)
        self.assertEqual(build_suggestions(list(PlanningJob.objects.all())), [])

    def test_different_material_does_not_merge(self):
        make_job('JC1', 10000)
        make_job('JC2', 10000, material='Duplex Board 350gsm')
        self.assertEqual(build_suggestions(list(PlanningJob.objects.all())), [])

    def test_savings_reported(self):
        make_job('JC1', 10000)
        make_job('JC2', 5000)
        make_job('JC3', 5000)
        suggestion = build_suggestions(list(PlanningJob.objects.all()))[0]
        self.assertEqual(suggestion['savings']['makereadies_saved'], 2)
        self.assertEqual(suggestion['savings']['plates_saved'], 8)

    def test_setup_sheets_saved_uses_25_per_colour(self):
        # 3 jobs, 4 colours -> (3-1) make-readies * 4 colours * 25 sheets
        make_job('JC1', 10000)
        make_job('JC2', 5000)
        make_job('JC3', 5000)
        suggestion = build_suggestions(list(PlanningJob.objects.all()))[0]
        self.assertEqual(suggestion['savings']['setup_sheets_saved'], 200)

    def test_make_ready_minutes_saved(self):
        make_job('JC1', 10000, total_mr_time_minutes=30)
        make_job('JC2', 5000, total_mr_time_minutes=30)
        make_job('JC3', 5000, total_mr_time_minutes=30)
        suggestion = build_suggestions(list(PlanningJob.objects.all()))[0]
        self.assertEqual(suggestion['savings']['mr_minutes_saved'], 60)


class MergeBoardViewTests(TestCase):
    def setUp(self):
        user = get_user_model().objects.create_user('planner', password='pw12345678')
        profile, _ = UserProfile.objects.get_or_create(user=user)
        profile.role = 'planner'
        profile.save()
        self.client.force_login(user)

    def test_printed_and_issued_jobs_are_excluded(self):
        job = make_job('JC1', 10000)
        make_job('JC2', 5000)
        make_job('JC3', 5000)
        PlanningPrintRun.objects.create(planning_job=job, run_index=1, print_qty=100)
        response = self.client.get(reverse('planning:merge_board'))
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.context['eligible_count'], 2)

    def test_accept_creates_group_and_removes_from_pool(self):
        jobs = [make_job('JC1', 10000), make_job('JC2', 5000), make_job('JC3', 5000)]
        response = self.client.post(
            reverse('planning:merge_accept'),
            {'job_ids': [job.id for job in jobs], 'notes': 'gang it'},
            follow=True,
        )
        self.assertEqual(response.status_code, 200)
        group = MergeGroup.objects.get()
        self.assertEqual(group.items.count(), 3)
        self.assertEqual(group.run_sheets, 5000)
        self.assertEqual(jobs[0].active_merge_group, group)

        board = self.client.get(reverse('planning:merge_board'))
        self.assertEqual(board.context['eligible_count'], 0)
        self.assertEqual(board.context['suggestions'], [])

    def test_artwork_code_and_source_awc_snapshot(self):
        jobs = [make_job('JC1', 10000), make_job('JC2', 5000), make_job('JC3', 5000)]
        for job, awc in zip(jobs, ['A-111', 'A-222', 'A-333']):
            set_awc(job, awc)
        self.client.post(reverse('planning:merge_accept'), {'job_ids': [j.id for j in jobs]}, follow=True)
        group = MergeGroup.objects.get()
        self.assertEqual(group.artwork_code, f'AWC-{group.code}')
        source_codes = set(group.items.values_list('source_awc_no', flat=True))
        self.assertTrue({'A-111', 'A-222', 'A-333'}.issubset(source_codes))

    def test_board_shows_sheet_size_and_awc(self):
        set_awc(make_job('JC1', 10000), 'A-111')
        set_awc(make_job('JC2', 5000), 'A-222')
        set_awc(make_job('JC3', 5000), 'A-333')
        response = self.client.get(reverse('planning:merge_board'))
        body = response.content.decode()
        self.assertIn('25x36', body)      # print sheet size
        self.assertIn('A-111', body)      # member source AWC

    def test_double_submit_does_not_create_second_group(self):
        jobs = [make_job('JC1', 10000), make_job('JC2', 5000), make_job('JC3', 5000)]
        payload = {'job_ids': [job.id for job in jobs]}
        self.client.post(reverse('planning:merge_accept'), payload, follow=True)
        self.client.post(reverse('planning:merge_accept'), payload, follow=True)
        self.assertEqual(MergeGroup.objects.count(), 1)

    def test_cancel_releases_jobs(self):
        jobs = [make_job('JC1', 10000), make_job('JC2', 5000), make_job('JC3', 5000)]
        self.client.post(reverse('planning:merge_accept'), {'job_ids': [j.id for j in jobs]}, follow=True)
        group = MergeGroup.objects.get()
        self.client.post(reverse('planning:merge_cancel', args=[group.id]), follow=True)
        group.refresh_from_db()
        self.assertEqual(group.status, 'cancelled')
        board = self.client.get(reverse('planning:merge_board'))
        self.assertEqual(board.context['eligible_count'], 3)


class MergeDownstreamTests(TestCase):
    def setUp(self):
        user = get_user_model().objects.create_user('planner2', password='pw12345678')
        profile, _ = UserProfile.objects.get_or_create(user=user)
        profile.role = 'planner'
        profile.save()
        self.user = user
        self.client.force_login(user)

    def _accept_group(self):
        jobs = [make_job('JC1', 10000), make_job('JC2', 5000), make_job('JC3', 5000)]
        self.client.post(reverse('planning:merge_accept'), {'job_ids': [j.id for j in jobs]}, follow=True)
        return MergeGroup.objects.get(), jobs

    def test_lead_is_largest_ups(self):
        group, jobs = self._accept_group()
        self.assertIsNotNone(group.lead_job)
        lead_item = group.items.get(is_lead=True)
        self.assertEqual(lead_item.allocated_ups, 2)   # the 10000-pc job
        self.assertEqual(group.items.filter(is_lead=True).count(), 1)

    def test_follower_blocked_from_plates(self):
        from planning.services import get_plate_making_prerequisite_errors
        from printing_plates.services import planning_job_should_skip_plate_making
        group, jobs = self._accept_group()
        follower = group.items.filter(is_lead=False).first().planning_job
        self.assertTrue(follower.is_merge_member_follower)
        self.assertTrue(planning_job_should_skip_plate_making(follower))
        errors = get_plate_making_prerequisite_errors(follower)
        self.assertTrue(any('merged into layout' in e for e in errors))

    def test_lead_not_blocked_from_plates(self):
        from printing_plates.services import planning_job_should_skip_plate_making
        group, jobs = self._accept_group()
        self.assertFalse(planning_job_should_skip_plate_making(group.lead_job))

    def test_cancelling_member_dissolves_group(self):
        from planning.services import cancel_planning_job
        group, jobs = self._accept_group()
        follower = group.items.filter(is_lead=False).first().planning_job
        cancel_planning_job(follower, actor=self.user, reason='customer dropped', reason_code='customer_cancelled')
        group.refresh_from_db()
        self.assertEqual(group.status, 'cancelled')

    def test_printing_split_to_member_cards(self):
        from core.models import JobCard, Production
        group, jobs = self._accept_group()
        cards = {job.id: make_card(job) for job in jobs}
        lead = group.lead_job
        entry = Production.objects.create(
            job_card=cards[lead.id],
            entry_type='printing',
            date='2026-07-21',
            shift='A',
            output_sheets=5000,
        )
        # Lead's own entry recounts at combined-sheet ups (2), not job-card ups (4).
        entry.refresh_from_db()
        self.assertEqual(entry.merge_allocated_ups, 2)
        self.assertEqual(entry.pcs_produced, 10000)

        # Each follower card received a derived entry at its allocated ups (1).
        for item in group.items.filter(is_lead=False):
            follower_entries = Production.objects.filter(
                job_card__planning_job=item.planning_job, entry_type='printing',
            )
            self.assertEqual(follower_entries.count(), 1)
            fe = follower_entries.first()
            self.assertEqual(fe.merge_parent_id, entry.id)
            self.assertEqual(fe.merge_allocated_ups, 1)
            self.assertEqual(fe.pcs_produced, 5000)
            self.assertEqual(fe.impressions, 0)

    def test_split_does_not_recurse(self):
        from core.models import JobCard, Production
        group, jobs = self._accept_group()
        for job in jobs:
            make_card(job, prefix='JCARD2')
        Production.objects.create(
            job_card=group.lead_job.job_card, entry_type='printing',
            date='2026-07-21', shift='A', output_sheets=5000,
        )
        # 1 lead + 2 followers, no cascade beyond that.
        self.assertEqual(Production.objects.filter(entry_type='printing').count(), 3)
