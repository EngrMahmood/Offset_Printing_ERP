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


class JobCardMergeBannerTests(TestCase):
    """The printed job card must tell the production team what to do."""

    def setUp(self):
        user = get_user_model().objects.create_user('planner3', password='pw12345678')
        profile, _ = UserProfile.objects.get_or_create(user=user)
        profile.role = 'planner'
        profile.save()
        self.user = user
        self.client.force_login(user)
        self.jobs = [make_job('JC1', 10000), make_job('JC2', 5000), make_job('JC3', 5000)]
        for job, awc in zip(self.jobs, ['A-111', 'A-222', 'A-333']):
            set_awc(job, awc)
        self.client.post(
            reverse('planning:merge_accept'),
            {'job_ids': [j.id for j in self.jobs]},
            follow=True,
        )
        self.group = MergeGroup.objects.get()

    def _render_card(self, job):
        from django.template.loader import render_to_string
        from planning.services import build_job_card_merge_context
        return render_to_string('Job Card.html', {
            'job': job,
            'recipe': job.sku_recipe,
            'merge': build_job_card_merge_context(job),
        })

    def test_context_none_for_unmerged_job(self):
        from planning.services import build_job_card_merge_context
        loner = make_job('JC-SOLO', 1000)
        self.assertIsNone(build_job_card_merge_context(loner))

    def test_lead_card_shows_combined_banner_and_all_skus(self):
        html = self._render_card(self.group.lead_job)
        self.assertIn('Combined layout', html)
        self.assertIn(self.group.code, html)
        self.assertIn(self.group.artwork_code, html)
        self.assertNotIn('Do not print separately', html)
        # every member SKU listed on the lead's sheet
        for item in self.group.items.all():
            self.assertIn(item.planning_job.jc_number, html)
        self.assertIn('A-111', html)  # source AWC for the designer/press

    def test_follower_card_warns_not_to_print(self):
        follower_item = self.group.items.filter(is_lead=False).first()
        html = self._render_card(follower_item.planning_job)
        self.assertIn('Do not print separately', html)
        self.assertIn(self.group.lead_job.jc_number, html)
        self.assertNotIn('Combined layout', html)

    def test_merged_ups_annotated_next_to_standalone_ups(self):
        html = self._render_card(self.group.lead_job)
        self.assertIn('(merged: 2)', html)   # lead holds 2 of the 4 ups

    def test_unmerged_card_has_no_banner(self):
        loner = make_job('JC-SOLO2', 1000)
        set_awc(loner, 'A-SOLO')
        loner = PlanningJob.objects.get(pk=loner.pk)  # drop the cached (empty) recipe
        html = self._render_card(loner)
        self.assertNotIn('<div class="merge-banner', html)  # element, not the CSS rule
        self.assertNotIn('Do not print separately', html)
        self.assertNotIn('(merged:', html)

    def test_layout_builder_registers_merge_fields(self):
        from planning.views import _job_card_layout_field_labels
        labels = _job_card_layout_field_labels()
        for key in ['merge_code', 'merge_artwork_code', 'merge_role',
                    'merge_allocated_ups', 'merge_run_sheets']:
            self.assertIn(key, labels)


class MergePlateUiTests(TestCase):
    def setUp(self):
        user = get_user_model().objects.create_user('plateuser', password='pw12345678')
        profile, _ = UserProfile.objects.get_or_create(user=user)
        profile.role = 'admin'
        profile.save()
        self.user = user
        self.client.force_login(user)
        jobs = [make_job('JC1', 10000), make_job('JC2', 5000), make_job('JC3', 5000)]
        self.client.post(reverse('planning:merge_accept'), {'job_ids': [j.id for j in jobs]}, follow=True)
        self.group = MergeGroup.objects.get()

    def test_merge_info_on_plate_request(self):
        from printing_plates.models import PlateRequest
        req = PlateRequest.objects.create(planning_job=self.group.lead_job)
        self.assertIsNotNone(req.merge_info)
        self.assertEqual(req.merge_info['code'], self.group.code)
        self.assertEqual(req.merge_info['member_count'], 3)

    def test_merge_info_none_for_normal_request(self):
        from printing_plates.models import PlateRequest
        loner = make_job('JC-SOLO', 1000)
        req = PlateRequest.objects.create(planning_job=loner)
        self.assertIsNone(req.merge_info)

    def test_merged_layouts_tab_lists_group(self):
        response = self.client.get(reverse('printing_plates:merged_list'))
        self.assertEqual(response.status_code, 200)
        body = response.content.decode()
        self.assertIn(self.group.code, body)
        self.assertIn(self.group.artwork_code, body)
        for item in self.group.items.all():
            self.assertIn(item.planning_job.jc_number, body)

    def test_plate_request_detail_shows_merge_panel(self):
        from printing_plates.models import PlateRequest
        req = PlateRequest.objects.create(planning_job=self.group.lead_job)
        response = self.client.get(reverse('printing_plates:request_detail', args=[req.pk]))
        self.assertEqual(response.status_code, 200)
        self.assertIn('Combined layout', response.content.decode())


class MergeParallelPathwayTests(TestCase):
    """A job with no merge group must behave exactly as before the feature."""

    def setUp(self):
        user = get_user_model().objects.create_user('normaluser', password='pw12345678')
        profile, _ = UserProfile.objects.get_or_create(user=user)
        profile.role = 'planner'
        profile.save()
        self.user = user
        self.job = make_job('JC-NORMAL', 10000)
        set_awc(self.job, 'A-NORM')
        self.job = PlanningJob.objects.get(pk=self.job.pk)

    def test_merge_flags_are_false(self):
        self.assertIsNone(self.job.active_merge_group)
        self.assertIsNone(self.job.active_merge_item)
        self.assertFalse(self.job.is_merge_lead)
        self.assertFalse(self.job.is_merge_member_follower)

    def test_plate_making_not_blocked(self):
        from planning.services import get_plate_making_prerequisite_errors
        from printing_plates.services import planning_job_should_skip_plate_making
        self.assertFalse(planning_job_should_skip_plate_making(self.job))
        errors = get_plate_making_prerequisite_errors(self.job)
        self.assertFalse(any('merged into layout' in e for e in errors))

    def test_production_entry_is_validated_and_not_split(self):
        from core.models import Production
        card = make_card(self.job, prefix='NORM')
        entry = Production.objects.create(
            job_card=card, entry_type='printing', date='2026-07-21', shift='A', output_sheets=100,
        )
        entry.refresh_from_db()
        self.assertIsNone(entry.merge_allocated_ups)
        self.assertIsNone(entry.merge_parent_id)
        # pcs still counted off the job card's own ups, exactly as before
        self.assertEqual(entry.pcs_produced, 100 * card.ups)
        self.assertEqual(Production.objects.filter(entry_type='printing').count(), 1)

    def test_cancelling_normal_job_touches_no_group(self):
        from planning.services import cancel_planning_job
        cancel_planning_job(self.job, actor=self.user, reason='no longer needed',
                            reason_code='customer_cancelled')
        self.job.refresh_from_db()
        self.assertTrue(self.job.is_cancelled)
        self.assertEqual(MergeGroup.objects.count(), 0)


class MergeEligibilityPlateGuardTests(TestCase):
    """Merging after plates exist saves nothing — those jobs must not be offered."""

    def setUp(self):
        user = get_user_model().objects.create_user('planner4', password='pw12345678')
        profile, _ = UserProfile.objects.get_or_create(user=user)
        profile.role = 'planner'
        profile.save()
        self.client.force_login(user)

    def test_job_with_plate_request_is_excluded(self):
        from printing_plates.models import PlateRequest
        job = make_job('JC1', 10000)
        make_job('JC2', 5000)
        make_job('JC3', 5000)
        PlateRequest.objects.create(planning_job=job, status=PlateRequest.STATUS_SENT)
        response = self.client.get(reverse('planning:merge_board'))
        self.assertEqual(response.context['eligible_count'], 2)

    def test_archived_plate_request_does_not_exclude(self):
        from printing_plates.models import PlateRequest
        job = make_job('JC1', 10000)
        make_job('JC2', 5000)
        make_job('JC3', 5000)
        PlateRequest.objects.create(planning_job=job, status=PlateRequest.STATUS_ARCHIVED)
        response = self.client.get(reverse('planning:merge_board'))
        self.assertEqual(response.context['eligible_count'], 3)

    def test_detail_warns_about_preexisting_member_plates(self):
        from printing_plates.models import PlateRequest
        jobs = [make_job('JC1', 10000), make_job('JC2', 5000), make_job('JC3', 5000)]
        self.client.post(reverse('planning:merge_accept'), {'job_ids': [j.id for j in jobs]}, follow=True)
        group = MergeGroup.objects.get()
        follower = group.items.filter(is_lead=False).first().planning_job
        PlateRequest.objects.create(planning_job=follower, status=PlateRequest.STATUS_SENT)
        response = self.client.get(reverse('planning:merge_detail', args=[group.id]))
        self.assertContains(response, 'Plates already exist for members of this group')
        self.assertContains(response, follower.jc_number)
