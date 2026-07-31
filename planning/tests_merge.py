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
        # The soft-coded access-control system (core/permissions.py) has no
        # seed data in a fresh test DB (Role/Permission rows are created by
        # the seed_access_control management command, not by a migration),
        # so grant this test user the specific permission its views require.
        from core.models import Permission, UserPermissionOverride
        permission, _ = Permission.objects.get_or_create(
            code='action.plan', defaults={'name': 'Plan'},
        )
        UserPermissionOverride.objects.get_or_create(
            user=user, permission=permission, defaults={'granted': True},
        )

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

        # Each follower card received a derived entry at its allocated ups (1),
        # and was auto-started out of 'released' so it becomes dispatchable.
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

            follower_card = cards[item.planning_job_id]
            follower_card.refresh_from_db()
            self.assertEqual(follower_card.status, 'in_production')

        from core.models import JOB_CARD_DISPATCHABLE_STATUSES
        for item in group.items.filter(is_lead=False):
            follower_card = cards[item.planning_job_id]
            self.assertIn(follower_card.status, JOB_CARD_DISPATCHABLE_STATUSES)

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

    def test_follower_card_watermarked_and_names_lead(self):
        follower_item = self.group.items.filter(is_lead=False).first()
        html = self._render_card(follower_item.planning_job)
        self.assertIn('DO NOT PRINT SEPARATELY', html)   # watermark
        self.assertIn(self.group.lead_job.jc_number, html)   # flag names the lead
        self.assertNotIn('<div class="merge-banner', html)

    def test_combined_sheet_lists_all_skus(self):
        from django.template.loader import render_to_string
        html = render_to_string(
            'planning/planning_merge_combined_sheet.html', {'group': self.group}
        )
        self.assertIn(self.group.code, html)
        self.assertIn(self.group.artwork_code, html)
        for item in self.group.items.all():
            self.assertIn(item.planning_job.jc_number, html)

    def test_unmerged_card_is_the_plain_card(self):
        loner = make_job('JC-SOLO2', 1000)
        set_awc(loner, 'A-SOLO')
        loner = PlanningJob.objects.get(pk=loner.pk)  # drop the cached (empty) recipe
        html = self._render_card(loner)
        # No merge markers at all on a normal card.
        self.assertNotIn('<div class="merge-watermark', html)
        self.assertNotIn('<div class="merge-flag', html)
        self.assertNotIn('DO NOT PRINT SEPARATELY', html)
        self.assertNotIn('printMerge(', html)          # no old style switcher
        self.assertEqual(html.count('window.print()'), 1)

    def test_merged_card_is_normal_plus_watermark(self):
        """Round 9: merged card stays the familiar single-SKU card + watermark."""
        lead = self.group.lead_job
        html = self._render_card(lead)
        # Watermark + small flag, but NOT the old banner / toolbar / combined table.
        self.assertIn('<div class="merge-watermark', html)
        self.assertIn('DO NOT PRINT SEPARATELY', html)
        self.assertIn('<div class="merge-flag', html)
        self.assertIn(self.group.code, html)
        self.assertNotIn('<div class="merge-banner', html)
        self.assertNotIn('printMerge(', html)
        self.assertNotIn('id="combined-sheet-page"', html)
        # Card body keeps the SKU's own numbers (no combined-run override).
        self.assertNotIn('(combined run)', html)
        self.assertNotIn('(merged:', html)
        # One plain Print button, same as a normal card.
        self.assertEqual(html.count('window.print()'), 1)


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

    def test_job_with_plate_request_is_still_offered(self):
        # Plates existing is no longer a blocker: the make-ready and setup saving
        # still applies, so the planner decides scrap/retain/exclude at accept.
        from printing_plates.models import PlateRequest
        job = make_job('JC1', 10000)
        make_job('JC2', 5000)
        make_job('JC3', 5000)
        plate = PlateRequest.objects.create(planning_job=job, status=PlateRequest.STATUS_SENT)
        response = self.client.get(reverse('planning:merge_board'))
        self.assertEqual(response.context['eligible_count'], 3)
        offered = {
            item['job'].id: item['job']
            for suggestion in response.context['suggestions']
            for item in suggestion['allocation']['items']
        }
        self.assertEqual(offered[job.id].existing_plate_request, plate)

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
        self.assertContains(response, 'Plates still exist for members of this group')
        self.assertContains(response, follower.jc_number)


class MergePlateDispositionTests(TestCase):
    """Existing plates are a decision, not a blocker."""

    def setUp(self):
        user = get_user_model().objects.create_user('planner5', password='pw12345678')
        profile, _ = UserProfile.objects.get_or_create(user=user)
        profile.role = 'admin'
        profile.save()
        self.user = user
        self.client.force_login(user)

    def _jobs_with_plates(self):
        from printing_plates.models import PlateRequest
        jobs = [make_job('JC1', 10000), make_job('JC2', 5000), make_job('JC3', 5000)]
        plates = {}
        for job in jobs:
            plates[job.id] = PlateRequest.objects.create(
                planning_job=job, status=PlateRequest.STATUS_AVAILABLE,
            )
        return jobs, plates

    def test_plated_jobs_are_offered_again(self):
        jobs, _ = self._jobs_with_plates()
        response = self.client.get(reverse('planning:merge_board'))
        self.assertEqual(response.context['eligible_count'], 3)
        self.assertTrue(response.context['suggestions'])

    def test_scrap_archives_follower_plates(self):
        from printing_plates.models import PlateRequest
        jobs, plates = self._jobs_with_plates()
        payload = {'job_ids': [j.id for j in jobs]}
        for job in jobs:
            payload[f'plate_action_{job.id}'] = 'scrap'
        self.client.post(reverse('planning:merge_accept'), payload, follow=True)
        group = MergeGroup.objects.get()
        for job in jobs:
            plate = PlateRequest.objects.get(pk=plates[job.id].pk)
            if job.id == group.lead_job_id:
                self.assertNotEqual(plate.status, PlateRequest.STATUS_ARCHIVED)
            else:
                self.assertEqual(plate.status, PlateRequest.STATUS_ARCHIVED)
                self.assertFalse(plate.retained_for_reuse)

    def test_retain_parks_plate_for_reuse(self):
        from printing_plates.models import PlateRequest
        from printing_plates.services import get_retained_plate_for_sku
        jobs, plates = self._jobs_with_plates()
        payload = {'job_ids': [j.id for j in jobs]}
        for job in jobs:
            payload[f'plate_action_{job.id}'] = 'retain'
        self.client.post(reverse('planning:merge_accept'), payload, follow=True)
        group = MergeGroup.objects.get()
        follower = group.items.filter(is_lead=False).first().planning_job
        plate = PlateRequest.objects.get(pk=plates[follower.id].pk)
        self.assertTrue(plate.retained_for_reuse)
        self.assertIsNotNone(plate.retained_at)
        self.assertEqual(plate.status, PlateRequest.STATUS_ARCHIVED)
        self.assertEqual(get_retained_plate_for_sku(follower.sku), plate)

    def test_exclude_drops_member_and_reallocates(self):
        jobs, _ = self._jobs_with_plates()
        payload = {'job_ids': [j.id for j in jobs]}
        payload[f'plate_action_{jobs[0].id}'] = 'exclude'   # the 10000-pc job
        payload[f'plate_action_{jobs[1].id}'] = 'scrap'
        payload[f'plate_action_{jobs[2].id}'] = 'scrap'
        self.client.post(reverse('planning:merge_accept'), payload, follow=True)
        group = MergeGroup.objects.get()
        self.assertEqual(group.items.count(), 2)
        self.assertNotIn(jobs[0].id, group.items.values_list('planning_job_id', flat=True))

    def test_excluding_too_many_is_rejected(self):
        jobs, _ = self._jobs_with_plates()
        payload = {'job_ids': [j.id for j in jobs]}
        for job in jobs[:2]:
            payload[f'plate_action_{job.id}'] = 'exclude'
        payload[f'plate_action_{jobs[2].id}'] = 'scrap'
        self.client.post(reverse('planning:merge_accept'), payload, follow=True)
        self.assertEqual(MergeGroup.objects.count(), 0)

    def test_retained_plate_shows_on_hold_tab_and_job_detail(self):
        jobs, _ = self._jobs_with_plates()
        payload = {'job_ids': [j.id for j in jobs]}
        for job in jobs:
            payload[f'plate_action_{job.id}'] = 'retain'
        self.client.post(reverse('planning:merge_accept'), payload, follow=True)
        group = MergeGroup.objects.get()
        follower = group.items.filter(is_lead=False).first().planning_job

        hold = self.client.get(reverse('printing_plates:on_hold_list'))
        self.assertEqual(hold.status_code, 200)
        self.assertIn(follower.sku, hold.content.decode())

        # A NEW job for the same SKU is told the plate set is waiting.
        later = make_job('JC-LATER', 4000, sku=follower.sku)
        detail = self.client.get(reverse('planning:job_detail', args=[later.id]))
        self.assertContains(detail, 'A plate set for this SKU is on hold')


class MergeBlockerTests(TestCase):
    def setUp(self):
        user = get_user_model().objects.create_user('planner6', password='pw12345678')
        profile, _ = UserProfile.objects.get_or_create(user=user)
        profile.role = 'planner'
        profile.save()
        self.client.force_login(user)

    def test_complete_job_has_no_blockers(self):
        from planning.merge_engine import merge_blockers
        self.assertEqual(merge_blockers(make_job('JC1', 10000)), [])

    def test_missing_ups_and_sheet_are_reported(self):
        from planning.merge_engine import merge_blockers
        job = make_job('JC1', 10000, ups=None, print_sheet_size='')
        reasons = merge_blockers(job)
        self.assertTrue(any('UPS missing' in r for r in reasons))
        self.assertTrue(any('Print sheet size missing' in r for r in reasons))

    def test_missing_awc_is_not_a_blocker(self):
        from planning.merge_engine import merge_blockers
        job = make_job('JC-NEWSKU', 10000)   # no SkuRecipe at all, so no AWC
        self.assertEqual(job.awc_no_display, '')
        self.assertEqual(merge_blockers(job), [])

    def test_board_lists_near_miss_job_with_reasons(self):
        make_job('JC1', 10000)
        make_job('JC2', 5000)
        make_job('JC3', 5000)
        make_job('JC-BAD', 5000, ups=None)   # same size/material, missing ups
        response = self.client.get(reverse('planning:merge_board'))
        rows = response.context['blocked_rows']
        self.assertEqual([row['job'].jc_number for row in rows], ['JC-BAD'])
        self.assertTrue(any('UPS missing' in r for r in rows[0]['reasons']))


class MergeLifecycleTests(TestCase):
    """Round 6: artwork -> combined plate -> release-all -> printing."""

    def setUp(self):
        user = get_user_model().objects.create_user('planner7', password='pw12345678')
        profile, _ = UserProfile.objects.get_or_create(user=user)
        profile.role = 'admin'
        profile.save()
        self.user = user
        self.client.force_login(user)

    def _accepted_group(self, n_jobs=3):
        qtys = [10000, 5000, 5000][:n_jobs]
        jobs = [make_job(f'JC{i+1}', q) for i, q in enumerate(qtys)]
        for job in jobs:
            job.application = 'Label'
            job.effective_machine_name and None
            job.save()
        payload = {'job_ids': [j.id for j in jobs]}
        self.client.post(reverse('planning:merge_accept'), payload, follow=True)
        group = MergeGroup.objects.get()
        return group, jobs

    def _make_machine(self):
        from core.models import Machine
        machine, _ = Machine.objects.get_or_create(name='Press-Merge')
        return machine

    def _release_card(self, card):
        card.status = 'production_approved'
        card.save(update_fields=['status'])
        from core.jobcard_service import release_to_production
        release_to_production(card, actor=self.user, reason='test')
        card.refresh_from_db()
        return card

    def test_cancel_allowed_before_combined_plate(self):
        group, jobs = self._accepted_group()
        response = self.client.post(reverse('planning:merge_cancel', args=[group.id]), follow=True)
        group.refresh_from_db()
        self.assertEqual(group.status, 'cancelled')

    def test_legacy_plate_set_no_does_not_block_cancel(self):
        # This is the exact bug the user hit: a repeat SKU's historical
        # plate_set_no text must not read as "plates issued" for the merge.
        group, jobs = self._accepted_group()
        group.lead_job.plate_set_no = '10576'
        group.lead_job.save(update_fields=['plate_set_no'])
        response = self.client.post(reverse('planning:merge_cancel', args=[group.id]), follow=True)
        group.refresh_from_db()
        self.assertEqual(group.status, 'cancelled')

    def test_raise_combined_plate_creates_one_request_on_lead_only(self):
        from printing_plates.models import PlateRequest
        group, jobs = self._accepted_group()
        for job in jobs:
            job.application = 'Label'
            job.machine_name = 'GTO'
            job.front_pass = 1
            job.save()
        # Raising the combined plate requires the layout to be approved first;
        # the approval path itself is covered by MergeLayoutApprovalTests.
        group.status = 'layout_approved'
        group.save(update_fields=['status'])
        response = self.client.post(reverse('planning:merge_raise_plate', args=[group.id]), follow=True)
        group.refresh_from_db()
        requests_on_group_jobs = PlateRequest.objects.filter(
            planning_job__in=[j.id for j in jobs]
        ).exclude(status=PlateRequest.STATUS_ARCHIVED)
        self.assertEqual(requests_on_group_jobs.count(), 1)
        self.assertEqual(requests_on_group_jobs.first().planning_job_id, group.lead_job_id)
        self.assertEqual(group.status, 'artwork_ready')

    def test_raise_combined_plate_is_idempotent(self):
        group, jobs = self._accepted_group()
        for job in jobs:
            job.application = 'Label'
            job.machine_name = 'GTO'
            job.front_pass = 1
            job.save()
        group.status = 'layout_approved'
        group.save(update_fields=['status'])
        self.client.post(reverse('planning:merge_raise_plate', args=[group.id]), follow=True)
        from printing_plates.models import PlateRequest
        count_after_first = PlateRequest.objects.filter(planning_job=group.lead_job_id).exclude(
            status=PlateRequest.STATUS_ARCHIVED
        ).count()
        self.client.post(reverse('planning:merge_raise_plate', args=[group.id]), follow=True)
        count_after_second = PlateRequest.objects.filter(planning_job=group.lead_job_id).exclude(
            status=PlateRequest.STATUS_ARCHIVED
        ).count()
        self.assertEqual(count_after_first, count_after_second)

    def test_cancel_blocked_once_combined_plate_sent(self):
        from printing_plates.models import PlateRequest
        group, jobs = self._accepted_group()
        PlateRequest.objects.create(
            planning_job=group.lead_job, status=PlateRequest.STATUS_SENT,
        )
        response = self.client.post(reverse('planning:merge_cancel', args=[group.id]), follow=True)
        group.refresh_from_db()
        self.assertNotEqual(group.status, 'cancelled')
        self.assertIn('already issued to production', response.content.decode())

    def test_plate_receipt_releases_all_members(self):
        from printing_plates.models import PlateRequest
        group, jobs = self._accepted_group()
        machine = self._make_machine()
        cards = {}
        for job in jobs:
            card = make_card(job)
            card.status = 'production_approved'
            card.save(update_fields=['status'])
            cards[job.id] = card

        plate = PlateRequest.objects.create(
            planning_job=group.lead_job, job_card=cards[group.lead_job_id],
            status=PlateRequest.STATUS_RECEIVED,
        )
        plate.status = PlateRequest.STATUS_AVAILABLE
        plate.save()

        for job in jobs:
            cards[job.id].refresh_from_db()
            self.assertEqual(cards[job.id].workflow_status, 'released')
        group.refresh_from_db()
        self.assertEqual(group.status, 'layout_done')

    def test_lead_printing_blocked_until_followers_released(self):
        from core.models import Production
        group, jobs = self._accepted_group()
        # make_card() defaults to status='released'; followers must NOT be
        # released for this test, so build them at qc_approved instead.
        cards = {}
        for job in jobs:
            card = make_card(job)
            if job.id != group.lead_job_id:
                card.status = 'qc_approved'
                card.save(update_fields=['status'])
            cards[job.id] = card
        self._release_card(cards[group.lead_job_id])
        with self.assertRaises(Exception):
            Production.objects.create(
                job_card=cards[group.lead_job_id], entry_type='printing',
                date='2026-07-22', shift='A', output_sheets=1000,
            )

    def test_lead_printing_succeeds_once_all_released(self):
        from core.models import Production
        group, jobs = self._accepted_group()
        cards = {job.id: make_card(job) for job in jobs}
        for job in jobs:
            self._release_card(cards[job.id])
        entry = Production.objects.create(
            job_card=cards[group.lead_job_id], entry_type='printing',
            date='2026-07-22', shift='A', output_sheets=1000,
        )
        self.assertIsNotNone(entry.pk)
        self.assertEqual(Production.objects.filter(entry_type='printing').count(), len(jobs))


class MergeLayoutApprovalTests(TestCase):
    """Round 7: group-level production gate replaces per-SKU release."""

    def setUp(self):
        user = get_user_model().objects.create_user('planner8', password='pw12345678')
        profile, _ = UserProfile.objects.get_or_create(user=user)
        profile.role = 'admin'
        profile.save()
        self.user = user
        self.client.force_login(user)

    def _complete_master(self, job):
        SkuRecipe.objects.update_or_create(sku=job.sku, defaults={
            'awc_no': f'A-{job.jc_number}', 'material': job.material,
            'color_spec': '4', 'application': 'Label', 'product_type': 'Sticker',
            'size_w_mm': job.size_w_mm, 'size_h_mm': job.size_h_mm,
            'print_sheet_size': job.print_sheet_size, 'purchase_sheet_size': job.purchase_sheet_size,
            'ups': job.ups, 'purchase_sheet_ups': 4, 'die_cutting': 'Kiss', 'print_passes': 1,
            'job_name': job.sku,
        })

    def _accepted_group(self, complete_master=True):
        jobs = [make_job('JC1', 10000), make_job('JC2', 5000), make_job('JC3', 5000)]
        for job in jobs:
            job.application = 'Label'
            job.machine_name = 'GTO'
            job.front_pass = 1
            job.save()
            if complete_master:
                self._complete_master(job)
        self.client.post(reverse('planning:merge_accept'), {'job_ids': [j.id for j in jobs]}, follow=True)
        return MergeGroup.objects.get(), jobs

    def _approve(self, group):
        return self.client.post(
            reverse('planning:merge_approve_layout', args=[group.id]),
            {'combined_wastage': 200, 'material_origin': 'local'}, follow=True,
        )

    def test_approve_refuses_when_master_incomplete(self):
        from planning.services import approve_merge_layout
        from django.core.exceptions import ValidationError
        group, jobs = self._accepted_group(complete_master=False)
        with self.assertRaises(ValidationError):
            approve_merge_layout(group, actor=self.user, combined_wastage=200, material_origin='local')
        group.refresh_from_db()
        self.assertEqual(group.status, 'accepted')

    def test_approve_requires_wastage_and_material_origin(self):
        from planning.services import approve_merge_layout
        from django.core.exceptions import ValidationError
        group, jobs = self._accepted_group()
        with self.assertRaises(ValidationError):
            approve_merge_layout(group, actor=self.user)  # no wastage/origin
        group.refresh_from_db()
        self.assertEqual(group.status, 'accepted')

    def test_approve_divides_wastage_and_reflects_origin(self):
        group, jobs = self._accepted_group()
        self._approve(group)
        members = [i.planning_job for i in group.items.select_related('planning_job')]
        self.assertEqual(sum(m.wastage_sheets for m in members), 200)
        self.assertTrue(all(m.purchase_material_origin == 'local' for m in members))

    def test_approve_hands_off_to_designer(self):
        from printing_plates.models import PlateRequest
        group, jobs = self._accepted_group()
        self._approve(group)
        group.refresh_from_db()
        # Approval production-approves members AND auto-creates the combined plate
        # into the designer's queue; the group is handed to design.
        self.assertEqual(group.status, 'artwork_requested')
        for item in group.items.select_related('planning_job__job_card'):
            self.assertEqual(item.planning_job.job_card.workflow_status, 'production_approved')
        plate = PlateRequest.objects.filter(planning_job=group.lead_job_id).exclude(
            status=PlateRequest.STATUS_ARCHIVED
        ).first()
        self.assertIsNotNone(plate)
        self.assertEqual(plate.awc_no, group.artwork_code)

    def test_no_combined_plate_before_approval(self):
        from printing_plates.models import PlateRequest
        group, jobs = self._accepted_group()
        self.assertFalse(
            PlateRequest.objects.filter(planning_job=group.lead_job_id).exclude(
                status=PlateRequest.STATUS_ARCHIVED
            ).exists()
        )

    def test_plate_receipt_releases_all_after_approval(self):
        from printing_plates.models import PlateRequest
        group, jobs = self._accepted_group()
        self._approve(group)  # auto-creates the plate
        plate = PlateRequest.objects.filter(planning_job=group.lead_job_id).exclude(
            status=PlateRequest.STATUS_ARCHIVED
        ).first()
        plate.status = PlateRequest.STATUS_AVAILABLE
        plate.save()
        for item in group.items.select_related('planning_job__job_card'):
            item.planning_job.job_card.refresh_from_db()
            self.assertEqual(item.planning_job.job_card.workflow_status, 'released')

    def test_group_advances_to_artwork_ready_on_send(self):
        from printing_plates.models import PlateRequest
        group, jobs = self._accepted_group()
        self._approve(group)
        plate = PlateRequest.objects.filter(planning_job=group.lead_job_id).exclude(
            status=PlateRequest.STATUS_ARCHIVED
        ).first()
        plate.status = PlateRequest.STATUS_SENT
        plate.save()
        group.refresh_from_db()
        self.assertEqual(group.status, 'artwork_ready')

    def test_merge_context_carries_combined_run(self):
        from planning.services import build_job_card_merge_context
        group, jobs = self._accepted_group()
        ctx = build_job_card_merge_context(group.lead_job)
        self.assertEqual(ctx['run_sheets'], group.run_sheets)
        self.assertEqual(ctx['total_sheet_ups'], group.total_sheet_ups)
        self.assertEqual(ctx['combined_impressions'], group.run_sheets * 1)

    def test_lead_card_is_normal_with_watermark(self):
        """Round 9: the printed card is the familiar single-SKU card + watermark;
        combined-run figures live only on the separate Combined Layout Sheet."""
        from django.template.loader import render_to_string
        from planning.services import build_job_card_merge_context
        group, jobs = self._accepted_group()
        lead = group.lead_job
        html = render_to_string('Job Card.html', {
            'job': lead, 'recipe': lead.sku_recipe,
            'merge': build_job_card_merge_context(lead),
        })
        self.assertNotIn('(combined run)', html)          # card keeps standalone numbers
        self.assertIn('DO NOT PRINT SEPARATELY', html)    # watermark instead
        sheet = render_to_string('planning/planning_merge_combined_sheet.html', {'group': group})
        self.assertIn(str(group.run_sheets), sheet)       # combined run lives here

    def test_approved_lead_can_print_and_split_lands_on_all(self):
        from core.models import Production
        from printing_plates.models import PlateRequest
        group, jobs = self._accepted_group()
        self._approve(group)  # auto-creates the combined plate on the lead
        plate = PlateRequest.objects.filter(planning_job=group.lead_job_id).exclude(
            status=PlateRequest.STATUS_ARCHIVED
        ).first()
        plate.status = PlateRequest.STATUS_AVAILABLE
        plate.save()  # releases all member cards

        lead_card = group.lead_job.job_card
        lead_card.refresh_from_db()
        entry = Production.objects.create(
            job_card=lead_card, entry_type='printing',
            date='2026-07-22', shift='A', output_sheets=group.run_sheets,
        )
        self.assertEqual(Production.objects.filter(entry_type='printing').count(), len(jobs))
        # Lead recounts at its combined-sheet ups (2), not standalone 14.
        entry.refresh_from_db()
        self.assertEqual(entry.merge_allocated_ups, 2)


class MergedJobsQueueVisibilityTests(TestCase):
    """Round 10: merged jobs stay hidden by default but are always findable."""

    def setUp(self):
        user = get_user_model().objects.create_user('planner10', password='pw12345678')
        profile, _ = UserProfile.objects.get_or_create(user=user)
        profile.role = 'admin'
        profile.save()
        self.user = user
        self.client.force_login(user)
        jobs = [make_job('JC1', 10000), make_job('JC2', 5000), make_job('JC3', 5000)]
        self.client.post(reverse('planning:merge_accept'), {'job_ids': [j.id for j in jobs]}, follow=True)
        self.group = MergeGroup.objects.get()
        self.member_jc = self.group.items.first().planning_job.jc_number

    def test_default_queue_hides_merged_jobs(self):
        response = self.client.get(reverse('planning:jobs'))
        self.assertNotContains(response, self.member_jc)
        self.assertContains(response, 'Merged (3)')
        self.assertContains(response, 'hidden here to keep the queue clean')

    def test_search_surfaces_a_merged_job(self):
        response = self.client.get(reverse('planning:jobs'), {'q': self.member_jc})
        self.assertContains(response, self.member_jc)
        self.assertContains(response, 'Merged \xb7')  # the row badge

    def test_merged_chip_lists_only_merged_jobs(self):
        response = self.client.get(reverse('planning:jobs'), {'merged': '1'})
        self.assertContains(response, self.member_jc)
        self.assertContains(response, 'Showing')
        self.assertContains(response, self.group.code)
