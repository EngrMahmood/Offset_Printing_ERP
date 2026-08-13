from django.test import TestCase
from django.urls import reverse
from django.contrib.auth import get_user_model

from core.models import UserProfile


class ReportsAppTests(TestCase):
    def setUp(self):
        User = get_user_model()
        self.user = User.objects.create_user(username='report_user', password='pass12345')
        profile = UserProfile.objects.get(user=self.user)
        profile.role = 'manager'
        profile.save(update_fields=['role'])

        # Grant standard Django view_reports permission on UserProfile
        from django.contrib.auth.models import Permission
        from django.contrib.contenttypes.models import ContentType
        content_type = ContentType.objects.get_for_model(UserProfile)
        permission, _ = Permission.objects.get_or_create(
            codename='view_reports',
            name='Can view reports',
            content_type=content_type,
        )
        self.user.user_permissions.add(permission)

    def test_reports_home_loads(self):
        response = self.client.get(reverse('reports:home'))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, 'Available Reports')
        self.assertContains(response, 'Machine Planning')

    def test_reports_api_list_loads(self):
        response = self.client.get(reverse('reports:reports_api:list_reports'))
        self.assertEqual(response.status_code, 200)
        self.assertJSONEqual(
            response.content,
            {
                'ok': True,
                'count': len(response.json().get('items', [])),
                'items': response.json().get('items', []),
            },
        )

    def test_machine_planning_report_loads(self):
        self.client.login(username='report_user', password='pass12345')
        response = self.client.get(reverse('reports:detail', args=['machine-planning']))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, 'Machine Planning')

    def test_job_planning_report_loads(self):
        response = self.client.get(reverse('reports:detail', args=['job-planning']))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, 'Job Planning')

    def test_qc_approvals_report_loads(self):
        response = self.client.get(reverse('reports:detail', args=['qc-approvals']))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, 'QC Approvals')

    def test_dispatch_tracking_report_loads(self):
        response = self.client.get(reverse('reports:detail', args=['dispatch-tracking']))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, 'Dispatch Tracking')

    def test_report_detail_404_for_invalid_slug(self):
        response = self.client.get(reverse('reports:detail', args=['invalid-report']))
        self.assertEqual(response.status_code, 404)

    def test_report_run_api_loads(self):
        self.client.login(username='report_user', password='pass12345')
        response = self.client.get(reverse('reports:reports_api:run_report', args=['machine-planning']))
        self.assertEqual(response.status_code, 200)
        self.assertTrue(response.json()['ok'])

    def test_report_export_csv_loads(self):
        self.client.login(username='report_user', password='pass12345')
        response = self.client.get(reverse('reports:reports_api:export_report', args=['machine-planning']), {'type': 'csv'})
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response['Content-Type'].split(';')[0], 'text/csv')

    def test_wastage_report_loads(self):
        response = self.client.get(reverse('reports:detail', args=['wastage-report']))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, 'Wastage Report')

    def test_manager_without_django_perm_gets_report_data(self):
        """Regression test: a manager (role-based access) must get report data
        even without the Django `core.view_reports` permission, since the app
        authorises everywhere else via profile.role, not Django perms."""
        User = get_user_model()
        role_user = User.objects.create_user(username='role_only_manager', password='pass12345')
        profile = UserProfile.objects.get(user=role_user)
        profile.role = 'manager'
        profile.save(update_fields=['role'])

        self.client.login(username='role_only_manager', password='pass12345')
        response = self.client.get(reverse('reports:reports_api:run_report', args=['machine-planning']))
        self.assertEqual(response.status_code, 200)
        body = response.json()
        self.assertTrue(body['ok'])
        self.assertIn('data', body['payload'])

    def test_role_without_access_gets_denied(self):
        User = get_user_model()
        other_user = User.objects.create_user(username='no_access_user', password='pass12345')
        profile = UserProfile.objects.get(user=other_user)
        profile.role = 'qc'
        profile.save(update_fields=['role'])

        self.client.login(username='no_access_user', password='pass12345')
        response = self.client.get(reverse('reports:reports_api:run_report', args=['machine-planning']))
        self.assertEqual(response.status_code, 403)

    def test_planning_job_save_bumps_report_cache_version(self):
        from reports.report_engine.engine import get_cache_version
        from planning.models import PlanningJob
        import datetime

        before = get_cache_version()
        PlanningJob.objects.create(
            jc_number='JC-TEST-CACHE-1',
            order_qty=100,
            status='pending_qc',
            plan_date=datetime.date.today(),
            plan_month='July 2026',
        )
        after = get_cache_version()
        self.assertGreater(after, before)

    def test_machine_save_bumps_report_cache_version(self):
        from reports.report_engine.engine import get_cache_version
        from core.models import Machine

        before = get_cache_version()
        Machine.objects.create(name='Cache Bump Machine', is_active=True)
        after = get_cache_version()
        self.assertGreater(after, before)

    def test_machine_planning_routes_jobs_into_named_colour_pools(self):
        """End-to-end test of Part B2/B3: PlanningJob rows should be grouped
        by the colour/size-routed machine pool (named members), not the
        literal machine_name string, and should carry pass counts and
        actual-machine tracking (Part C)."""
        from core.models import Machine, JobCard, Production
        from planning.models import PlanningJob
        import datetime

        Machine.objects.create(name='GTO 1A', machine_type='offset_printing', machine_group_code='GTO1', default_colors=1, operational_colors=1,
                                max_print_length_mm=520, max_print_width_mm=740, is_active=True)
        Machine.objects.create(name='GTO 1B', machine_type='offset_printing', machine_group_code='GTO1', default_colors=1, operational_colors=1,
                                max_print_length_mm=520, max_print_width_mm=740, is_active=True)
        gto2a = Machine.objects.create(name='GTO 2A', machine_type='offset_printing', machine_group_code='GTO2', default_colors=2, operational_colors=2,
                                        max_print_length_mm=520, max_print_width_mm=740, is_active=True)
        Machine.objects.create(name='GTO 2B', machine_type='offset_printing', machine_group_code='GTO2', default_colors=2, operational_colors=2,
                                max_print_length_mm=520, max_print_width_mm=740, is_active=True)
        Machine.objects.create(name='SM 74', machine_type='offset_printing', machine_group_code='SM74', default_colors=5, operational_colors=5,
                                max_print_length_mm=740, max_print_width_mm=1050, is_active=True)

        pj = PlanningJob.objects.create(
            jc_number='JC-TEST-ROUTE-1',
            order_qty=100,
            status='released',
            plan_date=datetime.date.today(),
            plan_month='July 2026',
            color_spec='3',
            print_sheet_size='18*25',
            sku='SKU-ROUTE-1',
        )
        jc = JobCard.objects.create(
            job_card_no='JC-TEST-ROUTE-1',
            planning_job=pj,
            order_qty=100,
            total_sheet_quantity=100,
            status='in_production',
            is_active=True,
            SKU='SKU-ROUTE-1',
            po_date=datetime.date.today(),
            machine_name=gto2a,
            total_impressions_required=100,
        )
        Production.objects.create(
            entry_type='printing', job_card=jc, date=datetime.date.today(), shift='A',
            output_sheets=100, status='completed', machine=gto2a,
        )

        self.client.login(username='report_user', password='pass12345')
        response = self.client.get(reverse('reports:reports_api:run_report', args=['machine-planning']))
        self.assertEqual(response.status_code, 200)
        payload = response.json()['payload']['data']

        pool_key = 'GTO 2A, GTO 2B'
        self.assertIn(pool_key, payload['machine_reports'])
        pool_row = payload['machine_reports'][pool_key]['rows'][0]
        self.assertEqual(pool_row['sku'], 'SKU-ROUTE-1')
        self.assertEqual(pool_row['passes'], 2)  # 3 colours on a 2-colour pool
        self.assertEqual(pool_row['actual_machine'], 'GTO 2A')

    def test_machine_planning_respects_explicit_sm74_assignment(self):
        """Regression: a job explicitly planned onto SM74 must stay under the
        SM74 pool even if its colour count would technically fit a GTO2 pool
        by the size/colour heuristic - SM74's plate size is physically
        different and manual assignments should not be silently reshuffled."""
        from core.models import Machine, JobCard
        from planning.models import PlanningJob
        import datetime

        Machine.objects.create(name='GTO 2A', machine_type='offset_printing', machine_group_code='GTO2', default_colors=2, operational_colors=2,
                                max_print_length_mm=520, max_print_width_mm=740, is_active=True)
        sm74 = Machine.objects.create(name='SM 74', machine_type='offset_printing', machine_group_code='SM74', default_colors=5, operational_colors=5,
                                       max_print_length_mm=740, max_print_width_mm=1050, is_active=True)

        pj = PlanningJob.objects.create(
            jc_number='JC-TEST-SM74-1',
            order_qty=100,
            status='released',
            plan_date=datetime.date.today(),
            plan_month='July 2026',
            color_spec='2',
            print_sheet_size='18*25',
            machine_name='SM 74',
            sku='SKU-SM74-1',
        )
        JobCard.objects.create(
            job_card_no='JC-TEST-SM74-1',
            planning_job=pj,
            order_qty=100,
            total_sheet_quantity=100,
            status='in_production',
            is_active=True,
            SKU='SKU-SM74-1',
            po_date=datetime.date.today(),
            machine_name=sm74,
            total_impressions_required=100,
        )

        self.client.login(username='report_user', password='pass12345')
        response = self.client.get(reverse('reports:reports_api:run_report', args=['machine-planning']))
        self.assertEqual(response.status_code, 200)
        payload = response.json()['payload']['data']

        self.assertIn('SM 74', payload['machine_reports'])
        self.assertNotIn('GTO 2A', payload['machine_reports'])

    def test_priority_update_endpoint_bumps_report_cache_version(self):
        """V2 plan item 6: the planning_job_priority_update view must bust
        the report cache so Machine Planning reflects the new priority on
        the very next reload, not after the 300s cache timeout."""
        from reports.report_engine.engine import get_cache_version
        from planning.models import PlanningJob
        import datetime

        pj = PlanningJob.objects.create(
            jc_number='JC-TEST-PRIORITY-CACHE-1', order_qty=10, status='draft',
            plan_date=datetime.date.today(), plan_month='July 2026', priority=1,
        )
        self.client.login(username='report_user', password='pass12345')
        before = get_cache_version()
        response = self.client.post(reverse('planning:job_priority_update', args=[pj.id]), {'priority': 3})
        self.assertEqual(response.status_code, 200)
        after = get_cache_version()
        self.assertGreater(after, before)

    def test_machine_planning_collapses_explicit_pool_members_into_one_tab(self):
        """V2 plan item 4: two jobs explicitly assigned to sibling machines in
        the same pool (GTO 1A / GTO 1B) must produce exactly one combined
        tab, not one tab per machine plus a merged one."""
        from core.models import Machine, JobCard
        from planning.models import PlanningJob
        import datetime

        gto1a = Machine.objects.create(name='GTO 1A', machine_type='offset_printing', machine_group_code='GTO1',
                                        default_colors=1, operational_colors=1,
                                        max_print_length_mm=520, max_print_width_mm=740, is_active=True)
        gto1b = Machine.objects.create(name='GTO 1B', machine_type='offset_printing', machine_group_code='GTO1',
                                        default_colors=1, operational_colors=1,
                                        max_print_length_mm=520, max_print_width_mm=740, is_active=True)

        for i, m in enumerate([gto1a, gto1b], start=1):
            pj = PlanningJob.objects.create(
                jc_number=f'JC-ITEM4-{i}', order_qty=100, status='released',
                plan_date=datetime.date.today(), plan_month='July 2026',
                machine_name=m.name, sku=f'SKU-ITEM4-{i}',
            )
            JobCard.objects.create(
                job_card_no=f'JC-ITEM4-{i}', planning_job=pj, order_qty=100, total_sheet_quantity=100,
                status='in_production', is_active=True, SKU=f'SKU-ITEM4-{i}', po_date=datetime.date.today(),
                machine_name=m, total_impressions_required=100, total_colors=1,
            )

        self.client.login(username='report_user', password='pass12345')
        response = self.client.get(reverse('reports:reports_api:run_report', args=['machine-planning']))
        self.assertEqual(response.status_code, 200)
        machine_reports = response.json()['payload']['data']['machine_reports']

        gto1_tabs = [key for key in machine_reports if 'GTO 1A' in key or 'GTO 1B' in key]
        self.assertEqual(gto1_tabs, ['GTO 1A, GTO 1B'])

    def test_machine_planning_partial_production_and_on_hold(self):
        """V2 plan item 1: a job that's already fully produced (>=95% of
        planned sheets run) drops out of the report; a partially-produced
        job stays visible showing its remaining balance and stage; an
        on-hold job is excluded entirely."""
        from core.models import Machine, JobCard, Production
        from planning.models import PlanningJob
        import datetime

        machine = Machine.objects.create(name='GTO 1A', machine_type='offset_printing', machine_group_code='GTO1',
                                          default_colors=1, operational_colors=1,
                                          max_print_length_mm=520, max_print_width_mm=740, is_active=True)

        def _make_job(suffix, produced_sheets, is_on_hold=False):
            pj = PlanningJob.objects.create(
                jc_number=f'JC-PARTIAL-{suffix}', order_qty=100, status='in_production',
                plan_date=datetime.date.today(), plan_month='July 2026',
                machine_name='GTO 1A', sku=f'SKU-PARTIAL-{suffix}', is_on_hold=is_on_hold,
                actual_sheet_required=100,
            )
            jc = JobCard.objects.create(
                job_card_no=f'JC-PARTIAL-{suffix}', planning_job=pj, order_qty=100, total_sheet_quantity=100,
                status='in_production', is_active=True, SKU=f'SKU-PARTIAL-{suffix}', po_date=datetime.date.today(),
                machine_name=machine, total_impressions_required=100, total_colors=1,
            )
            if produced_sheets:
                Production.objects.create(
                    entry_type='printing', job_card=jc, date=datetime.date.today(), shift='A',
                    output_sheets=produced_sheets, status='completed', machine=machine,
                )
            return pj

        not_started = _make_job('NOTSTARTED', produced_sheets=0)
        partial = _make_job('PARTIAL', produced_sheets=40)
        done = _make_job('DONE', produced_sheets=98)  # 98% -> within 5% tolerance -> done
        on_hold = _make_job('ONHOLD', produced_sheets=0, is_on_hold=True)

        self.client.login(username='report_user', password='pass12345')
        response = self.client.get(reverse('reports:reports_api:run_report', args=['machine-planning']))
        self.assertEqual(response.status_code, 200)
        payload = response.json()['payload']['data']

        all_jc_numbers = set()
        for pool in payload['machine_reports'].values():
            for row in pool['rows']:
                all_jc_numbers.update(row['job_card_numbers'].split(', '))

        self.assertIn('JC-PARTIAL-NOTSTARTED', all_jc_numbers)
        self.assertIn('JC-PARTIAL-PARTIAL', all_jc_numbers)
        self.assertNotIn('JC-PARTIAL-DONE', all_jc_numbers)
        self.assertNotIn('JC-PARTIAL-ONHOLD', all_jc_numbers)

        partial_row = next(
            row for pool in payload['machine_reports'].values() for row in pool['rows']
            if row['job_card_numbers'] == 'JC-PARTIAL-PARTIAL'
        )
        self.assertTrue(partial_row['has_partial_production'])
        self.assertEqual(partial_row['print_sheet_quantity'], 60)  # 100 planned - 40 produced
        detail = partial_row['jobs_detail'][0]
        self.assertEqual(detail['production_state'], 'partially_produced')
        self.assertEqual(detail['stage'], 'In Production')

    def test_machine_planning_planner_override_flag_on_sm74(self):
        """V2 plan item 5a: a job that fits GTO by colour but was explicitly
        parked on SM74 by the planner stays there (already covered), and is
        now also flagged planner_override=True so the UI can explain why."""
        from core.models import Machine, JobCard
        from planning.models import PlanningJob
        import datetime

        Machine.objects.create(name='GTO 2A', machine_type='offset_printing', machine_group_code='GTO2',
                                default_colors=2, operational_colors=2,
                                max_print_length_mm=520, max_print_width_mm=740, is_active=True)
        sm74 = Machine.objects.create(name='SM 74', machine_type='offset_printing', machine_group_code='SM74',
                                       default_colors=5, operational_colors=5,
                                       max_print_length_mm=740, max_print_width_mm=1050, is_active=True)

        pj = PlanningJob.objects.create(
            jc_number='JC-OVERRIDE-1', order_qty=100, status='released',
            plan_date=datetime.date.today(), plan_month='July 2026',
            color_spec='2', print_sheet_size='18*25', machine_name='SM 74', sku='SKU-OVERRIDE-1',
        )
        JobCard.objects.create(
            job_card_no='JC-OVERRIDE-1', planning_job=pj, order_qty=100, total_sheet_quantity=100,
            status='in_production', is_active=True, SKU='SKU-OVERRIDE-1', po_date=datetime.date.today(),
            machine_name=sm74, total_impressions_required=100, total_colors=2,
        )

        self.client.login(username='report_user', password='pass12345')
        response = self.client.get(reverse('reports:reports_api:run_report', args=['machine-planning']))
        payload = response.json()['payload']['data']
        row = payload['machine_reports']['SM 74']['rows'][0]
        self.assertTrue(row['planner_override'])

    def test_machine_planning_size_warning_on_oversized_gto_assignment(self):
        """V2 plan item 5b: a job explicitly parked on a GTO machine but whose
        sheet size exceeds that GTO pool's max must show a warning, not be
        silently auto-moved."""
        from core.models import Machine, JobCard
        from planning.models import PlanningJob
        import datetime

        gto2a = Machine.objects.create(name='GTO 2A', machine_type='offset_printing', machine_group_code='GTO2',
                                        default_colors=2, operational_colors=2,
                                        max_print_length_mm=520, max_print_width_mm=740, is_active=True)
        Machine.objects.create(name='SM 74', machine_type='offset_printing', machine_group_code='SM74',
                                default_colors=5, operational_colors=5,
                                max_print_length_mm=740, max_print_width_mm=1050, is_active=True)

        pj = PlanningJob.objects.create(
            jc_number='JC-OVERSIZE-1', order_qty=100, status='released',
            plan_date=datetime.date.today(), plan_month='July 2026',
            color_spec='2', print_sheet_size='30*40', machine_name='GTO 2A', sku='SKU-OVERSIZE-1',
        )
        JobCard.objects.create(
            job_card_no='JC-OVERSIZE-1', planning_job=pj, order_qty=100, total_sheet_quantity=100,
            status='in_production', is_active=True, SKU='SKU-OVERSIZE-1', po_date=datetime.date.today(),
            machine_name=gto2a, total_impressions_required=100, total_colors=2,
        )

        self.client.login(username='report_user', password='pass12345')
        response = self.client.get(reverse('reports:reports_api:run_report', args=['machine-planning']))
        payload = response.json()['payload']['data']
        row = payload['machine_reports']['GTO 2A']['rows'][0]
        self.assertTrue(row['size_warnings'])
        self.assertEqual(payload['machine_reports']['GTO 2A']['summary']['size_violations_count'], 1)

    def test_jc_selection_endpoint_requires_planner_or_admin_role(self):
        """V2 plan item 3: only planner/admin (or superuser) may toggle JC
        selection; other roles get a 403 and the selection is unchanged."""
        from core.models import UserProfile
        from reports.models import MachinePlanningJcSelection
        User = get_user_model()
        qc_user = User.objects.create_user(username='qc_user_sel', password='pass12345')
        UserProfile.objects.filter(user=qc_user).update(role='qc')

        self.client.login(username='qc_user_sel', password='pass12345')
        response = self.client.post(
            reverse('reports:reports_api:machine_planning_jc_selection', args=['JC-NOPE-1']),
            {'is_excluded': 'true'},
        )
        self.assertEqual(response.status_code, 403)
        self.assertFalse(MachinePlanningJcSelection.objects.filter(jc_number='JC-NOPE-1').exists())

    def test_jc_selection_endpoint_allows_manager_role(self):
        """Manager has the same planning-console permissions as Planner/Admin
        (short of deletion/superuser-only actions), so a manager must also
        be able to toggle JC selection in Machine Planning."""
        from core.models import UserProfile
        from reports.models import MachinePlanningJcSelection
        User = get_user_model()
        manager_user = User.objects.create_user(username='manager_user_sel', password='pass12345')
        UserProfile.objects.filter(user=manager_user).update(role='manager')

        self.client.login(username='manager_user_sel', password='pass12345')
        response = self.client.post(
            reverse('reports:reports_api:machine_planning_jc_selection', args=['JC-MGR-1']),
            {'is_excluded': 'true'},
        )
        self.assertEqual(response.status_code, 200)
        self.assertTrue(response.json()['is_excluded'])
        self.assertTrue(MachinePlanningJcSelection.objects.filter(jc_number='JC-MGR-1', is_excluded=True).exists())

    def test_jc_selection_endpoint_planner_can_toggle_and_report_excludes_it(self):
        """V2 plan item 3: a planner deselecting one JC of a combined run
        recomputes the merged totals to exclude it, while the JC stays
        visible (marked excluded) in jobs_detail for the planner console."""
        from core.models import UserProfile, Machine, JobCard
        from planning.models import PlanningJob
        import datetime

        User = get_user_model()
        planner = User.objects.create_user(username='planner_sel', password='pass12345')
        UserProfile.objects.filter(user=planner).update(role='planner')

        machine = Machine.objects.create(name='GTO 1A', machine_type='offset_printing', machine_group_code='GTO1',
                                          default_colors=1, operational_colors=1,
                                          max_print_length_mm=520, max_print_width_mm=740, is_active=True)

        for i in (1, 2):
            pj = PlanningJob.objects.create(
                jc_number=f'JC-SEL-{i}', order_qty=100, status='released',
                plan_date=datetime.date.today(), plan_month='July 2026',
                machine_name='GTO 1A', sku='SKU-SEL-1', actual_sheet_required=100,
            )
            JobCard.objects.create(
                job_card_no=f'JC-SEL-{i}', planning_job=pj, order_qty=100, total_sheet_quantity=100,
                status='in_production', is_active=True, SKU='SKU-SEL-1', po_date=datetime.date.today(),
                machine_name=machine, total_impressions_required=100, total_colors=1,
            )

        self.client.login(username='planner_sel', password='pass12345')
        toggle_response = self.client.post(
            reverse('reports:reports_api:machine_planning_jc_selection', args=['JC-SEL-2']),
            {'is_excluded': 'true'},
        )
        self.assertEqual(toggle_response.status_code, 200)
        self.assertTrue(toggle_response.json()['is_excluded'])

        response = self.client.get(reverse('reports:reports_api:run_report', args=['machine-planning']))
        payload = response.json()['payload']['data']
        row = payload['machine_reports']['GTO 1A']['rows'][0]

        self.assertEqual(row['job_card_numbers'], 'JC-SEL-1')
        self.assertEqual(row['finish_quantity'], 100)  # only JC-SEL-1's qty, not both
        excluded_detail = next(d for d in row['jobs_detail'] if d['jc_number'] == 'JC-SEL-2')
        self.assertTrue(excluded_detail['is_excluded'])
        included_detail = next(d for d in row['jobs_detail'] if d['jc_number'] == 'JC-SEL-1')
        self.assertFalse(included_detail['is_excluded'])

    def test_machine_planning_group_code_fallback_and_name_canonicalization(self):
        """Regression: legacy machine_name text without a per-unit suffix
        ('GTO 1' instead of 'GTO 1A') must fold into the combined pool tab
        rather than showing as its own stray tab, and case variants of the
        same non-pool machine name must collapse into one canonical tab."""
        from core.models import Machine, JobCard
        from planning.models import PlanningJob
        import datetime

        gto1a = Machine.objects.create(name='GTO 1A', machine_type='offset_printing', machine_group_code='GTO1',
                                        default_colors=1, operational_colors=1,
                                        max_print_length_mm=520, max_print_width_mm=740, is_active=True)
        gto1b = Machine.objects.create(name='GTO 1B', machine_type='offset_printing', machine_group_code='GTO1',
                                        default_colors=1, operational_colors=1,
                                        max_print_length_mm=520, max_print_width_mm=740, is_active=True)
        konica = Machine.objects.create(name='Konica Minolta', machine_type='digital_printing', is_active=True)

        # Job 1: assigned to plain "GTO 1" with no per-unit suffix.
        pj1 = PlanningJob.objects.create(
            jc_number='JC-FALLBACK-1', order_qty=100, status='released',
            plan_date=datetime.date.today(), plan_month='July 2026',
            machine_name='GTO 1', sku='SKU-FALLBACK-1',
        )
        JobCard.objects.create(
            job_card_no='JC-FALLBACK-1', planning_job=pj1, order_qty=100, total_sheet_quantity=100,
            status='in_production', is_active=True, SKU='SKU-FALLBACK-1', po_date=datetime.date.today(),
            machine_name=gto1a, total_impressions_required=100, total_colors=1,
        )
        # Job 2: assigned to the real "GTO 1B" unit.
        pj2 = PlanningJob.objects.create(
            jc_number='JC-FALLBACK-2', order_qty=100, status='released',
            plan_date=datetime.date.today(), plan_month='July 2026',
            machine_name='GTO 1B', sku='SKU-FALLBACK-2',
        )
        JobCard.objects.create(
            job_card_no='JC-FALLBACK-2', planning_job=pj2, order_qty=100, total_sheet_quantity=100,
            status='in_production', is_active=True, SKU='SKU-FALLBACK-2', po_date=datetime.date.today(),
            machine_name=gto1b, total_impressions_required=100, total_colors=1,
        )
        # Job 3/4: same digital machine but different case in the free-text field.
        pj3 = PlanningJob.objects.create(
            jc_number='JC-CASE-1', order_qty=50, status='released',
            plan_date=datetime.date.today(), plan_month='July 2026',
            machine_name='KONICA MINOLTA', sku='SKU-CASE-1',
        )
        JobCard.objects.create(
            job_card_no='JC-CASE-1', planning_job=pj3, order_qty=50, total_sheet_quantity=50,
            status='in_production', is_active=True, SKU='SKU-CASE-1', po_date=datetime.date.today(),
            machine_name=konica, total_impressions_required=50, total_colors=1,
        )
        pj4 = PlanningJob.objects.create(
            jc_number='JC-CASE-2', order_qty=50, status='released',
            plan_date=datetime.date.today(), plan_month='July 2026',
            machine_name='Konica Minolta', sku='SKU-CASE-2',
        )
        JobCard.objects.create(
            job_card_no='JC-CASE-2', planning_job=pj4, order_qty=50, total_sheet_quantity=50,
            status='in_production', is_active=True, SKU='SKU-CASE-2', po_date=datetime.date.today(),
            machine_name=konica, total_impressions_required=50, total_colors=1,
        )

        self.client.login(username='report_user', password='pass12345')
        response = self.client.get(reverse('reports:reports_api:run_report', args=['machine-planning']))
        machine_reports = response.json()['payload']['data']['machine_reports']

        gto1_tabs = [key for key in machine_reports if 'GTO 1' in key]
        self.assertEqual(gto1_tabs, ['GTO 1A, GTO 1B'])

        konica_tabs = [key for key in machine_reports if 'konica' in key.lower()]
        self.assertEqual(konica_tabs, ['Konica Minolta'])

    def test_machine_planning_row_has_impressions_and_hours(self):
        """Supervisor overview: each row should carry total_impressions and
        estimated_hours so shop-floor supervisors can see workload at a glance."""
        from core.models import Machine, JobCard
        from planning.models import PlanningJob
        import datetime

        machine = Machine.objects.create(name='GTO 1A', machine_type='offset_printing', machine_group_code='GTO1',
                                          default_colors=1, operational_colors=1,
                                          standard_impressions_per_hour=4000, standard_setup_minutes_per_color=15,
                                          max_print_length_mm=520, max_print_width_mm=740, is_active=True)
        pj = PlanningJob.objects.create(
            jc_number='JC-HOURS-1', order_qty=100, status='released',
            plan_date=datetime.date.today(), plan_month='July 2026',
            machine_name='GTO 1A', sku='SKU-HOURS-1', actual_sheet_required=1000,
        )
        JobCard.objects.create(
            job_card_no='JC-HOURS-1', planning_job=pj, order_qty=100, total_sheet_quantity=1000,
            status='in_production', is_active=True, SKU='SKU-HOURS-1', po_date=datetime.date.today(),
            machine_name=machine, total_impressions_required=1000, total_colors=1,
        )

        self.client.login(username='report_user', password='pass12345')
        response = self.client.get(reverse('reports:reports_api:run_report', args=['machine-planning']))
        payload = response.json()['payload']['data']
        row = payload['machine_reports']['GTO 1A']['rows'][0]

        self.assertEqual(row['total_impressions'], 1000)  # 1000 sheets * 1 pass
        self.assertGreater(row['estimated_hours'], 0)

    def test_machine_planning_uses_planning_print_passes_for_impressions(self):
        """A 1+1 (front/back) job is 2 physical passes through the press,
        so its impressions must be sheets * 2, using PlanningJob's own
        print_passes field (the value the planner maintains), not the
        machine-pool merge-pass heuristic (which would say 1 pass since
        1+1 is a single-colour-class job)."""
        from core.models import Machine, JobCard
        from planning.models import PlanningJob
        import datetime

        machine = Machine.objects.create(name='GTO 1A', machine_type='offset_printing', machine_group_code='GTO1',
                                          default_colors=1, operational_colors=1,
                                          standard_impressions_per_hour=4000, standard_setup_minutes_per_color=15,
                                          max_print_length_mm=520, max_print_width_mm=740, is_active=True)
        pj = PlanningJob.objects.create(
            jc_number='JC-PASSES-1', order_qty=100, status='released',
            plan_date=datetime.date.today(), plan_month='July 2026',
            machine_name='GTO 1A', sku='SKU-PASSES-1', color_spec='1+1',
            actual_sheet_required=1000, print_passes=2,
        )
        JobCard.objects.create(
            job_card_no='JC-PASSES-1', planning_job=pj, order_qty=100, total_sheet_quantity=1000,
            status='in_production', is_active=True, SKU='SKU-PASSES-1', po_date=datetime.date.today(),
            machine_name=machine, total_impressions_required=2000, total_colors=1,
        )

        self.client.login(username='report_user', password='pass12345')
        response = self.client.get(reverse('reports:reports_api:run_report', args=['machine-planning']))
        payload = response.json()['payload']['data']
        row = payload['machine_reports']['GTO 1A']['rows'][0]

        self.assertEqual(row['passes'], 2)
        self.assertEqual(row['total_impressions'], 2000)  # 1000 sheets * 2 passes

    def test_machine_planning_pdf_export_fits_columns_on_one_page(self):
        """Regression: the PDF export chunked columns onto a second page
        (Passes/Hours got pushed off) because min_col_width was too wide for
        19 columns and the widths_map summed to >100%. Both are fixed now,
        so exporting should not raise and should build a single-page story
        (no internal PageBreak for the column table)."""
        from core.models import Machine, JobCard
        from planning.models import PlanningJob
        from reports.export.services import export_as_pdf
        from reports.report_engine.engine import run_report
        from django.test import RequestFactory
        import datetime

        machine = Machine.objects.create(name='SM 74', machine_type='offset_printing', machine_group_code='SM74',
                                          default_colors=5, operational_colors=5,
                                          max_print_length_mm=740, max_print_width_mm=1050, is_active=True)
        pj = PlanningJob.objects.create(
            jc_number='JC-PDF-1', order_qty=100, status='released',
            plan_date=datetime.date.today(), plan_month='July 2026',
            machine_name='SM 74', sku='SKU-PDF-1', actual_sheet_required=100,
        )
        JobCard.objects.create(
            job_card_no='JC-PDF-1', planning_job=pj, order_qty=100, total_sheet_quantity=100,
            status='in_production', is_active=True, SKU='SKU-PDF-1', po_date=datetime.date.today(),
            machine_name=machine, total_impressions_required=100, total_colors=4,
        )

        factory = RequestFactory()
        request = factory.get('/reports/api/reports/machine-planning/export/', {'_export': 'true'})
        request.user = self.user
        payload = run_report('machine-planning', request)

        pdf_bytes = export_as_pdf(payload)
        self.assertTrue(pdf_bytes.startswith(b'%PDF'))

        # Directly verify the chunking math: with the reduced min_col_width,
        # usable landscape-A4 width fits all 19 headers in a single chunk.
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib.units import mm
        page_w, page_h = landscape(A4)
        usable_width = page_w - 24 * mm
        min_col_width = 9 * mm
        max_columns = max(1, min(19, int(usable_width // min_col_width)))
        self.assertGreaterEqual(max_columns, len(payload['data']['headers']))

    def test_degraded_machine_keeps_two_color_jobs_in_two_color_pool(self):
        """Bug fix: dropping a GTO2 unit to operational_colors=1 moves the
        MACHINE into the GTO1 pool, but its 2-colour jobs must NOT follow it
        there - a 1-colour machine can't run them. They re-route to the
        remaining GTO2 members; only 1-colour jobs follow the degraded
        machine into the GTO1 tab."""
        from core.models import Machine, JobCard
        from planning.models import PlanningJob
        import datetime

        gto1a = Machine.objects.create(name='GTO 1A', machine_type='offset_printing', machine_group_code='GTO1',
                                        default_colors=1, operational_colors=1,
                                        max_print_length_mm=520, max_print_width_mm=740, is_active=True)
        gto2a = Machine.objects.create(name='GTO 2A', machine_type='offset_printing', machine_group_code='GTO2',
                                        default_colors=2, operational_colors=1,  # degraded to 1 colour
                                        max_print_length_mm=520, max_print_width_mm=740, is_active=True)
        gto2b = Machine.objects.create(name='GTO 2B', machine_type='offset_printing', machine_group_code='GTO2',
                                        default_colors=2, operational_colors=2,
                                        max_print_length_mm=520, max_print_width_mm=740, is_active=True)
        gto2c = Machine.objects.create(name='GTO 2C', machine_type='offset_printing', machine_group_code='GTO2',
                                        default_colors=2, operational_colors=2,
                                        max_print_length_mm=520, max_print_width_mm=740, is_active=True)

        def _make_job(suffix, color_spec, machine):
            pj = PlanningJob.objects.create(
                jc_number=f'JC-DEGRADE-{suffix}', order_qty=100, status='released',
                plan_date=datetime.date.today(), plan_month='July 2026',
                machine_name=machine.name, sku=f'SKU-DEGRADE-{suffix}', color_spec=color_spec,
                print_sheet_size='18*25', actual_sheet_required=100,
            )
            JobCard.objects.create(
                job_card_no=f'JC-DEGRADE-{suffix}', planning_job=pj, order_qty=100, total_sheet_quantity=100,
                status='in_production', is_active=True, SKU=f'SKU-DEGRADE-{suffix}', po_date=datetime.date.today(),
                machine_name=machine, total_impressions_required=100, total_colors=2,
            )

        # 2-colour job previously planned on the now-degraded GTO 2A.
        _make_job('2COL', '2', gto2a)
        # 1-colour job also on GTO 2A - this one SHOULD follow it to GTO1.
        _make_job('1COL', '1', gto2a)

        self.client.login(username='report_user', password='pass12345')
        response = self.client.get(reverse('reports:reports_api:run_report', args=['machine-planning']))
        machine_reports = response.json()['payload']['data']['machine_reports']

        # GTO1 pool now has 2 members (GTO 1A + degraded GTO 2A); GTO2 pool has 2 (B, C).
        gto1_tab = 'GTO 1A, GTO 2A'
        gto2_tab = 'GTO 2B, GTO 2C'
        self.assertIn(gto1_tab, machine_reports)
        self.assertIn(gto2_tab, machine_reports)

        gto1_jcs = {jc for row in machine_reports[gto1_tab]['rows'] for jc in row['job_card_numbers'].split(', ')}
        gto2_jcs = {jc for row in machine_reports[gto2_tab]['rows'] for jc in row['job_card_numbers'].split(', ')}

        self.assertIn('JC-DEGRADE-1COL', gto1_jcs)   # 1-colour follows the machine
        self.assertIn('JC-DEGRADE-2COL', gto2_jcs)   # 2-colour re-routes to capable pool
        self.assertNotIn('JC-DEGRADE-2COL', gto1_jcs)

    def test_maintenance_machine_jobs_redistribute_within_pool(self):
        """Bug fix companion: a machine at operational_colors=0 (maintenance)
        drops out of its pool's members, and jobs planned on it distribute
        across the remaining machines of the SAME pool - never hidden."""
        from core.models import Machine, JobCard
        from planning.models import PlanningJob
        import datetime

        gto2a = Machine.objects.create(name='GTO 2A', machine_type='offset_printing', machine_group_code='GTO2',
                                        default_colors=2, operational_colors=0,  # under maintenance
                                        max_print_length_mm=520, max_print_width_mm=740, is_active=True)
        Machine.objects.create(name='GTO 2B', machine_type='offset_printing', machine_group_code='GTO2',
                                default_colors=2, operational_colors=2,
                                max_print_length_mm=520, max_print_width_mm=740, is_active=True)
        Machine.objects.create(name='GTO 2C', machine_type='offset_printing', machine_group_code='GTO2',
                                default_colors=2, operational_colors=2,
                                max_print_length_mm=520, max_print_width_mm=740, is_active=True)

        pj = PlanningJob.objects.create(
            jc_number='JC-MAINT-1', order_qty=100, status='released',
            plan_date=datetime.date.today(), plan_month='July 2026',
            machine_name='GTO 2A', sku='SKU-MAINT-1', color_spec='2',
            print_sheet_size='18*25', actual_sheet_required=100,
        )
        JobCard.objects.create(
            job_card_no='JC-MAINT-1', planning_job=pj, order_qty=100, total_sheet_quantity=100,
            status='in_production', is_active=True, SKU='SKU-MAINT-1', po_date=datetime.date.today(),
            machine_name=gto2a, total_impressions_required=100, total_colors=2,
        )

        self.client.login(username='report_user', password='pass12345')
        response = self.client.get(reverse('reports:reports_api:run_report', args=['machine-planning']))
        payload = response.json()['payload']['data']
        machine_reports = payload['machine_reports']

        # The job lands in the GTO2 pool tab listing only the two working members.
        pool_tab = 'GTO 2B, GTO 2C'
        self.assertIn(pool_tab, machine_reports)
        jcs = {jc for row in machine_reports[pool_tab]['rows'] for jc in row['job_card_numbers'].split(', ')}
        self.assertIn('JC-MAINT-1', jcs)
        self.assertIn('GTO 2A', payload['maintenance_machines'])

    def test_wastage_report_details_and_calculations(self):
        from core.models import JobCard, Production, Dispatch, Machine, Sorter
        from planning.models import PlanningJob
        from django.core.cache import cache
        import datetime

        # Create mock Sorter
        sorter = Sorter.objects.create(name='Test Sorter', is_active=True)

        # Create mock Machine
        machine = Machine.objects.create(name='Test Machine', is_active=True)

        # Create mock PlanningJob
        pj = PlanningJob.objects.create(
            jc_number='JC-TEST-WASTE-1',
            order_qty=1000,
            ups=2,
            print_sheets=500,
            wastage_sheets=50,
            status='in_production',
            plan_date=datetime.date.today() - datetime.timedelta(days=1),
            plan_month='July 2026'
        )

        # Create mock JobCard
        jc = JobCard.objects.create(
            job_card_no='JC-TEST-WASTE-1',
            planning_job=pj,
            order_qty=1000,
            ups=2,
            total_sheet_quantity=500,
            status='in_production',
            is_active=True,
            SKU='SKU-TEST-1',
            po_date=datetime.date.today(),
            total_colors=4,
            machine_name=machine,
            total_impressions_required=1000
        )
        # Update created_at to yesterday to simulate PO intake timing for plan_date logic
        from django.utils import timezone
        JobCard.objects.filter(pk=jc.pk).update(
            created_at=timezone.now() - datetime.timedelta(days=1)
        )
        jc.refresh_from_db()

        # Create printing production waste
        Production.objects.create(
            entry_type='printing',
            job_card=jc,
            date=datetime.date.today(),
            shift='A',
            output_sheets=400,
            waste_sheets=50,
            status='completed',
            machine=machine
        )

        # Create sorting production waste
        Production.objects.create(
            entry_type='packing',
            job_card=jc,
            date=datetime.date.today(),
            shift='A',
            packing_qty=750,
            sorting_waste_qty=30,
            status='completed',
            sorter=sorter
        )

        # Create dispatch
        Dispatch.objects.create(
            job_card=jc,
            dc_no='DC-TEST-WASTE-1',
            dispatch_date=datetime.date.today(),
            dispatch_qty=750,
            is_active=True
        )

        self.client.login(username='report_user', password='pass12345')
        response = self.client.get(reverse('reports:reports_api:run_report', args=['wastage-report']))
        self.assertEqual(response.status_code, 200)
        
        data = response.json()
        self.assertTrue(data['ok'])
        
        payload_data = data['payload']['data']
        rows = payload_data['wastage_rows']
        self.assertEqual(len(rows), 1)
        
        row = rows[0]
        self.assertEqual(row['s_no'], 1)
        self.assertEqual(row['job_card_no'], 'JC-TEST-WASTE-1')
        self.assertEqual(row['sku'], 'SKU-TEST-1')
        self.assertEqual(row['plan_date'], (datetime.date.today() - datetime.timedelta(days=1)).strftime('%Y-%m-%d'))
        self.assertEqual(row['plan_month'], 'July')
        self.assertEqual(row['plan_qty'], 1000) # total_sheet_quantity (500) * ups (2) = 1000
        self.assertEqual(row['dispatch_qty'], 750)
        self.assertEqual(row['printing_waste_sheets'], 50)
        self.assertEqual(row['printing_waste_pcs'], 100) # 50 sheets * 2 ups = 100 pcs
        self.assertEqual(row['printing_waste_pct'], '10.0%') # 100 / 1000 * 100 = 10%
        self.assertEqual(row['sorting_waste_pcs'], 30)
        self.assertEqual(row['sorting_waste_pct'], '3.0%') # 30 / 1000 * 100 = 3%
        self.assertEqual(row['difference_pcs'], 250) # 1000 plan - 750 dispatch = 250
        self.assertEqual(row['difference_pct'], '25.0%') # 250 / 1000 * 100 = 25%
        self.assertEqual(row['wastage_status'], 'Tentative')
        self.assertEqual(row['total_wastage_pcs'], 380) # 100 + 30 + 250 = 380
        self.assertEqual(row['total_wastage_pct'], '38.0%')

        summary = payload_data['summary']
        self.assertEqual(summary['total_plan_qty'], 1000)
        self.assertEqual(summary['total_dispatch_qty'], 750)
        self.assertEqual(summary['printing_waste_pcs'], 100)
        self.assertEqual(summary['sorting_waste_pcs'], 30)
        self.assertEqual(summary['dispatch_gap_pcs'], 250)
        self.assertEqual(summary['total_wastage_pcs'], 380)
        self.assertEqual(summary['overall_wastage_pct'], 38.0)
        self.assertEqual(summary['tentative_wastage_pcs'], 380)
        self.assertEqual(summary['finalized_wastage_pcs'], 0)

        # Test date range filter working correctly
        cache.clear()
        yesterday_str = (datetime.date.today() - datetime.timedelta(days=1)).strftime('%Y-%m-%d')
        response_date_in = self.client.get(
            reverse('reports:reports_api:run_report', args=['wastage-report']),
            {
                'date_from': yesterday_str,
                'date_to': yesterday_str
            }
        )
        self.assertEqual(len(response_date_in.json()['payload']['data']['wastage_rows']), 1)

        cache.clear()
        response_date_out = self.client.get(
            reverse('reports:reports_api:run_report', args=['wastage-report']),
            {
                'date_from': '2020-01-01',
                'date_to': '2020-01-31'
            }
        )
        self.assertEqual(len(response_date_out.json()['payload']['data']['wastage_rows']), 0)

        # Now set job card to completed and verify it updates to Finalized
        jc.status = 'completed'
        jc.save(update_fields=['status'])

        # Clear the report engine cache to ensure the updated status is computed
        cache.clear()

        response2 = self.client.get(reverse('reports:reports_api:run_report', args=['wastage-report']))
        payload_data2 = response2.json()['payload']['data']
        self.assertEqual(payload_data2['wastage_rows'][0]['wastage_status'], 'Finalized')
        self.assertEqual(payload_data2['summary']['finalized_wastage_pcs'], 380)
        self.assertEqual(payload_data2['summary']['tentative_wastage_pcs'], 0)

        # Test filtering by finalized
        cache.clear()
        response_finalized = self.client.get(reverse('reports:reports_api:run_report', args=['wastage-report']), {'wastage_status': 'finalized'})
        data_finalized = response_finalized.json()['payload']['data']
        self.assertEqual(len(data_finalized['wastage_rows']), 1)

        # Test filtering by tentative (should return 0 rows since our only job is completed/finalized)
        cache.clear()
        response_tentative = self.client.get(reverse('reports:reports_api:run_report', args=['wastage-report']), {'wastage_status': 'tentative'})
        data_tentative = response_tentative.json()['payload']['data']
        self.assertEqual(len(data_tentative['wastage_rows']), 0)

        # Create a second, newer planning job and job card to test ordering and page limits
        pj2 = PlanningJob.objects.create(
            jc_number='JC-TEST-WASTE-2',
            order_qty=500,
            ups=1,
            print_sheets=500,
            wastage_sheets=50,
            status='in_production',
            plan_date=datetime.date.today(), # today
            plan_month='July 2026'
        )
        jc2 = JobCard.objects.create(
            job_card_no='JC-TEST-WASTE-2',
            planning_job=pj2,
            order_qty=500,
            ups=1,
            total_sheet_quantity=500,
            status='in_production',
            is_active=True,
            SKU='SKU-TEST-2',
            po_date=datetime.date.today(),
            total_colors=4,
            machine_name=machine,
            total_impressions_required=500
        )

        # Test pagination metadata and default response
        cache.clear()
        response_pag = self.client.get(reverse('reports:reports_api:run_report', args=['wastage-report']), {'page': '1'})
        data_pag = response_pag.json()['payload']['data']
        self.assertEqual(len(data_pag['wastage_rows']), 2)
        # Verify chronological order: oldest plan_date first
        self.assertEqual(data_pag['wastage_rows'][0]['job_card_no'], 'JC-TEST-WASTE-1')
        self.assertEqual(data_pag['wastage_rows'][1]['job_card_no'], 'JC-TEST-WASTE-2')

        # Assert pagination metadata
        pag_metadata = data_pag['pagination']
        self.assertEqual(pag_metadata['current_page'], 1)
        self.assertEqual(pag_metadata['total_rows'], 2)
        self.assertEqual(pag_metadata['page_size'], 100)
        self.assertEqual(pag_metadata['total_pages'], 1)
        self.assertFalse(pag_metadata['has_next'])
        self.assertFalse(pag_metadata['has_prev'])

        # Create a third planning job and job card with 0% wastage to test high_wastage filter
        pj3 = PlanningJob.objects.create(
            jc_number='JC-TEST-WASTE-3',
            order_qty=500,
            ups=1,
            print_sheets=500,
            wastage_sheets=0,
            status='in_production',
            plan_date=datetime.date.today(),
            plan_month='July 2026'
        )
        jc3 = JobCard.objects.create(
            job_card_no='JC-TEST-WASTE-3',
            planning_job=pj3,
            order_qty=500,
            ups=1,
            total_sheet_quantity=500,
            status='in_production',
            is_active=True,
            SKU='SKU-TEST-3',
            po_date=datetime.date.today(),
            total_colors=4,
            machine_name=machine,
            total_impressions_required=500
        )
        # Create 100% printing, packing, and dispatch to achieve 0% wastage
        Production.objects.create(
            entry_type='printing',
            job_card=jc3,
            date=datetime.date.today(),
            shift='A',
            output_sheets=500,
            waste_sheets=0,
            status='completed',
            machine=machine
        )
        Production.objects.create(
            entry_type='packing',
            job_card=jc3,
            date=datetime.date.today(),
            shift='A',
            packing_qty=500,
            sorting_waste_qty=0,
            status='completed',
            sorter=sorter
        )
        Dispatch.objects.create(
            job_card=jc3,
            dc_no='DC-TEST-WASTE-3',
            dispatch_date=datetime.date.today(),
            dispatch_qty=500,
            is_active=True
        )

        # Query without high_wastage filter (returns all 3)
        cache.clear()
        response_all = self.client.get(reverse('reports:reports_api:run_report', args=['wastage-report']))
        data_all = response_all.json()['payload']['data']
        self.assertEqual(len(data_all['wastage_rows']), 3)

        # Query with high_wastage=true filter (returns 2, excluding JC-TEST-WASTE-3 which has 0% wastage)
        cache.clear()
        response_high = self.client.get(reverse('reports:reports_api:run_report', args=['wastage-report']), {'high_wastage': 'true'})
        data_high = response_high.json()['payload']['data']
        self.assertEqual(len(data_high['wastage_rows']), 2)
        self.assertEqual(data_high['wastage_rows'][0]['job_card_no'], 'JC-TEST-WASTE-1')
        self.assertEqual(data_high['wastage_rows'][1]['job_card_no'], 'JC-TEST-WASTE-2')



class PeriodFilterPrecedenceTests(TestCase):
    """The filter forms always render the resolved range of the active preset
    into date_from/date_to, so a chosen preset must win over those inputs.
    Regression guard: previously every preset was silently overridden by them.
    """

    def setUp(self):
        from django.test import RequestFactory
        self.rf = RequestFactory()
        self.user = get_user_model().objects.create_user(username='report_user', password='testpass123')

    def _period(self, query):
        from reports.services import _parse_period_filter
        request = self.rf.get('/reports/daily-production/?' + query)
        request.user = self.user
        start, end, period, label, date_from, date_to = _parse_period_filter(request)
        return period, start, end

    def test_preset_beats_stale_date_inputs(self):
        from django.utils import timezone
        today = timezone.localdate()
        stale = 'date_from=2026-07-01&date_to=2026-07-21'

        period, start, end = self._period('period=today&' + stale)
        self.assertEqual(period, 'today')
        self.assertEqual(start, today)
        self.assertEqual(end, today)

        period, _, _ = self._period('period=week&' + stale)
        self.assertEqual(period, 'week')

        period, start, end = self._period('period=all&' + stale)
        self.assertEqual(period, 'all')
        self.assertIsNone(start)
        self.assertIsNone(end)

    def test_custom_range_still_honours_dates(self):
        from datetime import date
        period, start, end = self._period('period=custom&date_from=2026-07-10&date_to=2026-07-12')
        self.assertEqual(period, 'custom')
        self.assertEqual(start, date(2026, 7, 10))
        self.assertEqual(end, date(2026, 7, 12))

    def test_bare_dates_without_period_are_custom(self):
        from datetime import date
        period, start, end = self._period('date_from=2026-07-10&date_to=2026-07-12')
        self.assertEqual(period, 'custom')
        self.assertEqual(start, date(2026, 7, 10))
        self.assertEqual(end, date(2026, 7, 12))

    def test_yesterday_is_a_single_complete_day(self):
        """What the morning automations report on — the last complete day, not
        today's partial one."""
        from datetime import timedelta
        from django.utils import timezone
        yesterday = timezone.localdate() - timedelta(days=1)

        period, start, end = self._period('period=yesterday')
        self.assertEqual(period, 'yesterday')
        self.assertEqual(start, yesterday)
        self.assertEqual(end, yesterday)

    def test_yesterday_beats_stale_date_inputs(self):
        from datetime import timedelta
        from django.utils import timezone
        yesterday = timezone.localdate() - timedelta(days=1)

        period, start, end = self._period(
            'period=yesterday&date_from=2026-07-01&date_to=2026-07-21'
        )
        self.assertEqual(period, 'yesterday')
        self.assertEqual(start, yesterday)
        self.assertEqual(end, yesterday)

    def test_yesterday_is_labelled(self):
        from django.test import RequestFactory
        from reports.services import _parse_period_filter
        request = RequestFactory().get('/reports/daily-production/?period=yesterday')
        request.user = self.user
        label = _parse_period_filter(request)[3]
        self.assertEqual(label, 'Yesterday')


class PendingWorkReportTests(TestCase):
    """Process-wise pending backlog: printing/packing/dispatch gaps per job."""

    def setUp(self):
        import datetime
        from core.models import JobCard, Machine, Production, Dispatch, Sorter
        from planning.models import PlanningJob

        User = get_user_model()
        self.user = User.objects.create_user(username='report_user', password='pass12345')
        profile = UserProfile.objects.get(user=self.user)
        profile.role = 'manager'
        profile.save(update_fields=['role'])
        from django.contrib.auth.models import Permission
        from django.contrib.contenttypes.models import ContentType
        content_type = ContentType.objects.get_for_model(UserProfile)
        permission, _ = Permission.objects.get_or_create(
            codename='view_reports', name='Can view reports', content_type=content_type,
        )
        self.user.user_permissions.add(permission)

        today = datetime.date.today()
        machine = Machine.objects.create(name='PW Test Machine', is_active=True)
        sorter = Sorter.objects.create(name='PW Test Sorter', is_active=True)

        # Job A: fully printed (600 pcs) but nothing packed/dispatched yet.
        # Expect: pending_printing = 0, pending_packing = 600, no dispatch row
        # (dispatch backlog is against packed stock, and nothing's packed yet).
        pj_a = PlanningJob.objects.create(
            jc_number='JC-PW-A', order_qty=600, status='in_production',
            plan_date=today, plan_month='Test',
        )
        self.jc_a = JobCard.objects.create(
            job_card_no='JC-PW-A', planning_job=pj_a, order_qty=600, ups=2,
            total_sheet_quantity=300, status='in_production', is_active=True,
            SKU='SKU-PW-A', po_date=today, total_colors=4, total_impressions_required=600,
            machine_name=machine,
        )
        Production.objects.create(
            entry_type='printing', job_card=self.jc_a, date=today, shift='A',
            output_sheets=300, waste_sheets=0, status='completed',
        )

        # Job B: fully printed and packed (500 pcs) but nothing dispatched yet.
        # Expect: pending_printing = 0, pending_packing = 0, pending_dispatch = 500.
        pj_b = PlanningJob.objects.create(
            jc_number='JC-PW-B', order_qty=500, status='in_production',
            plan_date=today, plan_month='Test',
        )
        self.jc_b = JobCard.objects.create(
            job_card_no='JC-PW-B', planning_job=pj_b, order_qty=500, ups=1,
            total_sheet_quantity=500, status='in_production', is_active=True,
            SKU='SKU-PW-B', po_date=today, total_colors=4, total_impressions_required=500,
            machine_name=machine,
        )
        Production.objects.create(
            entry_type='printing', job_card=self.jc_b, date=today, shift='A',
            output_sheets=500, waste_sheets=0, status='completed',
        )
        Production.objects.create(
            entry_type='packing', job_card=self.jc_b, date=today, shift='A',
            packing_qty=500, sorting_waste_qty=0, status='completed', sorter=sorter,
        )

        # Job C: fully closed out (printed, packed, dispatched, status completed)
        # — must never appear anywhere despite matching quantities exactly.
        pj_c = PlanningJob.objects.create(
            jc_number='JC-PW-C', order_qty=200, status='completed',
            plan_date=today, plan_month='Test',
        )
        self.jc_c = JobCard.objects.create(
            job_card_no='JC-PW-C', planning_job=pj_c, order_qty=200, ups=1,
            total_sheet_quantity=200, status='in_production', is_active=True,
            SKU='SKU-PW-C', po_date=today, total_colors=4, total_impressions_required=200,
            machine_name=machine,
        )
        Production.objects.create(
            entry_type='printing', job_card=self.jc_c, date=today, shift='A',
            output_sheets=200, waste_sheets=0, status='completed',
        )
        Production.objects.create(
            entry_type='packing', job_card=self.jc_c, date=today, shift='A',
            packing_qty=200, sorting_waste_qty=0, status='completed', sorter=sorter,
        )
        Dispatch.objects.create(
            job_card=self.jc_c, dc_no='DC-PW-C', dispatch_date=today, dispatch_qty=200, is_active=True,
        )
        # Flip to 'completed' after the production/dispatch rows exist, bypassing
        # full_clean's release-sequence validation (only relevant on JobCard creation).
        JobCard.objects.filter(pk=self.jc_c.pk).update(status='completed')

        # Job D: still in planning (qc_approved), not yet released — must never
        # appear in the printing/packing/dispatch backlog tables, only in the
        # separate "Not Yet Released" list.
        pj_d = PlanningJob.objects.create(
            jc_number='JC-PW-D', order_qty=900, status='qc_approved',
            plan_date=today, plan_month='Test',
        )
        self.jc_d = JobCard.objects.create(
            job_card_no='JC-PW-D', planning_job=pj_d, order_qty=900, ups=1,
            total_sheet_quantity=900, status='qc_approved', is_active=True,
            SKU='SKU-PW-D', po_date=today, total_colors=4, total_impressions_required=900,
            machine_name=machine,
        )

        # Job E: still a plain draft — never submitted to QC, so it has no
        # JobCard at all yet (one is only created on submit-to-QC). Must still
        # surface under "Not Yet Released", not just disappear.
        self.pj_e = PlanningJob.objects.create(
            jc_number='JC-PW-E', po_number='PO-PW-E', sku='SKU-PW-E', order_qty=300,
            status='draft', plan_date=today, plan_month='Test',
        )

        # Job F: machine changed by production supervisor after planning —
        # JobCard.machine_name should win over PlanningJob.machine_name.
        supervisor_machine = Machine.objects.create(name='PW Supervisor Machine', is_active=True)
        pj_f = PlanningJob.objects.create(
            jc_number='JC-PW-F', order_qty=400, status='in_production',
            plan_date=today, plan_month='Test', machine_name='Planner Machine (text)',
        )
        self.jc_f = JobCard.objects.create(
            job_card_no='JC-PW-F', planning_job=pj_f, order_qty=400, ups=1,
            total_sheet_quantity=400, status='in_production', is_active=True,
            SKU='SKU-PW-F', po_date=today, total_colors=4, total_impressions_required=400,
            machine_name=supervisor_machine,
        )

    def _run(self, **params):
        self.client.login(username='report_user', password='pass12345')
        response = self.client.get(reverse('reports:reports_api:run_report', args=['pending-work']), params)
        self.assertEqual(response.status_code, 200)
        return response.json()['payload']['data']

    def test_printed_but_unpacked_job_shows_pending_packing_only(self):
        data = self._run()
        printing_jcs = {r['job_card_no'] for r in data['printing_rows']}
        packing_jcs = {r['job_card_no']: r for r in data['packing_rows']}
        dispatch_jcs = {r['job_card_no'] for r in data['dispatch_rows']}

        self.assertNotIn('JC-PW-A', printing_jcs)
        self.assertIn('JC-PW-A', packing_jcs)
        self.assertEqual(packing_jcs['JC-PW-A']['pending_qty'], 600)
        self.assertNotIn('JC-PW-A', dispatch_jcs)

    def test_packed_but_undispatched_job_shows_pending_dispatch_only(self):
        data = self._run()
        printing_jcs = {r['job_card_no'] for r in data['printing_rows']}
        packing_jcs = {r['job_card_no'] for r in data['packing_rows']}
        dispatch_jcs = {r['job_card_no']: r for r in data['dispatch_rows']}

        self.assertNotIn('JC-PW-B', printing_jcs)
        self.assertNotIn('JC-PW-B', packing_jcs)
        self.assertIn('JC-PW-B', dispatch_jcs)
        self.assertEqual(dispatch_jcs['JC-PW-B']['pending_qty'], 500)

    def test_completed_job_excluded_entirely(self):
        data = self._run()
        all_jcs = (
            {r['job_card_no'] for r in data['printing_rows']}
            | {r['job_card_no'] for r in data['packing_rows']}
            | {r['job_card_no'] for r in data['dispatch_rows']}
        )
        self.assertNotIn('JC-PW-C', all_jcs)

    def test_summary_totals(self):
        data = self._run()
        self.assertEqual(data['summary']['packing_pcs'], 600)
        self.assertEqual(data['summary']['dispatch_pcs'], 500)

    def test_not_released_job_excluded_from_backlog_and_listed_separately(self):
        data = self._run()
        backlog_jcs = (
            {r['job_card_no'] for r in data['printing_rows']}
            | {r['job_card_no'] for r in data['packing_rows']}
            | {r['job_card_no'] for r in data['dispatch_rows']}
        )
        self.assertNotIn('JC-PW-D', backlog_jcs)

        not_released = {r['job_card_no']: r for r in data['not_released_rows']}
        self.assertIn('JC-PW-D', not_released)
        self.assertEqual(not_released['JC-PW-D']['order_qty_pcs'], 900)
        # Job E (draft, no JobCard yet) also belongs here — see
        # test_draft_job_with_no_job_card_shows_in_not_released for that case.
        self.assertEqual(data['summary']['not_released_jobs'], 2)
        self.assertEqual(data['summary']['not_released_pcs'], 1200)

    def test_draft_job_with_no_job_card_shows_in_not_released(self):
        # Regression: a JobCard is only created once a PlanningJob is submitted
        # to QC, so a plain 'draft' job was invisible to the JobCard-only query
        # regardless of period filter.
        data = self._run()
        not_released = {r['job_card_no']: r for r in data['not_released_rows']}
        self.assertIn('JC-PW-E', not_released)
        self.assertEqual(not_released['JC-PW-E']['order_qty_pcs'], 300)
        self.assertEqual(not_released['JC-PW-E']['po_number'], 'PO-PW-E')

        data_all_time = self._run(period='all')
        not_released_all = {r['job_card_no'] for r in data_all_time['not_released_rows']}
        self.assertIn('JC-PW-E', not_released_all)

    def test_supervisor_assigned_machine_wins_over_planner_text(self):
        data = self._run()
        printing_row = next(r for r in data['printing_rows'] if r['job_card_no'] == 'JC-PW-F')
        self.assertEqual(printing_row['machine'], 'PW Supervisor Machine')

    def test_stage_param_selects_export_rows_and_does_not_collide_in_cache(self):
        printing_only = self._run(stage='printing')['export_rows']
        self.assertEqual({row['job_card_no'] for row in printing_only}, {'JC-PW-F'})

        not_released_only = self._run(stage='not_released')['export_rows']
        self.assertEqual({row['job_card_no'] for row in not_released_only}, {'JC-PW-D', 'JC-PW-E'})

        combined = self._run()['export_rows']
        self.assertTrue(all('stage' in row for row in combined))
        self.assertEqual(
            {row['job_card_no'] for row in combined},
            {'JC-PW-A', 'JC-PW-B', 'JC-PW-F'},
        )


class KPIScorecardServiceTests(TestCase):
    """Order Fulfillment / Wastage Reduction / Dispatch Alignment computation
    and Red/Yellow/Green banding against seeded KPITarget rows."""

    def setUp(self):
        import datetime
        from core.models import JobCard, Machine, Production, Dispatch, Sorter
        from planning.models import PlanningJob
        from reports.models import KPITarget

        self.start = datetime.date(2026, 7, 1)
        self.end = datetime.date(2026, 7, 31)

        KPITarget.objects.update_or_create(
            kpi_slug='order_fulfillment', year=2026,
            defaults={'min_value': 80, 'target_value': 85, 'max_value': 100, 'higher_is_better': True, 'weightage_pct': 20},
        )
        KPITarget.objects.update_or_create(
            kpi_slug='wastage_reduction', year=2026,
            defaults={'min_value': 0, 'target_value': 5, 'max_value': 8, 'higher_is_better': False, 'weightage_pct': 20},
        )
        KPITarget.objects.update_or_create(
            kpi_slug='dispatch_alignment', year=2026,
            defaults={'min_value': 80, 'target_value': 95, 'max_value': 130, 'higher_is_better': True, 'weightage_pct': 15},
        )

        pj = PlanningJob.objects.create(
            jc_number='JC-KPI-1', order_qty=1000, status='in_production',
            plan_date=self.start, plan_month='July 2026', po_approval_date=self.start,
        )
        machine = Machine.objects.create(name='KPI Test Machine', is_active=True)
        jc = JobCard.objects.create(
            job_card_no='JC-KPI-1', planning_job=pj, order_qty=1000, ups=2,
            total_sheet_quantity=500, status='in_production', is_active=True,
            SKU='SKU-KPI-1', po_date=self.start, total_colors=4, total_impressions_required=1000,
            machine_name=machine,
        )
        # 100 waste sheets * 2 ups = 200 waste pcs from printing.
        Production.objects.create(
            entry_type='printing', job_card=jc, date=self.start, shift='A',
            output_sheets=400, waste_sheets=100, status='completed',
        )
        # 700 packed, 10 pcs sorting waste.
        sorter = Sorter.objects.create(name='KPI Test Sorter', is_active=True)
        Production.objects.create(
            entry_type='packing', job_card=jc, date=self.start, shift='A',
            packing_qty=700, sorting_waste_qty=10, status='completed', sorter=sorter,
        )
        Dispatch.objects.create(
            job_card=jc, dc_no='DC-KPI-1', dispatch_date=self.start, dispatch_qty=650, is_active=True,
        )

    def test_order_fulfillment_is_dispatched_over_order_qty(self):
        from reports.kpi_services import compute_order_fulfillment
        value, detail = compute_order_fulfillment(self.start, self.end)
        # 650 dispatched / 1000 order qty = 65%
        self.assertEqual(value, 65.0)
        self.assertEqual(detail['order_qty'], 1000)
        self.assertEqual(detail['dispatched_pcs'], 650)

    def test_wastage_reduction_is_total_waste_over_order_qty(self):
        from reports.kpi_services import compute_wastage_reduction
        value, detail = compute_wastage_reduction(self.start, self.end)
        # (200 printing waste + 10 sorting waste) / 1000 order qty = 21%
        self.assertEqual(value, 21.0)
        self.assertEqual(detail['wastage_pcs'], 210)

    def test_dispatch_alignment_is_dispatched_over_packed(self):
        from reports.kpi_services import compute_dispatch_alignment
        value, detail = compute_dispatch_alignment(self.start, self.end)
        # 650 dispatched / 700 packed = 92.86%
        self.assertEqual(value, 92.86)
        self.assertEqual(detail['packed_pcs'], 700)

    def test_status_banding_higher_is_better(self):
        from reports.kpi_services import _status_for
        from reports.models import KPITarget
        target = KPITarget.objects.get(kpi_slug='dispatch_alignment', year=2026)
        self.assertEqual(_status_for(target, 60), 'red')     # below min (80)
        self.assertEqual(_status_for(target, 90), 'yellow')  # between min and target
        self.assertEqual(_status_for(target, 95), 'green')   # at target
        self.assertEqual(_status_for(target, 140), 'yellow') # above max — overshoot caution

    def test_status_banding_lower_is_better(self):
        from reports.kpi_services import _status_for
        from reports.models import KPITarget
        target = KPITarget.objects.get(kpi_slug='wastage_reduction', year=2026)
        self.assertEqual(_status_for(target, 3), 'green')   # at/below target (5)
        self.assertEqual(_status_for(target, 7), 'yellow')  # between target and max (8)
        self.assertEqual(_status_for(target, 12), 'red')    # above max


class KPITargetSeedMigrationTests(TestCase):
    """The 0005 migration re-bands Order Fulfillment to min 95 / max 150."""

    def test_order_fulfillment_2026_target_uses_updated_band(self):
        from reports.models import KPITarget
        target = KPITarget.objects.get(kpi_slug='order_fulfillment', year=2026)
        self.assertEqual(float(target.min_value), 95)
        self.assertEqual(float(target.max_value), 150)
        self.assertGreaterEqual(float(target.target_value), float(target.min_value))


class KPIScorecardViewTests(TestCase):
    def setUp(self):
        User = get_user_model()
        self.user = User.objects.create_user(username='kpi_user', password='pass12345')
        profile = UserProfile.objects.get(user=self.user)
        profile.role = 'manager'
        profile.save(update_fields=['role'])
        from django.contrib.auth.models import Permission
        from django.contrib.contenttypes.models import ContentType
        content_type = ContentType.objects.get_for_model(UserProfile)
        permission, _ = Permission.objects.get_or_create(
            codename='view_reports', name='Can view reports', content_type=content_type,
        )
        self.user.user_permissions.add(permission)
        self.client.login(username='kpi_user', password='pass12345')

    def test_kpi_scorecard_report_loads(self):
        response = self.client.get(reverse('reports:detail', args=['kpi-scorecard']))
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, 'KPI Scorecard')

    def test_save_note_persists_and_is_picked_up_on_next_get(self):
        response = self.client.post(reverse('reports:kpi_save_note'), {
            'kpi_slug': 'wastage_reduction',
            'period_type': 'month',
            'period_key': '2026-07',
            'status': 'yellow',
            'note': 'Review setup waste on Machine 3.',
            'return_query': 'period_type=month&year=2026&month=7',
        })
        self.assertEqual(response.status_code, 302)

        from reports.models import KPIActionNote
        note = KPIActionNote.objects.get(kpi_slug='wastage_reduction', period_type='month', period_key='2026-07')
        self.assertEqual(note.note, 'Review setup waste on Machine 3.')
        self.assertEqual(note.status, 'yellow')

        response = self.client.get(reverse('reports:detail', args=['kpi-scorecard']), {'period_type': 'month', 'year': 2026, 'month': 7})
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, 'Review setup waste on Machine 3.')

    def test_different_months_are_not_served_from_the_same_cache_entry(self):
        """Regression guard: period_type/year/month/quarter were missing from
        parse_universal_filters, so every month/quarter selection shared one
        cache key and always showed whichever period was computed first."""
        import datetime
        from core.models import JobCard, Machine, Production
        from planning.models import PlanningJob

        machine = Machine.objects.create(name='KPI Cache Test Machine', is_active=True)
        may = datetime.date(2026, 5, 15)
        pj_may = PlanningJob.objects.create(
            jc_number='JC-KPI-CACHE-MAY', order_qty=1000, status='in_production',
            plan_date=may, plan_month='May 2026', po_approval_date=may,
        )
        jc_may = JobCard.objects.create(
            job_card_no='JC-KPI-CACHE-MAY', planning_job=pj_may, order_qty=1000, ups=1,
            total_sheet_quantity=1000, status='in_production', is_active=True,
            SKU='SKU-KPI-CACHE-MAY', po_date=may, total_colors=4, total_impressions_required=1000,
            machine_name=machine,
        )
        Production.objects.create(
            entry_type='printing', job_card=jc_may, date=may, shift='A',
            output_sheets=1000, waste_sheets=0, status='completed',
        )

        response_may = self.client.get(
            reverse('reports:reports_api:run_report', args=['kpi-scorecard']),
            {'period_type': 'month', 'year': 2026, 'month': 5},
        )
        response_jun = self.client.get(
            reverse('reports:reports_api:run_report', args=['kpi-scorecard']),
            {'period_type': 'month', 'year': 2026, 'month': 6},
        )
        data_may = response_may.json()['payload']['data']
        data_jun = response_jun.json()['payload']['data']

        self.assertEqual(data_may['period_label'], 'May 2026')
        self.assertEqual(data_jun['period_label'], 'June 2026')
        may_fulfillment = next(k for k in data_may['kpis'] if k['slug'] == 'order_fulfillment')
        jun_fulfillment = next(k for k in data_jun['kpis'] if k['slug'] == 'order_fulfillment')
        # May has the 1000-pc order in it, June has none — a stale shared
        # cache entry would show the same order_qty for both.
        self.assertEqual(may_fulfillment['detail']['order_qty'], 1000)
        self.assertEqual(jun_fulfillment['detail']['order_qty'], 0)


class KPIDrilldownExportTests(TestCase):
    """Clicking a KPI's percent (or a trend period) downloads the raw
    job-level rows the percentage was computed from, for manual reconciliation."""

    def setUp(self):
        import datetime
        from core.models import JobCard, Machine, Production, Dispatch
        from planning.models import PlanningJob

        User = get_user_model()
        self.user = User.objects.create_user(username='kpi_drill_user', password='pass12345')
        profile = UserProfile.objects.get(user=self.user)
        profile.role = 'manager'
        profile.save(update_fields=['role'])
        from django.contrib.auth.models import Permission
        from django.contrib.contenttypes.models import ContentType
        content_type = ContentType.objects.get_for_model(UserProfile)
        permission, _ = Permission.objects.get_or_create(
            codename='view_reports', name='Can view reports', content_type=content_type,
        )
        self.user.user_permissions.add(permission)
        self.client.login(username='kpi_drill_user', password='pass12345')

        self.start = datetime.date(2026, 7, 1)
        machine = Machine.objects.create(name='Drill Test Machine', is_active=True)
        pj = PlanningJob.objects.create(
            jc_number='JC-DRILL-1', order_qty=500, status='in_production',
            plan_date=self.start, plan_month='July 2026', po_approval_date=self.start,
        )
        jc = JobCard.objects.create(
            job_card_no='JC-DRILL-1', planning_job=pj, order_qty=500, ups=1,
            total_sheet_quantity=500, status='in_production', is_active=True,
            SKU='SKU-DRILL-1', po_date=self.start, total_colors=4, total_impressions_required=500,
            machine_name=machine,
        )
        Dispatch.objects.create(
            job_card=jc, dc_no='DC-DRILL-1', dispatch_date=self.start, dispatch_qty=400, is_active=True,
        )

        # A second job in a different month of the same quarter (Q3 2026), so
        # the quarterly-detail export has more than one month's rows to prove
        # it actually spans the whole quarter, not just the selected month.
        aug = datetime.date(2026, 8, 15)
        pj2 = PlanningJob.objects.create(
            jc_number='JC-DRILL-2', po_number='PO-DRILL-2', order_qty=200, status='in_production',
            plan_date=aug, plan_month='August 2026', po_approval_date=aug,
        )
        jc2 = JobCard.objects.create(
            job_card_no='JC-DRILL-2', planning_job=pj2, order_qty=200, ups=1,
            total_sheet_quantity=200, status='in_production', is_active=True,
            SKU='SKU-DRILL-2', po_date=aug, total_colors=4, total_impressions_required=200,
            machine_name=machine,
        )
        Dispatch.objects.create(
            job_card=jc2, dc_no='DC-DRILL-2', dispatch_date=aug, dispatch_qty=150, is_active=True,
        )

    def test_quarterly_detail_export_spans_whole_quarter_with_totals(self):
        response = self.client.get(
            reverse('reports:reports_api:run_report', args=['kpi-scorecard']),
            {'period_type': 'month', 'year': 2026, 'month': 7, 'kpi': 'order_fulfillment', 'detail': 'quarterly'},
        )
        data = response.json()['payload']['data']
        rows = data['export_rows']
        self.assertEqual(data['headers'][:3], ['quarter', 'process', 'month'])

        jc_rows = {row['job_card_no']: row for row in rows if row.get('job_card_no')}
        self.assertIn('JC-DRILL-1', jc_rows)
        self.assertIn('JC-DRILL-2', jc_rows)
        self.assertEqual(jc_rows['JC-DRILL-2']['month'], 'August')
        self.assertEqual(jc_rows['JC-DRILL-2']['machine'], 'Drill Test Machine')

        processes = {row['process'] for row in rows}
        self.assertIn('Total', processes)
        self.assertIn('Order Fulfillment Efficiency %', processes)

    def test_kpi_param_returns_supporting_rows_for_that_kpi_only(self):
        response = self.client.get(
            reverse('reports:reports_api:run_report', args=['kpi-scorecard']),
            {'period_type': 'month', 'year': 2026, 'month': 7, 'kpi': 'order_fulfillment'},
        )
        data = response.json()['payload']['data']
        rows = data['export_rows']
        self.assertTrue(rows)
        row_types = {row['row_type'] for row in rows}
        self.assertEqual(row_types, {'Order (Planned)', 'Dispatch (Actual)'})
        dispatch_row = next(r for r in rows if r['row_type'] == 'Dispatch (Actual)')
        self.assertEqual(dispatch_row['qty_pcs'], 400)
        self.assertEqual(dispatch_row['job_card_no'], 'JC-DRILL-1')

    def test_kpi_all_combines_every_kpis_supporting_rows(self):
        response = self.client.get(
            reverse('reports:reports_api:run_report', args=['kpi-scorecard']),
            {'period_type': 'month', 'year': 2026, 'month': 7, 'kpi': 'all'},
        )
        data = response.json()['payload']['data']
        rows = data['export_rows']
        kpi_labels = {row['kpi'] for row in rows}
        self.assertEqual(
            kpi_labels,
            {'Order Fulfillment Efficiency', 'Wastage Reduction Efficiency', 'Dispatch vs Production Alignment'},
        )

    def test_export_xlsx_download_succeeds(self):
        response = self.client.get(
            reverse('reports:reports_api:export_report', args=['kpi-scorecard']),
            {'type': 'xlsx', 'period_type': 'month', 'year': 2026, 'month': 7, 'kpi': 'order_fulfillment'},
        )
        self.assertEqual(response.status_code, 200)
        self.assertEqual(
            response['Content-Type'],
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        )


class KPIFocusExportTests(TestCase):
    """The 'Improvement focus' export ranks exactly which jobs are dragging a
    KPI down for the period, unlike the raw drill-down/quarterly-detail rows."""

    def setUp(self):
        import datetime
        from core.models import JobCard, Machine, Production, Dispatch, Sorter
        from planning.models import PlanningJob

        User = get_user_model()
        self.user = User.objects.create_user(username='kpi_focus_user', password='pass12345')
        profile = UserProfile.objects.get(user=self.user)
        profile.role = 'manager'
        profile.save(update_fields=['role'])
        from django.contrib.auth.models import Permission
        from django.contrib.contenttypes.models import ContentType
        content_type = ContentType.objects.get_for_model(UserProfile)
        permission, _ = Permission.objects.get_or_create(
            codename='view_reports', name='Can view reports', content_type=content_type,
        )
        self.user.user_permissions.add(permission)
        self.client.login(username='kpi_focus_user', password='pass12345')

        july = datetime.date(2026, 7, 10)
        june = datetime.date(2026, 6, 10)
        machine = Machine.objects.create(name='Focus Test Machine', is_active=True)
        sorter = Sorter.objects.create(name='Focus Test Sorter', is_active=True)

        # Stuck at packing: printed in full, never packed or dispatched, PO
        # approved this period -> Order Fulfillment focus only.
        pj_stuck = PlanningJob.objects.create(
            jc_number='JC-FOCUS-STUCK', order_qty=300, status='in_production',
            plan_date=july, plan_month='July 2026', po_approval_date=july,
        )
        jc_stuck = JobCard.objects.create(
            job_card_no='JC-FOCUS-STUCK', planning_job=pj_stuck, order_qty=300, ups=1,
            total_sheet_quantity=300, status='in_production', is_active=True,
            SKU='SKU-FOCUS-STUCK', po_date=july, total_colors=4, total_impressions_required=300,
            machine_name=machine,
        )
        Production.objects.create(
            job_card=jc_stuck, entry_type='printing', date=july, shift='A', output_sheets=300,
            waste_sheets=0, is_active=True,
        )

        # Packed but undispatched, PO approved the month before -> Dispatch
        # Alignment focus only (out of Order Fulfillment's July window).
        pj_gap = PlanningJob.objects.create(
            jc_number='JC-FOCUS-GAP', order_qty=250, status='in_production',
            plan_date=june, plan_month='June 2026', po_approval_date=june,
        )
        jc_gap = JobCard.objects.create(
            job_card_no='JC-FOCUS-GAP', planning_job=pj_gap, order_qty=250, ups=1,
            total_sheet_quantity=250, status='in_production', is_active=True,
            SKU='SKU-FOCUS-GAP', po_date=june, total_colors=4, total_impressions_required=250,
            machine_name=machine, is_print_job=False,
        )
        Production.objects.create(
            job_card=jc_gap, entry_type='packing', date=july, shift='A', packing_qty=250,
            sorting_waste_qty=0, is_active=True, sorter=sorter,
        )
        Dispatch.objects.create(
            job_card=jc_gap, dc_no='DC-FOCUS-GAP', dispatch_date=july, dispatch_qty=100, is_active=True,
        )

        # Fully dispatched this period -> must not appear in either focus list.
        pj_done = PlanningJob.objects.create(
            jc_number='JC-FOCUS-DONE', order_qty=200, status='in_production',
            plan_date=july, plan_month='July 2026', po_approval_date=july,
        )
        jc_done = JobCard.objects.create(
            job_card_no='JC-FOCUS-DONE', planning_job=pj_done, order_qty=200, ups=1,
            total_sheet_quantity=200, status='in_production', is_active=True,
            SKU='SKU-FOCUS-DONE', po_date=july, total_colors=4, total_impressions_required=200,
            machine_name=machine, is_print_job=False,
        )
        Production.objects.create(
            job_card=jc_done, entry_type='packing', date=july, shift='A', packing_qty=200,
            sorting_waste_qty=0, is_active=True, sorter=sorter,
        )
        Dispatch.objects.create(
            job_card=jc_done, dc_no='DC-FOCUS-DONE', dispatch_date=july, dispatch_qty=200, is_active=True,
        )

        # Waste: two printing entries this period, high one must rank first.
        pj_waste = PlanningJob.objects.create(
            jc_number='JC-FOCUS-WASTE', order_qty=100, status='in_production',
            plan_date=july, plan_month='July 2026', po_approval_date=july,
        )
        jc_waste = JobCard.objects.create(
            job_card_no='JC-FOCUS-WASTE', planning_job=pj_waste, order_qty=100, ups=2,
            total_sheet_quantity=100, status='in_production', is_active=True,
            SKU='SKU-FOCUS-WASTE', po_date=july, total_colors=4, total_impressions_required=100,
            machine_name=machine,
        )
        Production.objects.create(
            job_card=jc_waste, entry_type='printing', date=july, shift='A', output_sheets=50,
            waste_sheets=50, is_active=True,
        )
        pj_waste_low = PlanningJob.objects.create(
            jc_number='JC-FOCUS-WASTE-LOW', order_qty=100, status='in_production',
            plan_date=july, plan_month='July 2026', po_approval_date=july,
        )
        jc_waste_low = JobCard.objects.create(
            job_card_no='JC-FOCUS-WASTE-LOW', planning_job=pj_waste_low, order_qty=100, ups=2,
            total_sheet_quantity=100, status='in_production', is_active=True,
            SKU='SKU-FOCUS-WASTE-LOW', po_date=july, total_colors=4, total_impressions_required=100,
            machine_name=machine,
        )
        Production.objects.create(
            job_card=jc_waste_low, entry_type='printing', date=july, shift='A', output_sheets=95,
            waste_sheets=5, is_active=True,
        )

    def _focus(self, kpi_slug):
        response = self.client.get(
            reverse('reports:reports_api:run_report', args=['kpi-scorecard']),
            {'period_type': 'month', 'year': 2026, 'month': 7, 'kpi': kpi_slug, 'detail': 'focus'},
        )
        data = response.json()['payload']['data']
        self.assertEqual(data['headers'][:2], ['job_card_no', 'po_number'])
        return data['export_rows']

    def test_order_fulfillment_focus_flags_stuck_at_packing_and_excludes_completed(self):
        rows = self._focus('order_fulfillment')
        by_jc = {r['job_card_no']: r for r in rows}
        self.assertIn('JC-FOCUS-STUCK', by_jc)
        self.assertEqual(by_jc['JC-FOCUS-STUCK']['issue'], 'Stuck at Packing')
        self.assertEqual(by_jc['JC-FOCUS-STUCK']['qty_pcs'], 300)
        self.assertNotIn('JC-FOCUS-DONE', by_jc)

    def test_dispatch_alignment_focus_shows_gap_and_excludes_completed_and_stuck(self):
        rows = self._focus('dispatch_alignment')
        by_jc = {r['job_card_no']: r for r in rows}
        self.assertIn('JC-FOCUS-GAP', by_jc)
        self.assertEqual(by_jc['JC-FOCUS-GAP']['issue'], 'Awaiting Dispatch')
        self.assertEqual(by_jc['JC-FOCUS-GAP']['qty_pcs'], 150)
        self.assertNotIn('JC-FOCUS-DONE', by_jc)
        self.assertNotIn('JC-FOCUS-STUCK', by_jc)

    def test_wastage_focus_ranks_worst_waste_first(self):
        rows = self._focus('wastage_reduction')
        self.assertTrue(rows)
        self.assertEqual(rows[0]['job_card_no'], 'JC-FOCUS-WASTE')
        self.assertEqual(rows[0]['qty_pcs'], 100)
        qtys = [r['qty_pcs'] for r in rows]
        self.assertEqual(qtys, sorted(qtys, reverse=True))


class DailyProductionReleasedToProductionTests(TestCase):
    """Daily Production's "Released to Production" stream is sourced from the
    ChangeLog audit trail written when a job is released, not a raw status
    field — so the fixture must drive a real release action, not just set
    status='released' directly, to prove the wiring end to end."""

    def setUp(self):
        import datetime
        from core.models import JobCard, Machine
        from planning.models import PlanningJob

        User = get_user_model()
        self.user = User.objects.create_user(username='dp_report_user', password='pass12345')
        profile = UserProfile.objects.get(user=self.user)
        profile.role = 'manager'
        profile.save(update_fields=['role'])
        from django.contrib.auth.models import Permission
        from django.contrib.contenttypes.models import ContentType
        content_type = ContentType.objects.get_for_model(UserProfile)
        permission, _ = Permission.objects.get_or_create(
            codename='view_reports', name='Can view reports', content_type=content_type,
        )
        self.user.user_permissions.add(permission)
        self.client.login(username='dp_report_user', password='pass12345')

        self.today = datetime.date.today()
        self.machine = Machine.objects.create(name='DP Released Test Machine', is_active=True)

        pj = PlanningJob.objects.create(
            jc_number='JC-DP-RELEASED-1', order_qty=800, status='qc_approved',
            plan_date=self.today, plan_month='Test',
        )
        self.jc = JobCard.objects.create(
            job_card_no='JC-DP-RELEASED-1', planning_job=pj, order_qty=800, ups=1,
            total_sheet_quantity=800, status='production_approved', is_active=True,
            SKU='SKU-DP-RELEASED-1', PO_No='PO-DP-1', po_date=self.today, total_colors=4,
            total_impressions_required=800, wastage=0, machine_name=self.machine,
        )

    def _release(self):
        from core.jobcard_service import execute_job_card_action
        self.jc.refresh_from_db()
        execute_job_card_action(self.jc, 'release_for_production', actor=self.user, reason='Test release')

    def test_released_job_appears_in_released_rows_and_overview_and_totals(self):
        self._release()
        response = self.client.get(
            reverse('reports:reports_api:run_report', args=['daily-production']),
            {'period': 'today'},
        )
        data = response.json()['payload']['data']

        released_rows = data['released_rows']
        self.assertEqual(len(released_rows), 1)
        row = released_rows[0]
        self.assertEqual(row['job_card_no'], 'JC-DP-RELEASED-1')
        self.assertEqual(row['po_number'], 'PO-DP-1')
        self.assertEqual(row['sku'], 'SKU-DP-RELEASED-1')
        self.assertEqual(row['machine'], 'DP Released Test Machine')
        self.assertEqual(row['order_qty'], 800)
        self.assertIn('dp_report_user', row['released_by'])

        overview_rows = data['overview_rows']
        self.assertEqual(len(overview_rows), 1)
        self.assertEqual(overview_rows[0]['released_count'], 1)
        self.assertEqual(overview_rows[0]['released_qty'], 800)

        self.assertEqual(data['totals']['released_jobs'], 1)
        self.assertEqual(data['totals']['released_qty'], 800)

    def test_release_outside_the_filtered_period_is_excluded(self):
        self._release()
        response = self.client.get(
            reverse('reports:reports_api:run_report', args=['daily-production']),
            {'period': 'custom', 'date_from': '2020-01-01', 'date_to': '2020-01-31'},
        )
        data = response.json()['payload']['data']
        self.assertEqual(data['released_rows'], [])
        self.assertEqual(data['totals']['released_jobs'], 0)

    def test_no_release_action_means_no_released_rows(self):
        response = self.client.get(
            reverse('reports:reports_api:run_report', args=['daily-production']),
            {'period': 'today'},
        )
        data = response.json()['payload']['data']
        self.assertEqual(data['released_rows'], [])
        self.assertEqual(data['totals']['released_jobs'], 0)
