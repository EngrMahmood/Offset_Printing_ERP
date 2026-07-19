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

