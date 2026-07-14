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
        self.assertEqual(row['plan_month'], 'July 2026')
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

