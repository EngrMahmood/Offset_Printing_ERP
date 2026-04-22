import os
import sys
import django
from io import BytesIO
from datetime import date, timedelta

ROOT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'Offset_ERP.settings')
django.setup()

from django.contrib.auth.models import User
from django.contrib.auth import get_user_model
from django.core.files.uploadedfile import SimpleUploadedFile
from django.test import Client
from django.utils import timezone
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

from core.models import UserProfile
from planning.models import PlanningJob, PlanningPrintRun, PlanningDispatchRun, SkuRecipe, PoDocument


def create_admin_user(username='planning_admin', password='Admin123!'):
    user, created = User.objects.get_or_create(username=username, defaults={'email': 'planning_admin@example.com', 'is_staff': True, 'is_superuser': True})
    if created or not user.check_password(password):
        user.set_password(password)
        user.is_staff = True
        user.is_superuser = True
        user.save()
    profile, _ = UserProfile.objects.update_or_create(user=user, defaults={'role': 'admin'})
    return user


def clear_planning_data():
    print('Clearing existing planning records...')
    for doc in PoDocument.objects.all():
        try:
            if doc.po_file:
                doc.po_file.delete(save=False)
        except Exception:
            pass
    PlanningDispatchRun.objects.all().delete()
    PlanningPrintRun.objects.all().delete()
    PoDocument.objects.all().delete()
    PlanningJob.objects.all().delete()
    SkuRecipe.objects.all().delete()
    print('Planning records cleared.')


def build_po_pdf(lines):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    c.setFont('Helvetica', 12)
    y = 760
    for line in lines:
        c.drawString(50, y, line)
        y -= 16
    c.showPage()
    c.save()
    buffer.seek(0)
    return buffer.getvalue()


def setup_repeat_sku_history(user):
    print('Setting up repeat SKU historical recipe and job...')
    recipe, created = SkuRecipe.objects.get_or_create(
        sku='REP-SKU-001',
        defaults={
            'job_name': 'Repeat SKU 001',
            'material': 'Paper',
            'color_spec': '4+0',
            'application': 'UV',
            'size_w_mm': 210,
            'size_h_mm': 297,
            'ups': 1,
            'print_sheet_size': 'A4',
            'purchase_sheet_size': 'A4',
            'purchase_sheet_ups': 1,
            'purchase_material': 'Local',
            'machine_name': 'Machine A',
            'department': 'Printing',
            'default_unit_cost': 10.00,
            'master_data_status': 'approved',
            'approved_by': user,
            'approved_at': timezone.now(),
        },
    )
    recipe.save()
    job, created = PlanningJob.objects.get_or_create(
        jc_number='RJ-1001',
        defaults={
            'plan_date': date.today(),
            'po_number': 'PO-2026-0001',
            'sku': 'REP-SKU-001',
            'job_name': 'Repeat SKU Job',
            'repeat_flag': 'Repeat',
            'material': 'Paper',
            'color_spec': '4+0',
            'application': 'UV',
            'size_w_mm': 210,
            'size_h_mm': 297,
            'order_qty': 100,
            'print_sheet_size': 'A4',
            'purchase_sheet_size': 'A4',
            'purchase_sheet_ups': 1,
            'purchase_material': 'Local',
            'machine_name': 'Machine A',
            'department': 'Printing',
            'status': 'draft',
            'issued_to_production': False,
            'is_active': True,
        },
    )
    job.save()
    return recipe, job


def run_full_planning_flow(admin_username='planning_admin', admin_password='Admin123!'):
    print('Running full planning app workflow test...')
    client = Client()
    logged_in = client.login(username=admin_username, password=admin_password)
    assert logged_in, 'Login failed for admin user.'

    # Upload PO PDF
    po_lines = [
        'PURCHASE ORDER PO-2026-0001',
        'Dated Apr 22, 2026',
        'Delivery Location SITE-1',
        'Supplier Details Name SupplierA',
        'Buyer Details Name BuyerA',
        '1 REP-SKU-001 Apr 30, 2026 100 PIECE Rs 10 1000 180 1180',
        '2 NEW-SKU-001 May 05, 2026 200 PIECE Rs 20 4000 720 4720',
    ]
    pdf_bytes = build_po_pdf(po_lines)
    uploaded = SimpleUploadedFile('test_po.pdf', pdf_bytes, content_type='application/pdf')

    res = client.post('/planning/po/upload/', {'po_pdf': uploaded})
    assert res.status_code in (302, 303), f'Upload PO redirect failed: {res.status_code}'

    po_doc = PoDocument.objects.filter(extracted_payload__po_number='PO-2026-0001').order_by('-id').first()
    assert po_doc is not None, 'PO document was not created.'
    assert len(po_doc.extracted_payload.get('items', [])) == 2, 'PO parser did not detect 2 line items.'

    # Verify PO intake queue and review page
    res = client.get('/planning/po/inbox/')
    assert res.status_code == 200, '/planning/po/inbox/ failed.'

    res = client.get(f'/planning/po/{po_doc.id}/review/')
    assert res.status_code == 200, f'/planning/po/{po_doc.id}/review/ failed.'

    items = po_doc.extracted_payload.get('items', [])
    data = {'action': 'create_jobs'}
    for item in items:
        prefix = f"item_{item['line_no']}_"
        data[f'{prefix}print_sheet_size'] = item.get('print_sheet_size', '')
        data[f'{prefix}purchase_sheet_size'] = item.get('purchase_sheet_size', '')
        data[f'{prefix}machine_name'] = item.get('machine_name', '')
        data[f'{prefix}ups'] = item.get('ups', '')

    res = client.post(f'/planning/po/{po_doc.id}/review/', data)
    assert res.status_code in (302, 303), 'PO review create_jobs did not redirect.'

    # Ensure missing SKU recipe flows are available
    res = client.get(f'/planning/po/{po_doc.id}/new-skus/')
    assert res.status_code == 200, '/planning/po_new_skus/ failed.'

    # Save missing SKU recipe as draft
    new_sku = 'NEW-SKU-001'
    recipe_data = {
        f'sku_{new_sku}_job_name': 'New SKU Job',
        f'sku_{new_sku}_material': 'Paper',
        f'sku_{new_sku}_color_spec': '4+0',
        f'sku_{new_sku}_application': 'UV',
        f'sku_{new_sku}_machine_name': 'Machine A',
        f'sku_{new_sku}_department': 'Printing',
        f'sku_{new_sku}_print_sheet_size': 'A4',
        f'sku_{new_sku}_purchase_sheet_size': 'A4',
        f'sku_{new_sku}_ups': '1',
        f'sku_{new_sku}_purchase_material': 'Local',
    }
    res = client.post(f'/planning/po/{po_doc.id}/new-skus/', recipe_data)
    assert res.status_code in (302, 303), 'PO new skus save did not redirect.'

    recipe = SkuRecipe.objects.filter(sku__iexact=new_sku).order_by('-id').first()
    assert recipe is not None, 'Missing SKU recipe was not created.'
    assert recipe.master_data_status == 'draft', 'New SKU recipe was not saved as draft.'

    # Submit recipe for review and approve via master data flow
    edit_data = {
        'sku': recipe.sku,
        'job_name': recipe.job_name,
        'material': recipe.material,
        'color_spec': recipe.color_spec,
        'application': recipe.application,
        'size_w_mm': str(recipe.size_w_mm or ''),
        'size_h_mm': str(recipe.size_h_mm or ''),
        'ups': str(recipe.ups or ''),
        'print_sheet_size': recipe.print_sheet_size,
        'purchase_sheet_size': recipe.purchase_sheet_size,
        'purchase_sheet_ups': str(recipe.purchase_sheet_ups or ''),
        'purchase_material': recipe.purchase_material,
        'machine_name': recipe.machine_name,
        'department': recipe.department,
        'default_unit_cost': str(recipe.default_unit_cost or ''),
        'daily_demand': str(recipe.daily_demand or ''),
        'awc_no': recipe.awc_no,
        'plate_set_no': recipe.plate_set_no,
        'die_cutting': recipe.die_cutting,
        'notes': recipe.notes,
    }
    edit_data['action'] = 'submit_review'
    res = client.post(f'/planning/sku-recipes/{recipe.id}/edit/', edit_data)
    assert res.status_code in (302, 303), 'Recipe submit_review did not redirect.'
    recipe.refresh_from_db()
    assert recipe.master_data_status == 'pending_review', 'Recipe was not moved to Pending Review.'

    edit_data['action'] = 'review'
    res = client.post(f'/planning/sku-recipes/{recipe.id}/edit/', edit_data)
    assert res.status_code in (302, 303), 'Recipe review action did not redirect.'
    recipe.refresh_from_db()
    assert recipe.master_data_status == 'reviewed', 'Recipe was not moved to Reviewed.'

    # Approve the recipe via pending_skus approval flow and sync new job into planning
    res = client.post('/planning/pending-skus/', {'action': 'approve', 'sku': recipe.sku})
    assert res.status_code in (302, 303), 'Pending SKUs approve did not redirect.'
    recipe.refresh_from_db()
    assert recipe.master_data_status == 'approved', 'Recipe approval failed.'

    new_job = PlanningJob.objects.filter(sku__iexact=new_sku, is_active=True).first()
    assert new_job is not None, 'Approved new SKU did not create a planning job.'

    # Check planning job and approval queue pages
    res = client.get('/planning/jobs/')
    assert res.status_code == 200, '/planning/jobs/ failed.'
    res = client.get('/planning/approval-queue/')
    assert res.status_code == 200, '/planning/approval-queue/ failed.'

    # Use scan to open job by JC
    res = client.post('/planning/scan/', {'scan_code': new_job.jc_number})
    assert res.status_code in (302, 303), 'Scan post did not redirect.'

    # Open scan direct JC URL
    res = client.get(f'/planning/scan/open/{new_job.jc_number}/')
    assert res.status_code in (302, 303), 'Scan open JC URL did not redirect.'

    res = client.get('/planning/report/')
    assert res.status_code == 200, '/planning/report/ failed.'

    print('Full planning app workflow test passed successfully.')


if __name__ == '__main__':
    admin = create_admin_user()
    clear_planning_data()
    setup_repeat_sku_history(admin)
    run_full_planning_flow()
