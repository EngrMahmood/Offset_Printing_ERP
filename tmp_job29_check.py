import os
os.environ['DJANGO_SETTINGS_MODULE'] = 'Offset_ERP.settings'
import django
django.setup()
from planning.models import PlanningJob
from planning.views import _normalize_status

job = PlanningJob.objects.filter(id=29).first()
print('job found:', bool(job))
if not job:
    raise SystemExit(1)
print('status raw:', repr(job.status))
print('status normalized:', _normalize_status(job.status))
print('workflow status:', job.status)
print('job_card_version', job.job_card_version)
print('repeat_flag', repr(job.repeat_flag))
print('plate_set_no', repr(job.plate_set_no))
print('wastage_sheets', job.wastage_sheets)
print('machine_name', repr(job.machine_name))
print('remarks', repr(job.remarks))
print('print_sheet_size', repr(job.print_sheet_size))
print('purchase_sheet_size', repr(job.purchase_sheet_size))
print('ups', job.ups)
print('order_qty', job.order_qty)
print('total_colors', job.total_colors)
print('actual_sheet_required', job.actual_sheet_required)
print('calculated_sheets_required', job.calculated_sheets_required)
print('qc_validation_errors', job.qc_validation_errors())
print('can_send_to_qc raw validation?')
try:
    job.clean()
    print('clean OK')
except Exception as e:
    print('clean exception:', type(e).__name__, e)
