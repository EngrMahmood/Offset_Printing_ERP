import os
import sys
import django

# Set up path and Django settings module
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'Offset_ERP.settings')
django.setup()

from core.models import JobCard
from production.wip_service import evaluate_and_update_job_wip_status

def main():
    print("==================================================")
    print("Starting production data synchronization script...")
    print("==================================================")
    
    # 1. Sync WIP Statuses
    print("\n1. Syncing WIP statuses based on logs...")
    jobs = JobCard.objects.filter(is_active=True, status__in=['released', 'in_production'])
    wip_count = 0
    for jc in jobs:
        updated = evaluate_and_update_job_wip_status(jc, force=True)
        if updated:
            wip_count += 1
    print(f"-> Successfully updated WIP status for {wip_count} Job Cards.")
    
    # 2. Sync PO Dates
    print("\n2. Syncing Job Card PO Dates using PO Approval Dates...")
    po_jobs = JobCard.objects.filter(is_active=True, planning_job__isnull=False)
    po_count = 0
    for jc in po_jobs:
        if jc.planning_job:
            old_date = jc.po_date
            new_date = jc.planning_job.po_approval_date or jc.planning_job.po_received_date
            if old_date != new_date:
                jc.po_date = new_date
                jc.save(update_fields=['po_date'])
                po_count += 1
    print(f"-> Successfully resynced PO Date for {po_count} Job Cards.")
    
    print("\n==================================================")
    print("Sync completed successfully!")
    print("==================================================")

if __name__ == '__main__':
    main()
