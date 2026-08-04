import os
import sys
import django

# Set up path and Django settings module
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'Offset_ERP.settings')
django.setup()

from core.models import JobCard, JobCardWipStatus
from production.wip_service import evaluate_and_update_job_wip_status

def main():
    print("==================================================")
    print("Starting production data synchronization script...")
    print("==================================================")
    
    # 1. Mark all existing WIP records as Manual Override (since they were set manually by supervisors previously)
    print("\n1. Marking all existing supervisor WIP statuses as manual overrides...")
    existing_wip = JobCardWipStatus.objects.all()
    marked_count = 0
    for wip in existing_wip:
        if not wip.is_manual:
            wip.is_manual = True
            wip.save(update_fields=['is_manual'])
            marked_count += 1
    print(f"-> Marked {marked_count} existing WIP statuses as manual (retaining physical check values).")
    
    # 2. Auto-initialize only Job Cards that have NO WIP status record at all
    print("\n2. Initializing auto WIP status for Job Cards with no status set...")
    jobs = JobCard.objects.filter(is_active=True, status__in=['released', 'in_production'], production_wip_status__isnull=True)
    wip_count = 0
    for jc in jobs:
        # force=False to ensure we do not touch overridden status (though these have none)
        updated = evaluate_and_update_job_wip_status(jc, force=False)
        if updated:
            wip_count += 1
    print(f"-> Automatically set WIP status for {wip_count} new Job Cards.")
    
    # 3. Sync PO Dates
    print("\n3. Syncing Job Card PO Dates using PO Approval Dates...")
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
