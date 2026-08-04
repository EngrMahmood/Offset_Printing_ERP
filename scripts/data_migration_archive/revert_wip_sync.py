import os
import sys
import django

# Set up path and Django settings module
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'Offset_ERP.settings')
django.setup()

from core.models import JobCard, JobCardWipStatus, ProductionWipStatus, ChangeLog

def main():
    print("==================================================")
    print("Starting production WIP sync revert process...")
    print("==================================================")
    
    # Find all ChangeLogs created today by the sync run
    logs = ChangeLog.objects.filter(
        entity_type='job_card',
        change_reason__contains="WIP Status updated to",
        created_at__gte='2026-07-15'
    )
    
    print(f"Found {logs.count()} sync log entries to revert.")
    reverted_count = 0
    
    for log in logs:
        if 'wip_status' in log.field_changes:
            old_val = log.field_changes['wip_status']['from']
            
            # If the status was previously set to something else (not 'Not Set')
            if old_val and old_val != 'Not Set':
                try:
                    jc = JobCard.objects.get(pk=log.record_id)
                    status_obj = ProductionWipStatus.objects.filter(name=old_val, is_active=True).first()
                    if status_obj:
                        # Restore the old status and lock it as manual override
                        JobCardWipStatus.objects.update_or_create(
                            job_card=jc,
                            defaults={
                                'status': status_obj,
                                'is_manual': True
                            }
                        )
                        reverted_count += 1
                except JobCard.DoesNotExist:
                    pass
            else:
                # If it was 'Not Set', delete the status or set it to none
                try:
                    jc = JobCard.objects.get(pk=log.record_id)
                    JobCardWipStatus.objects.filter(job_card=jc).delete()
                    reverted_count += 1
                except JobCard.DoesNotExist:
                    pass

    print(f"-> Reverted {reverted_count} records to their original pre-sync state.")
    print("Revert completed successfully!")
    print("==================================================")

if __name__ == '__main__':
    main()
