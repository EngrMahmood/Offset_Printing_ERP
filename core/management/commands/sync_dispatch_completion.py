from django.core.management.base import BaseCommand
from django.db import transaction
from django.db.models import Sum, F, Q
from django.db.models.functions import Coalesce
from core.models import JobCard
from core.jobcard_service import transition_job_card_status

class Command(BaseCommand):
    help = "Syncs workflow status of fully dispatched Job Cards (and linked Planning Jobs) to 'completed'."

    def handle(self, *args, **options):
        self.stdout.write("Scanning for active Job Cards with dispatch ratio >= 95% in 'in_production' status...")
        
        jc_qs = JobCard.objects.filter(is_active=True).annotate(
            dispatched_total=Coalesce(
                Sum('dispatch__dispatch_qty', filter=Q(dispatch__is_active=True)),
                0,
            )
        )
        
        complete_jcs = list(jc_qs.filter(dispatched_total__gte=F('order_qty') * 0.95, status='in_production'))
        
        if not complete_jcs:
            self.stdout.write(self.style.SUCCESS("No unsynced fully-dispatched Job Cards found."))
            return

        self.stdout.write(f"Found {len(complete_jcs)} Job Cards to update. Syncing...")
        
        success_count = 0
        failed_count = 0

        for jc in complete_jcs:
            try:
                transition_job_card_status(jc, 'completed', reason='One-time sync: Dispatch completion reached >= 95%')
                success_count += 1
            except Exception as e:
                self.stdout.write(self.style.ERROR(f"Failed to sync {jc.job_card_no}: {e}"))
                failed_count += 1

        self.stdout.write(self.style.SUCCESS(f"Sync complete. Successfully completed: {success_count}, Failed: {failed_count}"))
