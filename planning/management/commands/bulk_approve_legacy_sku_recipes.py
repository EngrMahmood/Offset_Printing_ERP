"""Bulk-approve legacy SKU master records (admin migration path)."""

from __future__ import annotations

from django.contrib.auth import get_user_model
from django.core.management.base import BaseCommand, CommandError
from django.db import transaction
from django.utils import timezone

from planning.models import PlanningJob, SkuRecipe
from planning.services import _sync_new_jobs_for_approved_sku
from workflow.services import _missing_required_master_fields

LEGACY_PLACEHOLDER = '-'


class Command(BaseCommand):
    help = (
        'Bulk approve legacy_produced SKU master records. '
        'Fills blank AWC # and Plate Set No. with "-" (legacy sheet convention) before approval.'
    )

    def add_arguments(self, parser):
        parser.add_argument(
            '--username',
            help='Approver username (defaults to first superuser).',
        )
        parser.add_argument(
            '--dry-run',
            action='store_true',
            help='Report what would change without saving.',
        )
        parser.add_argument(
            '--include-non-legacy',
            action='store_true',
            help='Also approve non-legacy SKUs that pass required-field checks.',
        )
        parser.add_argument(
            '--sync-jobs',
            action='store_true',
            help='Refresh planning jobs for approved SKUs that have jobs.',
        )

    def handle(self, *args, **options):
        User = get_user_model()
        approver = None
        if options.get('username'):
            approver = User.objects.filter(username=options['username']).first()
            if not approver:
                raise CommandError(f'User not found: {options["username"]}')
        else:
            approver = User.objects.filter(is_superuser=True).order_by('id').first()
        if not approver:
            raise CommandError('No superuser found to record as approver.')

        now = timezone.now()
        qs = SkuRecipe.objects.filter(is_active=True).exclude(master_data_status='approved')
        if not options['include_non_legacy']:
            qs = qs.filter(legacy_produced=True)

        approved = 0
        placeholder_filled = 0
        skipped_missing = []
        sync_targets = []

        def _run():
            nonlocal approved, placeholder_filled
            for recipe in qs.iterator(chunk_size=500):
                changed = []
                if not (recipe.awc_no or '').strip():
                    recipe.awc_no = LEGACY_PLACEHOLDER
                    changed.append('awc_no')
                if not (recipe.plate_set_no or '').strip():
                    recipe.plate_set_no = LEGACY_PLACEHOLDER
                    changed.append('plate_set_no')
                if changed:
                    placeholder_filled += 1

                missing = _missing_required_master_fields(recipe)
                if missing:
                    skipped_missing.append((recipe.sku, missing))
                    continue

                recipe.legacy_produced = True
                recipe.master_data_status = 'approved'
                recipe.reviewed_by = approver
                recipe.reviewed_at = now
                recipe.approved_by = approver
                recipe.approved_at = now
                update_fields = list(dict.fromkeys(changed + [
                    'legacy_produced',
                    'master_data_status',
                    'reviewed_by',
                    'reviewed_at',
                    'approved_by',
                    'approved_at',
                    'updated_at',
                ]))
                recipe.save(update_fields=update_fields)
                approved += 1
                if options['sync_jobs'] and PlanningJob.objects.filter(sku__iexact=recipe.sku).exists():
                    sync_targets.append(recipe.sku)

        if options['dry_run']:
            with transaction.atomic():
                _run()
                transaction.set_rollback(True)
        else:
            _run()

        synced = 0
        if options['sync_jobs'] and not options['dry_run']:
            for sku in sync_targets:
                result = _sync_new_jobs_for_approved_sku(sku, actor=approver)
                synced += result.get('updated', 0)

        mode = 'DRY-RUN' if options['dry_run'] else 'APPLIED'
        self.stdout.write(self.style.SUCCESS(
            f'[{mode}] approver={approver.username} approved={approved} '
            f'placeholder_filled={placeholder_filled} skipped_missing={len(skipped_missing)} '
            f'jobs_synced={synced}'
        ))
        if skipped_missing:
            self.stdout.write('Skipped (still missing required fields):')
            for sku, missing in skipped_missing[:25]:
                self.stdout.write(f'  {sku}: {", ".join(missing)}')
            if len(skipped_missing) > 25:
                self.stdout.write(f'  ... and {len(skipped_missing) - 25} more')
