import logging
from decimal import Decimal, InvalidOperation

from django.core.files.base import ContentFile
from django.db import transaction
from django.db.models import Q

from migration.models import (
    JobStatus,
    MigrationImportLog,
    PlanningImportStaging,
    RowImportStatus,
)
from planning.models import PoDocument, PlanningJob
from planning.services import _sync_repeat_jobs_from_po

logger = logging.getLogger(__name__)


def _normalize_purchase_material_origin(raw_value):
    value = (raw_value or '').strip().lower()
    if value in {'local'}:
        return 'local'
    if value in {'import', 'imported'}:
        return 'import'
    return ''


def _get_raw_value(raw_data, keys):
    for key in keys:
        if key in raw_data and raw_data[key] not in (None, ''):
            return raw_data[key]
    return None


def _to_decimal(value):
    if value is None or value == '':
        return None
    try:
        return Decimal(str(value).replace(',', '').strip())
    except (InvalidOperation, TypeError, ValueError):
        return None


def _build_po_payload_from_staging(row):
    raw = row.raw_data or {}
    delivery_date_value = _get_raw_value(raw, ['delivery_date', 'date', 'plan_date', 'po_approval_date', 'po_received_date', 'po_date'])
    parsed_delivery_date = row.delivery_date

    def _serialize_date(value):
        if value is None or value == '':
            return ''
        if hasattr(value, 'isoformat'):
            return value.isoformat()
        return str(value).strip()

    return {
        'jc_number': _get_raw_value(raw, ['jc', 'jc_number', 'jc_no', 'job_card_no', 'jobcard_no', 'job_card', 'jobcard']),
        'plan_month': _get_raw_value(raw, ['month', 'plan_month']),
        'po_number': row.po_number,
        'pr_number': _get_raw_value(raw, ['pr', 'indent', 'indent_pr', 'indent/pr']),
        'po_date': _serialize_date(parsed_delivery_date or delivery_date_value or ''),
        'delivery_location': row.customer or _get_raw_value(raw, ['destination', 'delivery_location']),
        'department': _get_raw_value(raw, ['department']),
        'purchase_material': _normalize_purchase_material_origin(_get_raw_value(raw, ['purchase_material', 'purchase_material_origin'])),
        'machine_name': _get_raw_value(raw, ['machine_name', 'machine']),
        'items': [
            {
                'line_no': row.row_number,
                'sku': row.sku,
                'job_name': _get_raw_value(raw, ['job_name']),
                'repeat_flag': _get_raw_value(raw, ['repeat', 'repeat_flag']),
                'material': _get_raw_value(raw, ['material']),
                'color_spec': _get_raw_value(raw, ['color', 'color_spec']),
                'application': _get_raw_value(raw, ['application']),
                'quantity': row.quantity,
                'size_w_mm': _get_raw_value(raw, ['size_w_mm']),
                'size_h_mm': _get_raw_value(raw, ['size_h_mm']),
                'ups': _get_raw_value(raw, ['ups', 'no_of_ups']),
                'print_sheet_size': _get_raw_value(raw, ['print_sheet_size']),
                'actual_sheet_required': _get_raw_value(raw, ['actual_sheet_required', 'actual_sheet_require', 'sheet']),
                'wastage': _get_raw_value(raw, ['wastage', 'wastage_sheets']),
                'purchase_sheet_size': _get_raw_value(raw, ['purchase_sheet_size']),
                'purchase_sheet_ups': _get_raw_value(raw, ['purchase_sheet_ups']),
                'purchase_sheet_required': _get_raw_value(raw, ['purchase_sheet_required', 'purchase_sheet_require']),
                'pkt': _get_raw_value(raw, ['pkt', 'pkt_value']),
                'remarks': _get_raw_value(raw, ['remarks']),
                'requirement': _get_raw_value(raw, ['requirement', 'requirement_special_instructions']),
                'purchase_material_origin': _normalize_purchase_material_origin(_get_raw_value(raw, ['purchase_material', 'purchase_material_origin'])),
                'stock_qty': _get_raw_value(raw, ['stock', 'stock_qty']),
                'awc_no': _get_raw_value(raw, ['awc_no']),
                'plate_set_no': _get_raw_value(raw, ['plate_set_no', 'p_set_no']),
                'die_cutting': _get_raw_value(raw, ['die_cutting']),
                'machine_name': _get_raw_value(raw, ['machine_name', 'machine']),
                'cost': _get_raw_value(raw, ['cost', 'unit_cost']),
                'status': _get_raw_value(raw, ['status']),
                'balance_qty': _get_raw_value(raw, ['balance', 'balance_qty']),
                'delivery_date': _serialize_date(parsed_delivery_date or delivery_date_value),
            }
        ],
    }


def _find_existing_planning_job_for_row(row):
    raw = row.raw_data or {}
    jc_number = _get_raw_value(raw, ['jc', 'jc_number', 'jc_no', 'job_card_no', 'jobcard_no', 'job_card', 'jobcard'])
    try:
        quantity = int(str(row.quantity).replace(',', '').strip())
    except (TypeError, ValueError, AttributeError):
        quantity = None
    print_sheet_size = _get_raw_value(raw, ['print_sheet_size'])

    if not (row.po_number and row.sku):
        return None

    if jc_number:
        filters = {
            'po_number__iexact': row.po_number,
            'sku__iexact': row.sku,
            'jc_number__iexact': jc_number,
        }

        if quantity is not None:
            exact_match = PlanningJob.objects.filter(**filters, order_qty=quantity)
            if print_sheet_size:
                exact_match = exact_match.filter(print_sheet_size__iexact=print_sheet_size)
            exact_job = exact_match.order_by('-updated_at', '-id').first()
            if exact_job:
                return exact_job

        if print_sheet_size:
            broader_match = PlanningJob.objects.filter(**filters, print_sheet_size__iexact=print_sheet_size)
        else:
            broader_match = PlanningJob.objects.filter(**filters)
        broader_job = broader_match.order_by('-updated_at', '-id').first()
        if broader_job:
            return broader_job

    po_sku_match = PlanningJob.objects.filter(
        po_number__iexact=row.po_number,
        sku__iexact=row.sku,
    )
    if print_sheet_size:
        po_sku_match = po_sku_match.filter(print_sheet_size__iexact=print_sheet_size)
    return po_sku_match.order_by('-updated_at', '-id').first()


def _import_planning_row(row, actor):
    payload = _build_po_payload_from_staging(row)
    file_name = f'migration_job_{row.import_job_id}_row_{row.row_number}.txt'

    po_doc = PoDocument.objects.create(
        po_file=ContentFile(b'Generated by migration app', name=file_name),
        extracted_payload=payload,
        extraction_status='processed',
        uploaded_by=actor,
    )

    result = _sync_repeat_jobs_from_po(po_doc, actor=actor, bypass_recipe_check=True)
    imported = (result.get('created', 0) + result.get('updated', 0)) > 0
    linked_job = None

    if not imported:
        linked_job = _find_existing_planning_job_for_row(row)
        if linked_job:
            # If a matching existing PlanningJob exists, force update it instead of creating a duplicate.
            forced_payload = po_doc.extracted_payload or {}
            forced_payload['jc_number'] = linked_job.jc_number
            po_doc.extracted_payload = forced_payload
            po_doc.save(update_fields=['extracted_payload'])

            result = _sync_repeat_jobs_from_po(po_doc, actor=actor, bypass_recipe_check=True)
            imported = (result.get('created', 0) + result.get('updated', 0)) > 0

            if not imported:
                # Preserve the link if the existing job was still a valid match.
                imported = True
                row.imported_reference = linked_job.jc_number
                row.error_message = ''

    if imported and linked_job is None:
        linked_job = PlanningJob.objects.filter(
            po_number__iexact=row.po_number,
            sku__iexact=row.sku,
        ).order_by('-updated_at', '-id').first()

    if imported:
        row.import_status = RowImportStatus.IMPORTED
        row.error_message = ''
        row.imported_reference = linked_job.jc_number if linked_job else row.imported_reference
    else:
        row.import_status = RowImportStatus.ERROR
        row.error_message = 'No planning record was created/updated by planning service.'

    row.save(update_fields=['import_status', 'error_message', 'imported_reference', 'updated_at'])
    return imported


def import_planning_job(import_job, actor):
    """Import only VALID rows via existing planning business logic.

    Each row runs inside its own transaction so one failure does not block others.
    """
    valid_rows = PlanningImportStaging.objects.filter(
        import_job=import_job,
        import_status=RowImportStatus.VALID,
    ).order_by('row_number', 'id')

    success_count = 0
    error_count = 0

    for row in valid_rows:
        try:
            with transaction.atomic():
                imported = _import_planning_row(row, actor)
                if imported:
                    success_count += 1
                else:
                    error_count += 1
        except Exception as exc:
            logger.exception('Import failed for job=%s row=%s', import_job.id, row.row_number)
            row.import_status = RowImportStatus.ERROR
            row.error_message = str(exc)
            row.save(update_fields=['import_status', 'error_message', 'updated_at'])
            error_count += 1

    total_rows = import_job.planning_rows.count()
    imported_rows = import_job.planning_rows.filter(import_status=RowImportStatus.IMPORTED).count()
    pending_valid_rows = import_job.planning_rows.filter(import_status=RowImportStatus.VALID).count()
    current_error_rows = import_job.planning_rows.filter(import_status=RowImportStatus.ERROR).count()

    if imported_rows and pending_valid_rows:
        import_job.status = JobStatus.PARTIAL
    elif imported_rows and not pending_valid_rows:
        import_job.status = JobStatus.COMPLETED
    elif current_error_rows and not imported_rows:
        import_job.status = JobStatus.FAILED
    else:
        import_job.status = JobStatus.VALIDATED

    import_job.imported_rows = imported_rows
    import_job.error_rows = current_error_rows
    import_job.save(update_fields=['status', 'imported_rows', 'error_rows', 'updated_at'])

    MigrationImportLog.objects.create(
        import_job=import_job,
        imported_by=actor,
        rows_count=total_rows,
        success_count=success_count,
        error_count=error_count,
        message='Planning import executed via planning service layer.',
    )

    return {
        'total': total_rows,
        'success': success_count,
        'errors': error_count,
        'imported': imported_rows,
    }


def get_imported_planning_jobs(import_job):
    """Return PlanningJob records that were imported by a MigrationImportJob."""
    imported_rows = import_job.planning_rows.filter(import_status=RowImportStatus.IMPORTED)
    if not imported_rows.exists():
        return []

    jobs = []
    for row in imported_rows:
        if row.po_number and row.sku:
            if row.imported_reference:
                matches = PlanningJob.objects.filter(
                    jc_number__iexact=row.imported_reference,
                    po_number__iexact=row.po_number,
                    sku__iexact=row.sku,
                ).filter(
                    Q(created_at__gte=import_job.created_at) | Q(updated_at__gte=import_job.created_at)
                )
            else:
                matches = PlanningJob.objects.filter(
                    po_number__iexact=row.po_number,
                    sku__iexact=row.sku,
                ).filter(
                    Q(created_at__gte=import_job.created_at) | Q(updated_at__gte=import_job.created_at)
                )
            jobs.extend(matches)

    unique_jobs = {job.id: job for job in jobs}.values()
    return list(unique_jobs)


def rollback_imported_planning_jobs(import_job, dry_run=False):
    jobs = get_imported_planning_jobs(import_job)
    if not jobs:
        return 0

    if dry_run:
        return len(jobs)

    deleted_count = 0
    for job in jobs:
        job.delete()
        deleted_count += 1
    return deleted_count


def cleanup_imported_planning_job(import_job):
    imported_rows = import_job.planning_rows.filter(import_status=RowImportStatus.IMPORTED)
    if not imported_rows.exists():
        return {'deleted': 0, 'reset': 0}

    referenced_jcs = set(
        imported_rows.exclude(imported_reference='').values_list('imported_reference', flat=True)
    )
    deleted_count = 0
    if referenced_jcs:
        deleted_count, _ = PlanningJob.objects.filter(jc_number__in=referenced_jcs).delete()

    reset_count = imported_rows.update(
        import_status=RowImportStatus.VALID,
        imported_reference='',
        error_message='',
    )

    valid_count = import_job.planning_rows.filter(import_status=RowImportStatus.VALID).count()
    error_count = import_job.planning_rows.filter(import_status=RowImportStatus.ERROR).count()
    import_job.imported_rows = 0
    import_job.valid_rows = valid_count
    import_job.error_rows = error_count
    if valid_count:
        import_job.status = JobStatus.VALIDATED
    elif error_count:
        import_job.status = JobStatus.FAILED
    else:
        import_job.status = JobStatus.STAGED
    import_job.save(update_fields=['imported_rows', 'valid_rows', 'error_rows', 'status', 'updated_at'])

    return {'deleted': deleted_count, 'reset': reset_count}
