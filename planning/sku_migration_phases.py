"""Three-phase SKU master migration safety tiers.

Phase 1 — Never entered planning or production (no PlanningJob for SKU).
Phase 2 — In planning only (has jobs, none released to production).
Phase 3 — Released to production (at least one active job in production).
"""

from __future__ import annotations

from collections import defaultdict

from planning.models import PlanningJob

PRODUCTION_PLANNING_STATUSES = frozenset({
    'released',
    'in_production',
    'completed',
})

PRODUCTION_JOB_CARD_STATUSES = frozenset({
    'released',
    'in_production',
    'completed',
    'closed',
})

SKU_PREFIX_PRODUCT_TYPE_RULES = (
    ('INSERTCARD', 'Insert Card'),
    ('LABELCARE', 'Care Label'),
    ('SIZELABEL', 'Size Label'),
    ('LAWLABEL', 'Law Label'),
    ('STICKER', 'Sticker'),
    ('WRAPPAPER', 'Wrap Paper'),
    ('IMPORTERLABEL', 'Importer Label'),
    ('IMPORTERSLEEVE', 'Importer Sleeve'),
    ('FLUFFINGINSTRUCTION', 'Fluffing Instruction'),
    ('BELLYBAND', 'Belly Band'),
    ('COLORBOX', 'Color Box'),
    ('GAZETTED', 'Poly Bag'),
    ('HEADER', 'Header Card'),
    ('INNERPAPER', 'Inner Paper'),
    ('CATALOG', 'Catalog'),
    ('WARNINGLABEL', 'Warning Label'),
    ('IDENTIFICATIONCARD', 'Identification Card'),
)


def planning_job_is_in_production(job):
    """True when the job has been released to production or finished."""
    if not job:
        return False
    status = (getattr(job, 'workflow_status', None) or getattr(job, 'status', None) or '').strip()
    if status in PRODUCTION_PLANNING_STATUSES:
        return True
    try:
        job_card = job.job_card
    except PlanningJob.job_card.RelatedObjectDoesNotExist:
        job_card = None
    except Exception:
        job_card = None
    if job_card and (job_card.workflow_status or '').strip() in PRODUCTION_JOB_CARD_STATUSES:
        return True
    return False


def infer_product_type_from_sku(sku):
    """Best-effort product type from SKU prefix (phase-1 ERP-only SKUs)."""
    upper = str(sku or '').strip().upper()
    if not upper:
        return ''
    for prefix, product_type in SKU_PREFIX_PRODUCT_TYPE_RULES:
        if upper.startswith(prefix):
            return product_type
    return ''


def build_sku_phase_map():
    """
    Return {sku_casefold: phase_int} for every SKU seen on jobs or recipes.

    Phase 1: no PlanningJob ever.
    Phase 2: has jobs, none in production.
    Phase 3: at least one active job in production.
    """
    jobs_by_sku = defaultdict(list)
    for job in PlanningJob.objects.exclude(sku='').select_related('job_card').iterator(chunk_size=500):
        key = job.sku.strip().casefold()
        if key:
            jobs_by_sku[key].append(job)

    phase_map = {}
    for key, jobs in jobs_by_sku.items():
        active_jobs = [job for job in jobs if job.is_active]
        check_jobs = active_jobs or jobs
        if any(planning_job_is_in_production(job) for job in check_jobs):
            phase_map[key] = 3
        else:
            phase_map[key] = 2
    return phase_map


def get_sku_migration_phase(sku, phase_map=None):
    """Return migration phase (1, 2, or 3) for a SKU string."""
    key = str(sku or '').strip().casefold()
    if not key:
        return 1
    if phase_map is None:
        phase_map = build_sku_phase_map()
    return phase_map.get(key, 1)


def sku_eligible_for_migration_phase(sku, migration_phase, phase_map=None):
    """True when SKU is in the requested migration phase."""
    if migration_phase is None:
        return True
    return get_sku_migration_phase(sku, phase_map=phase_map) == migration_phase
