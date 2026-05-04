"""QC service layer stubs for approval and review workflows."""

# TODO: Move approval workflow and QC sync logic from planning.views here.


def validate_planning_job_for_qc(planning_job):
    """Validate a PlanningJob before QC approval."""
    raise NotImplementedError


def transition_job_card_status(job_card, action, actor=None, reason=''):
    """Perform job card approval transitions in a centralized service."""
    raise NotImplementedError
