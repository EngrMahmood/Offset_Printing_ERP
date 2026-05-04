import os

views_path = r"d:\Offset_Printing_ERP\core\views.py"
services_path = r"d:\Offset_Printing_ERP\core\services.py"

with open(views_path, "r", encoding="utf-8") as f:
    lines = f.readlines()

start_idx = None
end_idx = None

for i, line in enumerate(lines):
    if "def format_audit_value(value):" in line:
        start_idx = i
        break

for i in range(start_idx, len(lines)):
    if "@login_required" in lines[i] and "def home(request):" in lines[i+1]:
        end_idx = i
        break

services_code = [
    "from datetime import datetime, date, timedelta\n",
    "import re\n",
    "from django.db.models import Sum\n",
    "from django.utils import timezone\n",
    "from django.conf import settings\n",
    "from django.contrib import messages\n",
    "from .models import JobCard, Production, Dispatch, ChangeLog, EditOverrideRequest\n",
    "from .constants import AUDIT_CONFIG\n",
    "from core.views import add_unique_message\n",  # ensure_edit_lock_allowed uses add_unique_message
    "\n"
]

services_code.extend(lines[start_idx:end_idx])

with open(services_path, "w", encoding="utf-8") as f:
    f.writelines(services_code)

# Now we remove AUDIT_CONFIG and the services from views.py
audit_start_idx = None
for i, line in enumerate(lines):
    if "AUDIT_CONFIG = {" in line:
        audit_start_idx = i
        break

new_views = lines[:audit_start_idx] + lines[end_idx:]

imports = [
    "from .constants import AUDIT_CONFIG\n",
    "from .services import (\n",
    "    format_audit_value, normalize_colour_notation, extract_total_colors,\n",
    "    compute_planned_minutes, get_remaining_planned_minutes, build_audit_snapshot,\n",
    "    build_change_summary, log_change, user_has_entity_permission,\n",
    "    user_can_archive_records, user_can_bypass_edit_lock, get_record_edit_lock_days,\n",
    "    get_record_edit_lock_cutoff, record_is_time_locked, get_valid_override,\n",
    "    ensure_edit_lock_allowed, get_accessible_entities, validate_delete_allowed,\n",
    "    validate_restore_allowed, archive_record, restore_record_state,\n",
    "    run_bulk_archive, run_bulk_permanent_delete\n",
    ")\n"
]

# Insert imports after the other core imports
insert_idx = 0
for i, line in enumerate(new_views):
    if "from workflow.services import start_production" in line:
        insert_idx = i + 1
        break

final_views = new_views[:insert_idx] + ["\n"] + imports + ["\n"] + new_views[insert_idx:]

with open(views_path, "w", encoding="utf-8") as f:
    f.writelines(final_views)

print("Done extracting services.")
