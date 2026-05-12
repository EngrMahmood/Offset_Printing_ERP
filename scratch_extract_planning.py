import os

views_path = r"d:\Offset_Printing_ERP\planning\views.py"
services_path = r"d:\Offset_Printing_ERP\planning\services.py"

with open(views_path, "r", encoding="utf-8") as f:
    lines = f.readlines()

def is_def(line):
    return line.startswith("def _")

blocks_to_remove = []
services_lines = []

i = 0
while i < len(lines):
    if is_def(lines[i]):
        start = i
        i += 1
        # Read until next outdent
        while i < len(lines):
            line = lines[i]
            if line.strip() == "" or line.startswith(" ") or line.startswith("\t"):
                i += 1
            else:
                break
        end = i
        blocks_to_remove.append((start, end))
        services_lines.extend(lines[start:end])
        services_lines.append("\n")
    else:
        i += 1

# Remove blocks backwards
new_views = lines[:]
for start, end in reversed(blocks_to_remove):
    del new_views[start:end]

# Services file content
services_header = [
    "import io\n",
    "import json\n",
    "from datetime import datetime, date\n",
    "from decimal import Decimal\n",
    "from collections import defaultdict\n",
    "from django.db import transaction\n",
    "from django.db.models import Sum, Q\n",
    "from django.utils import timezone\n",
    "from core.models import Machine, Department, Material\n",
    "from .models import PlanningJob, PoDocument, SkuRecipe\n",
    "from workflow.services import _append_unique_note_line\n",
    "\n"
]

with open(services_path, "w", encoding="utf-8") as f:
    f.writelines(services_header + services_lines)

# Update views.py imports
imports = [
    "from .services import (\n",
    "    _user_is_admin, _planning_status_filter_values, _parse_date_filter,\n",
    "    _build_job_card_pdf_bytes, _sku_key, _missing_required_master_fields,\n",
    "    _sync_new_sku_requirement, _build_recipe_map, _to_optional_positive_int,\n",
    "    _to_optional_decimal, _sanitize_po_payload_items, _po_payload_items,\n",
    "    _annotate_items_with_recipe, _deduplicate_po_items_by_sku,\n",
    "    _history_repeat_new_counts, _sync_repeat_jobs_from_po,\n",
    "    _sync_new_jobs_for_approved_sku, _merge_po_items_for_existing_po,\n",
    "    _collect_pending_sku_rows\n",
    ")\n"
]

insert_idx = 0
for j, line in enumerate(new_views):
    if "from .models import" in line:
        insert_idx = j + 1
        break

final_views = new_views[:insert_idx] + ["\n"] + imports + ["\n"] + new_views[insert_idx:]

with open(views_path, "w", encoding="utf-8") as f:
    f.writelines(final_views)

print("Planning refactored.")
