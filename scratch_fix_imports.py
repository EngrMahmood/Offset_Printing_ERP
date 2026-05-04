import os

views_path = r"d:\Offset_Printing_ERP\core\views.py"
services_path = r"d:\Offset_Printing_ERP\core\services.py"

# Read views
with open(views_path, "r", encoding="utf-8") as f:
    views_lines = f.readlines()

# Find add_unique_message
start_idx = -1
end_idx = -1
for i, line in enumerate(views_lines):
    if "def add_unique_message(request, level, text):" in line:
        start_idx = i
    if start_idx != -1 and "messages.add_message(request, level, text)" in line:
        end_idx = i + 1
        break

func_lines = views_lines[start_idx:end_idx]

# Remove func from views
new_views = views_lines[:start_idx] + views_lines[end_idx:]

# Update views import list to include add_unique_message
for i, line in enumerate(new_views):
    if "from .services import (" in line:
        new_views.insert(i + 1, "    add_unique_message,\n")
        break

with open(views_path, "w", encoding="utf-8") as f:
    f.writelines(new_views)

# Read services
with open(services_path, "r", encoding="utf-8") as f:
    services_lines = f.readlines()

# Remove circular import
new_services = [line for line in services_lines if "from core.views import add_unique_message" not in line]

# Insert func_lines after imports
insert_idx = -1
for i, line in enumerate(new_services):
    if "from .constants import AUDIT_CONFIG" in line:
        insert_idx = i + 1
        break

final_services = new_services[:insert_idx] + ["\n"] + func_lines + ["\n"] + new_services[insert_idx:]

with open(services_path, "w", encoding="utf-8") as f:
    f.writelines(final_services)

print("Fixed circular import.")
