# Machine Planning Upgrade & Permission Fix — Implementation Plan

_Date: 2026-07-19_

This plan covers two things:

1. **Bug fix** — Machine Planning report (and Audit) show data only to Admin, not to Manager and below.
2. **Feature build** — Colour-based machine master data, colour-driven machine merging in planning
   (GTO1×2 / GTO2×3), min/max print size, maintenance handling, and planned-vs-actual machine
   tracking that feeds the SKU master.

---

## PART A — Permission Bug (fix first, ship independently)

### Root cause (confirmed for Machine Planning)

- The report engine (`reports/report_engine/engine.py → _has_access`) authorises every report with the
  **Django model permission** `user.has_perm('core.view_reports')`. In practice only the superuser/admin
  holds that permission.
- The rest of the app authorises with **`request.user.profile.role`** (`manager`, `planner`,
  `production_manager`, …) via `core/navigation.py` (`REPORTS_NAV_ROLES`, `AUDIT_NAV_ROLES`).
- These two authorisation systems have diverged. Managers see the *link* (role-based nav) but fail the
  *data* check (Django-perm based engine).
- When Machine Planning got its custom template, `reports/views.py → report_detail` now calls
  `run_report(...)` and swallows the resulting `PermissionDenied` with `except Exception: pass`, so the
  page renders **with empty data** instead of a clear error. Admin passes the perm check, so admin sees data.

### Fix

1. **Unify authorisation on `profile.role`.** In `engine.py`, change `_has_access` so that, in addition to
   `has_perm`, it accepts users whose `profile.role` is in the reports-allowed set (reuse
   `REPORTS_NAV_ROLES` from `core/navigation.py` — import the set, do not duplicate it). Superusers keep
   full access. This is the smallest change and keeps a single source of truth for roles.

   _Alternative (if you prefer Django perms):_ create Django Groups for each role and assign
   `core.view_reports` to the manager/planner/production_manager groups via a data migration. More moving
   parts; not recommended given the app already standardises on `profile.role`.

2. **Stop silently swallowing the auth error.** In `report_detail`, catch `PermissionDenied` explicitly and
   surface a proper 403 / "you don't have access" message, so a genuine permission problem is never
   disguised as an empty report again. Keep a narrow `except` for other runtime errors only, and log them.

3. **Audit page — confirm before fixing.** Audit is a separate app (`audit/views.py`, `@login_required`
   only; managers already in `AUDIT_NAV_ROLES`) with its own `get_audit_data()` and no `has_perm` gate, so
   the same root cause does not obviously apply. First reproduce as a manager and check: (a) does the page
   403/redirect, or render empty? (b) browser console / server log. Likely candidates: a shared base
   template or context processor added in the same release, or an exception in `get_audit_data()` for
   non-admin sessions. Fix once reproduced — do not blind-patch.

### Verification for Part A

- Log in as a Manager, Planner, Production Manager, and a below-role user; open Machine Planning and Audit.
  Confirm data renders for the intended roles and a clean 403 for the rest.
- Add/adjust a regression test in `reports/tests.py` that asserts a non-superuser with `role='manager'`
  gets a populated `data` payload from `run_report('machine-planning', ...)`.

---

## PART B — Machine Master Data: colour, size, maintenance

### B1. Extend the `Machine` model (`core/models.py`)

Add fields:

| Field | Type | Purpose |
|-------|------|---------|
| `default_colors` | PositiveSmallInteger (e.g. 1 or 2) | Nominal colour capacity of the machine |
| `operational_colors` | PositiveSmallInteger | Currently working colour units. `0` ⇒ under maintenance; `1` on a 2-colour machine ⇒ temporarily treated as single-colour |
| `min_print_length_mm` / `min_print_width_mm` | Decimal | Smallest sheet the machine can run |
| `max_print_length_mm` / `max_print_width_mm` | Decimal | Largest sheet the machine can run |
| `machine_group_code` | Char (e.g. `GTO1`, `GTO2`) | Stable code used to group machines by colour class |

Add helper properties:

- `is_under_maintenance` → `operational_colors == 0`
- `effective_colors` → `operational_colors` (falls back to `default_colors` when not overridden)

Fleet reference (from your description): **2 single-colour** and **3 double-colour** machines →
`GTO1 × 2` and `GTO2 × 3`.

Expose all new fields in the Machine admin/master-data form and `core/admin.py`.

Migration: `core/migrations/00XX_machine_color_size_fields.py` with sensible defaults so existing rows
stay valid; then a one-off data step to set `default_colors`/`operational_colors`/group codes for the 5
real machines.

### B2. Machine routing logic (confirmed rules)

Data source for colour = **planning colour spec** (`color_spec_display` / recipe). Routing is decided by
**print sheet size first, then colour**, in this order:

1. **Size gate first (SM74 override).** If the job's print sheet size exceeds the GTO groups' max print
   size, route it to the **SM74 5-colour machine** regardless of colour count. SM74 has the largest sheet
   size, so it handles anything (1 → 5+ colour) that is too big for the GTOs. Size is the primary
   constraint, colour is secondary.
2. **1+1 handling.** `1+1` (front/back) means **2 passes of 1 colour**, so it is a **1-colour** job → route
   to the GTO1 (single-colour) pool. Generalise: `X+X` front/back = single-colour-class, multi-pass.
3. **1-colour jobs** → GTO1 pool.
4. **2-and-above colour jobs** that fit the GTO size range → GTO2 (double-colour) pool, expressed as
   **number of passes** (e.g. a 3-colour job = 2 passes on a double-colour machine; 4-colour = 2 passes;
   odd colours round up). Compute and display `passes = ceil(colours / effective_colors_of_group)`.

So the grouping in `build_machine_planning_context` (`reports/services.py`) changes from literal
`machine_name` string to **size-gate → colour-class pool**, and each merged row carries a pass count.

### B3. Show which machines are combined (named)

Instead of a generic `GTO1×N`, the report must **name the actual machines merged** into each pool:

- 1-colour jobs combined across the single-colour machines → show e.g. **`GTO1A, GTO1B`** (or
  `GTO1 (A+B)`).
- 2-colour jobs combined across the double-colour machines → show e.g. **`GTO2A, GTO2B, GTO2C`** (or
  `GTO2 (A+B+C)`).

This means each real machine needs a stable per-unit identity (name/suffix) so the report can list the
members of a pool, not just a count. Use the machine `name` plus `machine_group_code`; the pool label is
built by joining the member machine names.

Pool membership rules (as before):
- A 2-colour machine with `operational_colors == 1` drops into the GTO1 pool for that period and is listed
  among the GTO1 members.
- A machine with `operational_colors == 0` is **excluded** (under maintenance) and shown separately as
  unavailable/maintenance capacity.

### B4. Master-data changes apply forward-only, and auto-update

- When admin edits machine master data (colours, size, maintenance), the system uses the **new values from
  that moment forward**. **Already-executed planning/production is not retro-changed**; only
  **not-yet-executed** (pending/unexecuted) jobs re-plan on the new data.
- The report/planning must **auto-update** to reflect master-data edits — no manual rebuild. This ties
  directly to the cache-invalidation work in Part D: a Machine master save must bust the machine-planning
  cache so pending jobs re-group immediately.

### Verification for Part B

- Unit tests: given the 5-machine fleet with set `operational_colors`, assert pools list the correct member
  machine names (`GTO1A, GTO1B` / `GTO2A, GTO2B, GTO2C`), that a degraded 2-colour machine appears in GTO1,
  and a `0`-colour machine is excluded as maintenance.
- Size-gate test: an oversized job routes to SM74 regardless of colour; a 3-colour in-range job routes to
  GTO2 with `passes = 2`; a `1+1` job routes to GTO1.
- Forward-only test: editing master data changes grouping for pending jobs but leaves executed jobs
  untouched.

---

## PART C — Planned vs Actual Machine Tracking + SKU master learning

Goal: plan generically (by colour group), but track which **specific** machine actually ran the job, and
learn a default machine per SKU.

1. **Capture actual machine at production entry.** `Production.machine` (FK) already stores the real
   machine. Ensure the production entry form requires/records it (`production/views.py`,
   `printing_entry_helpers.py`).

2. **Write back to the SKU master.** On production save, update the job's `SkuRecipe.machine_name`
   (and/or a new `last_actual_machine` / `preferred_machine` field on `SkuRecipe`) with the machine that
   actually ran it — so future planning can suggest a default machine per SKU. Do this in a signal or the
   production-save service, guarded so manual overrides aren't clobbered silently.

3. **Show plan vs actual in the report.** In the Machine Planning template, display both the **planned
   machine/colour group** and the **actual production machine** side by side, to track per-machine
   performance independent of the generic plan. The data is already available:
   `PlanningJob.machine_name` (plan) and `Production.machine` (actual) aggregated per job/SKU.

### Verification for Part C

- Enter a production run on a machine different from the plan; confirm the report shows both plan and
  actual, and the SKU master's default machine updates (without overwriting a manual lock).

---

## PART D — Priority not updating immediately (cache-invalidation bug)

### Root cause (confirmed)

The report engine caches the machine-planning payload for **300 seconds**
(`cache_timeout=300` in `builtin_reports.py`, `cache.set(..., timeout=...)` in `engine.py`), keyed by
`(slug, user_id, filters)`. When a planner changes a job's priority, the report keeps serving the stale
cached payload until the key expires or changes — which is exactly the reported symptom: "sometimes it
updates, sometimes only after multiple refreshes."

### Fix

- **Bust the cache on write.** Add a signal (or hook the priority/planning save path) so that saving a
  `PlanningJob` — priority or any planning-relevant field — and saving `Machine` master data clears the
  cached machine-planning payloads. Simplest robust approach: maintain a per-report cache **version** key
  and include it in `_cache_key`; bump it on any relevant write, which invalidates all user/filter variants
  at once (avoids trying to enumerate `reports:engine:*`).
- Optionally lower `cache_timeout` for machine-planning (e.g. 30–60s) as a backstop, but proper
  invalidation is the real fix — planners need immediate reflection of priority edits.

### Verification for Part D

- Change a job's priority as a planner; reload once; assert the new order/priority appears without waiting
  or multiple refreshes. Add a test asserting the cache version bumps on `PlanningJob` save.

---

## Suggested delivery order

1. **Part A** (permission fix) — highest urgency, small, ship on its own with the regression test.
2. **Part D** (priority cache invalidation) — quick, high daily impact for planners.
3. **Part B1** (model + migration + master-data form).
4. **Part B2/B3** (size-gate + colour-pool routing, named combined machines in the report).
5. **Part C** (actual capture, SKU write-back, plan-vs-actual display).

## Resolved decisions (from review)

- **Colour source** = planning colour spec. `1+1` = front/back = 2 passes of 1 colour ⇒ single-colour
  machine. 2+ colour jobs run on double-colour machines expressed as passes (3-colour = 2 passes). No new
  `required_colors` field needed unless the spec proves unreliable in data.
- **Size is the primary constraint.** Oversized jobs go to **SM74 (5-colour, largest sheet)** regardless of
  colour; SM74 covers 1 → 5+ colour by size.
- **Report shows named combined machines** (`GTO1A, GTO1B` / `GTO2A, GTO2B, GTO2C`), not just counts.
- **Master data is forward-only**: admin edits apply from the edit moment onward to not-yet-executed
  planning; executed history is untouched. Planning **auto-updates** (via cache invalidation) on master
  save.
- **SKU master auto-updates** from actual production machine (Part C) — automatic, per the requirement,
  with a guard so an explicit manual lock isn't overwritten.
