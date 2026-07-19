# Machine Planning Report V2 — Implementation Plan

_Date: 2026-07-19_

Builds on the existing routing engine (`core/machine_routing.py`) and
`build_machine_planning_context` (`reports/services.py`). Scope = the six changes below plus suggested
smart additions. Grounded in current code, so each item names the file/function to touch.

---

## 1. Partially-produced jobs: show remaining state, don't hard-exclude

**Decision (revised):** do **not** drop every job that has a production entry. Only fully-**completed**
jobs leave the report. A job that has started but still has balance left stays visible as a
**"Partially produced"** run showing the *remaining* work, and the planner decides whether to exclude it.

**Now:** `build_machine_planning_context` starts from
`PlanningJob.objects.filter(is_active=True).exclude(status='completed').exclude(job_process_type='cut_and_pack')`,
and shows full order/sheet quantities regardless of what's already been printed.

**Change:**
- **Fully exclude** only completed jobs (and `cut_and_pack`), as today.
- **Classify each remaining job's production state:**
  - `not_started` — no production entries → normal plannable run (full quantity).
  - `partially_produced` — has production entries but balance remains → keep in the report, but display
    **remaining sheets / impressions** (from `balance_qty` and remaining vs `actual_sheet_required`,
    already aggregated in the builder as `total_balance_qty`), the **actual machine** run so far, and a
    "Partially produced — X% done, N sheets left" badge.
  - `done` — **remaining balance ≤ 5% of order qty** → treat as completed, exclude (5% tolerance so tiny
    leftovers don't clutter the plan).
- **Planner exclude/hold toggle:** a per-job control (persisted, like the JC selection in item 3) that lets
  the planner remove a partially-produced job from the active plan or keep re-planning its balance. Default
  = show remaining. Reuse/extend the existing on-hold flag (`is_on_hold`, already counted in the builder)
  rather than inventing a new field where possible.
- Keep all of this in one helper (`_plannable_jobs(...)` / `_production_state(job)`) so the report, merge
  grouping, and exports stay consistent.

**Verify:** a job printed in full → excluded; a job half-printed → shows as "Partially produced" with the
remaining sheet count and its actual machine, and the planner's exclude toggle removes it from the active
plan; a never-started job → normal full-quantity run.

---

## 2. Combine all same-SKU jobs, show stages, planner decides

**Requirement (confirmed):** show **all** jobs combined by **same SKU**, with each job's **stage**
displayed. Do not auto-drop split-stage jobs — the planner manually selects/deselects which JCs run
together (see item 3). Same-SKU grouping is the default; the planner is the final arbiter.

**Change (in the `sku_groups` build loop):**
- Group by `(pool, sku)` across the plannable set from item 1 (exclude completed; keep partially-produced
  with their **remaining** quantities). Do **not** filter out JCs just because their stage differs.
- Display each JC's **stage** and **production state** (`planning`, `qc_approved`, `released`,
  `in_production`, `partially_produced`) inside the combined row so the planner can judge before selecting.
- Merged sheet / finish totals sum the **remaining** balance of partially-produced JCs, not full order qty,
  so consolidated numbers reflect what's left to print.
- The report output reflects the planner's current selection (item 3): deselected JCs are excluded from the
  merged totals and report/exports, but still shown-as-excluded in the planner console.

**Verify:** SKU with 3 JCs where 1 is in production → combined run shows the 2 plannable JCs only; SKU with
all JCs in production → absent from the report.

---

## 3. Interactive JC selection for combined runs (planner opt-in/opt-out)

**Requirement:** for each combined run show the planning stage and a JC checklist; the planner chooses
which JCs actually run together. Deselecting a JC must re-compute the combined run (sheets, finish qty,
sequence, savings) accordingly.

**Change:**
- Add a per-JC checkbox inside each combined row (default all checked), keyed by `planning_job_id` /
  `jc_number`, with the JC's stage shown next to it.
- **Persist as a shared selection** in a small `MachinePlanningSelection` table (SKU/pool + deselected JC
  set). Everyone sees the same state; the selection is not per-user.
- **Permissions:** only **Planner and Admin** (and superuser) can change it — gate the selection endpoint
  on `profile.role in {'planner', 'admin'}`. All other roles render the checkboxes read-only/disabled.
- **Deselected-JC behaviour:** excluded from the Machine Planning **report output and exports** (removed
  from merged totals and AI scoring), but still **shown-as-excluded in the planner console** (greyed row /
  "Excluded by planner" tag) so the planner always sees what was taken out.
- On change, POST the selection and re-render via the same cache-bust path as priority (item 6) so totals
  update immediately.

**Verify:** planner deselects one JC of a 3-JC combined run → report/exports recompute to 2 JCs with
updated sheet/finish totals; the deselected JC stays visible in the planner UI marked "excluded"; a
non-planner/non-admin user sees the same state but cannot change it.

---

## 4. Show combined pools only — remove the duplicate individual GTO tabs

**Now:** `machine_reports` ends up keyed by a mix of **pool labels** (`GTO1A, GTO1B`) *and* **single
machine names** (`GTO1A`) — because when a job has an explicit machine the builder sometimes keys on the
literal name, and when it routes it keys on the pool label. Result: both an individual tab and a combined
tab for the same fleet.

**Change:**
- Always key `machine_reports` on the **pool** (`pool_key` / pool label), never the individual machine
  name. Route every plannable job through `route_job` / `find_pool_for_machine` and collapse to the pool.
- Keep the member machine names *inside* the pool label (`GTO1A, GTO1B` / `GTO2A, GTO2B, GTO2C`) so the
  planner still sees which units are combined — but only **one tab per pool**.
- SM74 remains its own pool/tab.

**Verify:** tabs list exactly one entry per active pool (GTO1 pool, GTO2 pool, SM74) — no standalone
single-machine duplicates.

---

## 5. SM74 / GTO override rules (respect planner, warn on size violations)

**Two rules:**

a. **Planner picks SM74 even though it fits GTO → honour the planner.** When a job has an explicit machine
   assignment, the builder already routes via `find_pool_for_machine` and keeps that pool. Ensure the
   size-gate does **not** override an explicit planner choice — explicit assignment always wins. Add a flag
   `planner_override=True` on the row so it's visible why it's on SM74.

b. **Planner puts a job on a GTO that exceeds the GTO max sheet size → warn to correct.** Add a validation
   pass: for any job explicitly assigned to a GTO pool, compare its parsed sheet size
   (`parse_sheet_size_mm`) against that pool's max `_fits(...)`. If it doesn't fit, attach a
   `size_warning` to the row and render a visible warning badge ("Exceeds GTO max — move to SM74"). Do not
   silently re-route; the planner corrects it. Optionally surface a count of size-violation warnings in the
   summary cards.

**Verify:** oversized job on GTO shows a warning badge and is *not* auto-moved; a fits-GTO job the planner
put on SM74 stays on SM74 with an override note.

---

## 6. Priority not updating immediately — cache bust (carried from prior plan, still open)

**Confirmed root cause:** `planning_job_priority_update` (`planning/views.py:2214`) saves the priority but
never clears the report cache. The report engine caches the machine-planning payload for 300s
(`cache_timeout=300`, keyed by user+filters), so the JS `window.location.reload()` after a priority change
re-reads the stale cached page — hence "sometimes updates, sometimes only after several refreshes."

**Change:** add a cache-version bust. Maintain a `reports:machine-planning:version` key included in
`_cache_key`; bump it on any `PlanningJob` save (priority, machine, stage) and on `Machine` master save.
`planning_job_priority_update`, the JC-selection endpoint (item 3), and the Machine admin save all bump it.
Optionally drop `cache_timeout` for this report to ~30s as a backstop.

**Verify:** change priority as a planner → one reload reflects the new order every time.

---

## 7. Master-data colour change (operational_colors 2→1) must re-flow planning live

**Question:** if admin changes a GTO2 machine's `operational_colors` from 2 to 1 in master data, does the
Machine Planning report change accordingly?

**Current state:** the *routing logic already handles it.* `build_pools` (`core/machine_routing.py`) folds
a machine with `effective_colors == 1` and `default_colors > 1` into the single-colour (GTO1) pool, and
`route_job` then routes 2-colour jobs to the remaining GTO2 members. The model fields
(`operational_colors`, `effective_colors`, `is_under_maintenance`) already exist and are used. So the
degraded machine moves GTO2 → GTO1 and the double-colour pool shrinks by one.

**Gap (must fix):** saving a `Machine` does **not** invalidate the report cache (same 300s cache as
item 6), so the change would only surface after the cache expires — not immediately. Setting
`operational_colors = 0` (maintenance) has the same lag before the machine drops out of all pools.

**Change:**
- Include **`Machine` save** in the cache-version bust from item 6, so any master-data edit
  (`operational_colors`, `default_colors`, sizes, `machine_group_code`) re-flows the report on the next
  reload.
- Forward-only, as designed: only plannable/not-yet-executed jobs re-group; executed/in-production runs are
  untouched.

**Verify:** set a GTO2 machine to `operational_colors = 1`, reload once → that machine appears in the GTO1
pool tab, GTO2 pool shows one fewer member, and pending 2-colour jobs route to the remaining GTO2
machines. Set it to `0` → machine drops into the maintenance list and out of all pools.

---

## Suggested smart additions

- **Sequence drag-and-drop / manual pin:** let the planner reorder or pin a run above the AI sequence, with
  the manual order persisted and respected over the AI score.
- **Plate-life awareness:** the `Machine.plate_life_impressions` field already exists — flag runs whose
  impressions exceed one plate set so the planner expects a plate change mid-run.
- **Capacity/overload banner per pool:** you already compute `utilization_pct` against 22h/day; show an
  "overloaded — spills to next day" banner when a pool exceeds 100%, and estimate the spill date.
- **Colour wash-up sequencing:** order runs light→dark within a pool to minimise wash-ups (partly in the AI
  reason already); make it an explicit sequencing rule.
- **"Ready to release" vs "blocked" split:** mark runs missing plate set / AWC (data already in the audit
  gaps logic) so planners don't schedule work that can't be released.
- **Delivery-date SLA alongside PO age:** current status pill is PO-age driven; add delivery-date urgency so
  a near-delivery job ranks up even if the PO is fresh.
- **What-if merge preview:** before committing a JC selection, show the setup-time saved / sheets delta so
  the planner sees the benefit of combining.

---

## Delivery order

1. **Item 6 + 7** (cache bust on `PlanningJob` *and* `Machine` save) — smallest, fixes both the priority
   lag and the master-data re-flow lag in one change.
2. **Item 1 + 2** (plannable-jobs filter + same-stage combine) — one shared helper, correctness foundation.
3. **Item 4** (combined-only tabs) — display cleanup, depends on 1/2.
4. **Item 5** (override + size warnings).
5. **Item 3** (interactive JC selection) — largest, needs persistence + endpoint.
6. Smart additions, prioritised with you.

## Resolved decisions (confirmed)

- **Combine + stages + planner chooses (items 2/3):** show **all** jobs combined by **same SKU** with each
  job's **stage** displayed. Do not auto-exclude split-stage jobs — the planner manually selects/deselects
  which JCs run together. Same-SKU grouping is the default; the planner is the final arbiter.
- **Selection visibility & permissions:** the JC selection is **shared** — everyone sees the same
  selection state — but **only Planner and Admin roles can change it** (others are read-only). Gate the
  selection endpoint on `profile.role in {'planner', 'admin'}` (superuser allowed).
- **Deselected JC behaviour:** a deselected JC is **excluded from the Machine Planning report output and
  exports**, but **remains visible (marked excluded) in the planner UI** so the planner always knows which
  jobs they took out. i.e. hidden from the report, shown-as-excluded in the planner console.
