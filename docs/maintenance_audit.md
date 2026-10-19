# Maintenance module — audit and upgrade log

Scope: the existing `maintenance/` Django app. Nothing was rebuilt; the work is incremental and reviewable.
Status below separates what is **done** (Phases 1–3 and the downtime/workflow part of Phase 2) from what is **not yet done** (Phases 4–5).

## 1. Existing architecture (as inspected)

- Models: `MaintenanceRecord` (ticket), `MaintenanceSparePart`, `MaintenanceServiceJob`, `MachineDowntime`, `MaintenanceApproval`, `PreventiveMaintenancePlan`, `MaintenanceAttachment`, `MaintenanceActivityLog`, `FaultCategory`.
- Flow: operator files a complaint (`REPORTED`) → engineer triage (`DIAGNOSED`) → generic status move → `COMPLETED` → `VERIFIED` → `CLOSED`.
- Access: `nav.maintenance` soft-coded permission (module access) plus hard-coded role tuples in decorators.
- Integrations (preserved): Supply Chain `ItemRequest` (types MNT/SRV, department Maintenance), `core.notifications.notify_roles/notify_users`, Google Sheets registry (`sheets_sync/registry.py`, serializers), `core.Machine`.
- `core.Machine.is_under_maintenance` is `operational_colors == 0`. **The maintenance module never changes it**, and planning/routing read it. This was left untouched on purpose: machine availability for planning is not redefined by whether a ticket exists.

## 2. Confirmed defects (verified in code, before changes)

| # | Defect | Status |
|---|---|---|
| 1 | Status rules lived in a dict inside `views.py` and were re-implemented in the template. | **Fixed** — `maintenance/workflow.py` is the only rule set. |
| 2 | Downtime closed when a ticket reached `COMPLETED`, i.e. before any test run. | **Fixed** — closes at `VERIFIED` (passed test) or on confirmed restoration at cancellation. |
| 3 | Cancelling a ticket never closed its downtime (interval stayed open forever) and needed no reason. | **Fixed** — reason mandatory; for stopped-machine tickets the user must say "running again" or "still down"; "still down" needs another active stopped ticket and the interval is handed to it. |
| 4 | Two tickets on one machine opened two overlapping intervals (double counting). | **Fixed** — one open interval per machine; later tickets reuse it. |
| 5 | Closing one ticket could end downtime while another fault still stopped the machine. | **Fixed** — interval is handed over to the other stopped ticket instead of closing. |
| 6 | "Stopped machine" was only implied by a downtime row; the complaint question was inverted ("Is the machine running? — No"). | **Fixed** — explicit `machine_stopped` flag and a clear "Has the machine stopped?" question. |
| 7 | Any user with module access could read every ticket, add comments/attachments to any ticket, open downtime and report screens, and export (incl. costs). | **Fixed** — non-staff see only tickets they reported; downtime, PM and reports screens and export are staff-only. |
| 8 | Staff checks disagreed: templates used `is_superuser`, the decorator did not, so a superuser could see buttons the server refused. | **Fixed** — one helper, `maintenance/permissions.py`. |
| 9 | No diagnosis / repair action / test result / verifier recorded; completion needed nothing. | **Fixed** — new fields; completion requires diagnosis and repair action, verification requires a passed test and records who/when. |
| 10 | Labour hours could be negative or free text. | **Fixed** — validated, stored as Decimal. |
| 11 | Double-click or refresh on the complaint form created duplicate tickets/intervals. | **Fixed** — same user/machine/description within 2 minutes returns the existing ticket. |
| 12 | Attachments had no type or size validation. | **Fixed** — allow-list of extensions, 10 MB limit (complaint photo and attachment form). |
| 13 | Soft-deleting a ticket left its downtime running. | **Fixed**. |
| 14 | Notifications could fail inside the business transaction. | **Fixed** — sent after commit, failures logged, never undo the change. |
| 15 | Preventive maintenance never ran: the command was not scheduled; impression plans were stored but never processed; no per-occurrence idempotency. | **Fixed** — in-process scheduler (project convention, `MAINTENANCE_INPROCESS_SCHEDULER`), manual fallback command, DB-enforced one ticket per plan occurrence, impression plans counted from Production entries, per-plan failure isolation, overdue reminders. |
| 16 | Costs: `unit_price × received_qty` silently becomes 0 when missing; labour rate not snapshotted; no work sessions. | **Left as is by decision** (cost matters deferred). |
| 17 | Reports: no date filter, no MTBF/PM compliance, no repair-vs-waiting. | **Fixed** for everything except cost (date range, downtime hours, MTTR, repeat faults, repair vs waiting, PM compliance). The existing cost column was kept unchanged and still queries per machine. MTBF is not computed (needs operating hours). |
| 18 | `MaintenanceApproval` / `PENDING_APPROVAL` / `review_record` unreachable. | **Decision: no approval step at this stage.** The ticket default status is now `REPORTED`; the approval model and status are kept for audit history and future use. |

## 3. What changed

- `maintenance/workflow.py` (new): `transition_record()` plus `TRANSITIONS`, atomic, row-locked, audit-logged, friendly `TransitionError`; action labels and plain-English status help for the UI.
- `maintenance/permissions.py` (new): `STAFF_ROLES`, `is_maintenance_staff`, `can_view_record`, `require_record_access`.
- `maintenance/services.py`: `open_downtime` (one interval per machine), `release_downtime` (close or hand over, records who/why), new notifications (repair completed → verify, waiting, assigned).
- `maintenance/models.py` + migration `0005_workflow_repair_and_downtime_fields` (additive, with a backfill): `machine_stopped`, `root_cause`, `repair_action`, `test_result`, `verified_by`, `verified_at` on the ticket; `closed_by`, `closed_reason` on downtime.
- `maintenance/views.py`, `forms.py`, `urls.py`, `decorators.py`, `record_detail.html`: use the above; new endpoint `records/<id>/repair/` for diagnosis, repair action, labour and test result.
- Tests: `maintenance/tests/` (46 tests).

## 4. Migration impact

Additive only (new nullable/default columns, no data rewritten except the backfill that sets `machine_stopped=True` on tickets that already had a breakdown downtime). Existing statuses and values are unchanged. Tickets already `COMPLETED` keep their already-closed downtime. Safe on SQLite and PostgreSQL.

## 5. Decisions taken

1. No approval of costs / outsourced repairs at this stage.
2. Only admin and manager (and superusers) may reopen a closed ticket.
3. Cost-related work (labour rate snapshot, work sessions, cost reconciliation) is deferred.
4. Item Requests (`nav.item_request`) are open to every role, including custom roles (granted by `seed_access_control` on every start).

## 6. Additional changes (second pass)

- Preventive maintenance: models `MaintenanceChecklistItem`, plan fields (instructions, checklist, responsible technician), ticket fields `pm_plan` / `pm_due_key` with a unique constraint; migration `0006_pm_occurrences_checklist`.
- `maintenance/scheduler.py`, `maintenance/reporting.py`, PM plan edit / enable-disable screens, checklist on the ticket, PM completion needs the checklist ticked (no root cause required for scheduled PM), `last_done_at` set only when a PM ticket is verified.
- Dashboard (operator vs engineer), record-list filters (date range, technician), date-filtered reports, duplicate-safe "Raise demand".
- Tests: 82 in `maintenance/tests/`.

## 7. Still open

- Costing (deferred by decision) and technician work sessions.
- MTBF (needs operating hours).
- Impression plans are only as accurate as the Production entries behind them; there is no dedicated machine counter.
