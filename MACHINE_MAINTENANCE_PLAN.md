# Machine Maintenance Module — Implementation Plan

**Location:** new Django app `maintenance` (registered in `INSTALLED_APPS`, mounted at `/maintenance/`), reusing `core.Machine`, `core.Machine.work_schedule` (for downtime), and the existing `supply_chain.ItemRequest` module for spare parts and outsource services.
**Goal:** Replace the "Offset Maintenance" Google Sheet with a proper ERP module that logs every maintenance activity against a machine, drives spare-part and repair-service procurement through the existing Item Request workflow, tracks machine downtime, and reports maintenance KPIs.

---

## 1. Why a new app (not an extension of supply_chain)

Maintenance is its own domain (fault logging, scheduling, downtime, spares). It *consumes* the Item Request module but shouldn't live inside it. A dedicated `maintenance` app keeps models, permissions, and URLs clean, and mirrors how `planning`, `production`, `qc` are separated. It reads `core.Machine` and creates `supply_chain.ItemRequest` records — no schema changes to those apps except two small additions (§6).

---

## 2. Mapping the current Google Sheet → the module

| Sheet column | Where it lives in the module |
|---|---|
| S.no | `MaintenanceRecord.record_no` (auto, e.g. `MNT-2026-0042`) |
| Date | `reported_date` |
| Machine | FK → `core.Machine` |
| Fault Type | `fault_type` (Mechanical / Electrical / Electronic / Pneumatic / Hydraulic / Other) |
| Maintenance Type | `maintenance_type` (Preventive / Corrective / Breakdown / Predictive) |
| Priority | `priority` (Low / Medium / Major / Critical) |
| Fault Description | `fault_description` |
| Propose Solution | `proposed_solution` |
| Spare Parts Needed (Yes/No) | `spare_parts_needed` (bool) → drives linked Item Requests |
| Spare Parts Detail | line items on `MaintenanceSparePart` / linked `ItemRequest` |
| Repair Needed (Yes/No) | `repair_needed` (bool) |
| Repair Details | `repair_details` + linked service `ItemRequest` when outsourced |
| Demand Raise (Yes/No) + Demand Date | derived from linked `ItemRequest` existence + its `request_date` |
| Indent / PR / PO / Receive Date | already tracked in `ItemProcurementTimeline` of the linked `ItemRequest` — **no duplication** |
| Work Start Date / Work End Date | `work_start_at` / `work_end_at` |
| Total DownTime | `MachineDowntime` computed field (§5) |
| Remarks | `remarks` |

**Key improvement over the sheet:** procurement columns (Indent/PR/PO/Receive) are not re-typed here — they come from the linked Item Request, so there is one source of truth.

---

## 3. New requested capabilities

1. **In-house vs Outsource** — added as `execution_type` on the record (and per repair job), which decides the downstream path:
   - *In-house* → work done by internal team; may still raise **spare-part** Item Requests.
   - *Outsource* → raises a **service** Item Request (type `SRV`) to the vendor, and optionally spare-part requests too.
2. **Item Request / Service linkage** — spare parts and outsourced repairs generate Item Requests using the existing approval + procurement workflow, back-linked to the maintenance record so status flows both ways.
3. **Downtime tracking** — every stoppage recorded as a `MachineDowntime` interval, netted against the machine's work schedule so downtime reflects only scheduled running time.
4. **Full activity history** — immutable `MaintenanceActivityLog` for every status change, assignment, and linkage (reusing the audit/change-history pattern already in the ERP).

**Suggested additions (recommended):**

- **Preventive Maintenance (PM) schedules** — recurring PM plans per machine (interval by days or by impression count, using `Machine.standard_impressions_per_hour`) that auto-generate due `MaintenanceRecord`s. Turns the module from reactive log into planned maintenance.
- **Downtime & MTBF/MTTR KPI dashboard** — availability %, mean time between failures, mean time to repair, top failing machines, spares cost per machine.
- **Technician assignment & labour hours** — who worked, hours spent (feeds MTTR and cost).
- **Cost roll-up per record/machine** — spares cost (from linked IR unit prices) + service cost + labour.
- **Attachments** — fault photos, vendor quotes, service reports.
- **Notifications** — reuse `core.notifications` to alert maintenance/manager roles on new breakdown, PM due, IR received (part arrived → schedule work).

---

## 4. Data model (`maintenance/models.py`)

### 4.1 `MaintenanceRecord` (core)
| Field | Type | Notes |
|---|---|---|
| `record_no` | CharField unique, blank until saved | `MNT-{YEAR}-{seq}`, generated like the IR-ID helper |
| `machine` | FK → `core.Machine` PROTECT | |
| `reported_date` | DateField default today | |
| `reported_by` | FK → User | |
| `fault_types` | **M2M → FaultCategory** | **multi-select** — e.g. a fault can be Mechanical *and* Electrical together (§4.1a) |
| `maintenance_type` | CharField choices | Preventive/Corrective/Breakdown/Predictive |
| `execution_type` | CharField choices | `IN_HOUSE` / `OUTSOURCE` |
| `priority` | CharField choices | Low/Medium/Major/Critical |
| `fault_description` | TextField | |
| `proposed_solution` | TextField blank | |
| `spare_parts_needed` | Bool | |
| `repair_needed` | Bool | |
| `repair_details` | TextField blank | |
| `status` | CharField choices | see §4.5 |
| `assigned_to` | FK → User null | technician/team lead |
| `work_start_at` | DateTimeField null | |
| `work_end_at` | DateTimeField null | |
| `remarks` | TextField blank | |
| `created_at`/`updated_at` | auto | |

### 4.1a `FaultCategory` (lookup, multi-select)
Configurable so categories aren't hard-coded (same pattern as `ItemRequestType`). Seeded via data migration with: Mechanical, Electrical, Electronic, Pneumatic, Hydraulic, Other. A record links to **one or more** via `fault_types` M2M — so "Mechanical + Electrical" on the same fault is fully supported. Form uses a multi-select / checkbox widget; list filters allow filtering by any category.
| Field | Type | Notes |
|---|---|---|
| `name` | CharField unique | |
| `is_active` | Bool | |

### 4.2 `MaintenanceSparePart`
Line items describing needed spares before/instead of a full IR (mirrors the sheet's "Spare Parts Detail"). Each line can be linked to an `ItemRequest` once demand is raised.
| Field | Type | Notes |
|---|---|---|
| `record` | FK → MaintenanceRecord | |
| `description` | CharField | e.g. "Kompac Teflon seal for GTO 52" |
| `quantity` | Decimal | |
| `uom` | CharField | |
| `existing_sku` | FK → RawMaterialSku null | link if already in master |
| `item_request` | FK → supply_chain.ItemRequest null | set when demand raised |

### 4.3 `MaintenanceServiceJob`
For outsourced repair work → one service Item Request.
| Field | Type | Notes |
|---|---|---|
| `record` | FK → MaintenanceRecord | |
| `vendor` | CharField blank | |
| `scope` | TextField | what the vendor will do |
| `item_request` | FK → supply_chain.ItemRequest null | the `SRV` request |
| `sent_out_date` / `returned_date` | DateField null | asset-out/return tracking |

### 4.4 `MachineDowntime`
| Field | Type | Notes |
|---|---|---|
| `machine` | FK → core.Machine | |
| `record` | FK → MaintenanceRecord null | breakdown that caused it |
| `start_at` / `end_at` | DateTimeField | end null = still down |
| `reason` | CharField choices | Breakdown/Planned PM/Awaiting Part/Awaiting Vendor/Other |
| `scheduled_minutes_lost` | Int (computed) | netted vs `MachineWorkSchedule` (§5) |

### 4.5 `MaintenanceActivityLog` (immutable)
`record`, `actor`, `action`, `from_status`, `to_status`, `note`, `created_at`. Written on every transition.

### 4.6 `PreventiveMaintenancePlan` (phase 2)
`machine`, `title`, `interval_type` (DAYS/IMPRESSIONS), `interval_value`, `last_done_at`, `next_due_at`/`next_due_impressions`, `is_active`. A management command / scheduled task generates due `MaintenanceRecord`s.

### 4.7 Status lifecycle
`REPORTED → DIAGNOSED → AWAITING_PARTS → AWAITING_VENDOR → IN_PROGRESS → COMPLETED → VERIFIED → CLOSED` (plus `CANCELLED`). Only some transitions valid; each logged.

---

## 5. Downtime calculation

`MachineDowntime` stores raw start/end. Effective downtime = overlap of `[start_at, end_at]` with the machine's **scheduled working** windows from `core.MachineWorkSchedule` + `ShiftConfig`, so a breakdown over a Friday the machine is off doesn't inflate downtime. Availability % = 1 − (scheduled_minutes_lost / scheduled_available_minutes) over a period. This reuses existing shift/schedule infrastructure rather than reinventing a calendar.

---

## 6. Integration with Item Request / Services

1. **Seed a `SRV` (Service) request type** if not present, alongside the existing `MNT`/`SPR` types (data migration in `supply_chain`).
2. From a maintenance record, a **"Raise Demand"** action creates one or more `ItemRequest`s:
   - Spare parts → type `SPR`/`MNT`, `machine` pre-filled from the record, `item_title`/`specifications` pre-filled from the spare line, `raised_by` = current user, `department` = Maintenance.
   - Outsource repair → type `SRV`, description = service scope.
   The created IR enters the normal manager/supply-chain approval + procurement flow untouched.
3. **Back-link:** `MaintenanceSparePart.item_request` / `MaintenanceServiceJob.item_request` FK. The maintenance detail page shows each linked IR's live status and procurement timeline (Indent/PR/PO/Received) — read from `ItemProcurementTimeline`.
4. **Two-way status hint:** when a linked IR reaches `RECEIVED`, signal/notify maintenance that parts arrived → prompt to move record to `IN_PROGRESS`. When all linked IRs are `CLOSED` and work_end set → record can be `COMPLETED`.
5. **Two small additions to supply_chain** (non-breaking): a nullable `source_module`/`source_ref` on `ItemRequest` (or a reverse relation is enough via the FKs above — preferred, so *no* change to `ItemRequest` is strictly required). Chosen approach: keep FKs on the maintenance side only → **zero migration on `supply_chain` except seeding the `SRV` type.**

---

## 7. Views / URLs (`maintenance/urls.py`, namespace `maintenance`)

- `` — dashboard (open records, machines currently down, PM due, KPIs)
- `records/` — list + filters (machine, status, type, priority, date range)
- `records/new/` — create fault/maintenance record
- `records/<id>/` — detail: activity log, spare lines, service jobs, linked IR statuses, downtime, actions
- `records/<id>/raise-demand/` — spawn Item Request(s)
- `records/<id>/transition/` — status change (logged)
- `downtime/` — downtime log + availability report
- `pm-plans/` — preventive plans (phase 2)
- `reports/` — MTBF/MTTR/availability/cost dashboard

Follow existing conventions: `login_required`, role checks via the app's `decorators.py` pattern, templates under `maintenance/templates/maintenance/`, add a nav entry in the base template.

---

## 8. Permissions / roles (confirmed)
- **Maintenance Engineer** — logs fault/maintenance records, adds spare lines & service jobs, raises demand, updates work progress and downtime.
- **Manager** — approves the raised demand / Item Requests (existing manager approval stage in the IR workflow) and can verify/close records.
- `admin` — full access; report views open to all maintenance-access roles.

Add a `maintenance_engineer` role group; reuse the existing role-management screen and IR approval stages (no new approval engine needed).

---

## 9. Build phases

**Phase 1 — Core log + downtime + IR linkage (replaces the sheet)**
Models 4.1–4.5, record CRUD, raise-demand → Item Request, back-linked IR status display, downtime capture + basic availability, activity log, list/detail/dashboard, `SRV` type seed.

**Phase 2 — Planning & analytics**
`PreventiveMaintenancePlan` + auto-generation (scheduled task), MTBF/MTTR/availability/cost dashboard, notifications, technician labour hours & cost roll-up.

**Phase 3 — Polish**
Attachments, Excel export (reuse `excel_io` patterns), saved filters, mobile-friendly quick-report form for the shop floor. (No sheet-history import — starting fresh at go-live.)

---

## 10. Decisions — confirmed

1. **Roles:** Maintenance Engineer logs the fault; Manager approves the raised demand (§8).
2. **Downtime:** auto-starts when a Breakdown record is created, and remains editable by the engineer.
3. **Fault type is multi-select** (`FaultCategory` M2M, §4.1a) — Mechanical + Electrical etc. can be selected together. Terminology matches the shop-floor list; adjust via the lookup table anytime.
4. **Cost tracking → Phase 2.**
5. **Start fresh** at go-live — no Google Sheet history migration. (Excel export still available in Phase 3.)
