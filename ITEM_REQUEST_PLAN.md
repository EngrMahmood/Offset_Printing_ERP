# Item Request Module — Implementation Plan

**Location:** extends the existing `supply_chain` app
**Goal:** Let any department raise a request for an item they need (production, maintenance, etc.), route it through manager / supply-chain approval, and track the full procurement timeline (code opening → indent/PR → PO → receipt) against supply-chain KPIs.

---

## 1. Scope & workflow overview

```
Requester            Approver (Mgr / SC)         Supply Chain
---------            -------------------         ------------
Fill request  ─────► Review                ──┐
form + log         │  • Approve            │
                   │  • Reject             │
                   │  • Need Amend         │   (on Approve)
                   │  • Need More Info     │        │
                   └─◄ (back to requester) │        ▼
                                           └─► IR-ID generated
                                                     │
                                                     ▼
                                        Procurement timeline entry:
                                        item code opening → PR/indent →
                                        PO → received date → unit price
                                                     │
                                                     ▼
                                            Feeds SC KPI dashboard
```

Every state change is written to an immutable log (reuse the pattern of the `audit` app / existing change-history views).

---

## 2. Data model (new models in `supply_chain/models.py`)

### 2.1 `ItemRequestType`
Configurable lookup so types aren't hard-coded.

| Field | Type | Notes |
|---|---|---|
| `name` | CharField unique | Raw Material, Consumable, Maintenance Item, Service, Spare Part |
| `code` | CharField(4) | short prefix for IR-ID, e.g. `RM`, `CON`, `MNT`, `SPR` |
| `is_active` | BooleanField | |

Seed via a data migration with the five types above.

### 2.2 `ItemRequest`
The core request record. Fields mirror the existing **UTOPIA Printing & Packaging Request Form** (see attached paper form). `*` = mandatory on the form.

| Field | Type | Form label | Notes |
|---|---|---|---|
| `request_no` | CharField unique, blank until approved | (Approved By area) | the IR-ID, e.g. `IR-MNT-2026-0042` |
| `request_type` | FK → ItemRequestType | Item Request Type | **dropdown + "＋ add" button** (§4.1) |
| `request_date` | DateField, default today | Date | |
| `item_title` | CharField | Item Title | e.g. "Silver Spring Heat Remove Pipe" |
| `machine` | FK → core.Machine, null, blank | Machine Name / Model | **optional** — fill only when the item is machine-specific; dropdown from existing machine master (§2.7) |
| `machine_other` | CharField, blank | Machine Name / Model | fallback free text if the machine isn't in the master yet |
| `uom` | CharField * | Unit of Measure (UOM) * | |
| `specifications` | TextField * | Specifications / Technical Details * | |
| `description` | TextField | Description / Purpose of Use | |
| `dimensions` | CharField | Dimensions (if applicable) | e.g. `8" inch` |
| `local_import` | CharField choices | Local / Import | Local / Import |
| `part_number` | CharField | Part Number | |
| `required_quantity` | Decimal | Required Quantity | e.g. "50 fit" |
| `department` | FK → Department | — | fixed choice list + "＋ add" (§2.6) |
| `existing_sku` | FK → RawMaterialSku, null | — | link if the item already exists |
| `estimated_unit_price` | Decimal, null | — | requester's estimate |
| `attachment` | FileField, null | — | photo of paper form / quote / drawing |
| `status` | CharField choices | — | see §2.4 |
| `raised_by` | FK → User | Request by | |
| `created_at` / `updated_at` | DateTime | — | auto |

### 2.7 Machine field — reuse `core.Machine`
The "Machine Name / Model" field is **optional**, shown only when the requested item is specific to a machine (e.g. a spare part for GTO 1B UV). It's a dropdown populated from the **existing `core.Machine` master** (no new machine model). If the needed machine isn't in the master, `machine_other` free text is captured and can be promoted to a real machine record later. On the form, hide/disable the machine dropdown unless the request type is maintenance/spare-part, or leave it always-available but non-mandatory.

### 2.6 `Department`
Fixed choice list, editable via a "＋ add" button on the form (§4.1).

| Field | Type | Notes |
|---|---|---|
| `name` | CharField unique | Production, Maintenance, Prepress, Dispatch, Store, etc. |
| `is_active` | BooleanField | |

### 2.3 `ItemRequestApproval` (log of every decision/action)
| Field | Type | Notes |
|---|---|---|
| `request` | FK → ItemRequest, related_name `approvals` | |
| `actor` | FK → User | |
| `action` | CharField choices | APPROVE / REJECT / AMEND / MORE_INFO / RESUBMIT |
| `stage` | CharField | MANAGER / SUPPLY_CHAIN |
| `comment` | TextField | required for reject/amend/more-info |
| `created_at` | DateTime auto | immutable — never edited |

### 2.4 Status choices on `ItemRequest`
Manager approval is **always required**, then supply chain:
`SUBMITTED → MGR_REVIEW → SC_REVIEW → APPROVED → IN_PROCUREMENT → RECEIVED → CLOSED`
plus branches: `REJECTED`, `NEEDS_AMENDMENT`, `NEEDS_INFO` (bounce back to requester).

### 2.5 `ItemProcurementTimeline` (supply-chain data entry, 1:1 with an approved request)
This is what SC fills after approval and what drives KPIs.

| Field | Type | KPI use |
|---|---|---|
| `request` | OneToOne → ItemRequest | |
| `item_code` | CharField | code opening reference |
| `code_opened_date` | DateField, null | KPI: code-opening lead time |
| `indent_pr_no` | CharField, null | |
| `pr_date` | DateField, null | KPI: PR raise time |
| `po_no` | CharField, null | |
| `po_date` | DateField, null | KPI: PR→PO cycle |
| `supplier` | CharField / FK, null | |
| `received_date` | DateField, null | KPI: PO→receipt lead time |
| `unit_price` | Decimal, null | actual vs estimate variance |
| `received_qty` | Decimal, null | |
| `remarks` | TextField | |
| `updated_by` / `updated_at` | | |

---

## 3. Derived KPIs (add to `supply_chain/kpis.py`)

Computed from `ItemProcurementTimeline` date fields:

- **Approval cycle time** = `APPROVED` timestamp − `SUBMITTED` timestamp (from approval log).
- **Code-opening lead time** = `code_opened_date − approved_date`.
- **PR turnaround** = `pr_date − code_opened_date`.
- **PR→PO cycle** = `po_date − pr_date`.
- **PO→Receipt lead time** = `received_date − po_date`.
- **Total fulfilment time** = `received_date − submitted_date`.
- **On-time rate** = % received on/before `required_by_date`.
- **Price variance** = `unit_price − estimated_unit_price`.
- **Open vs closed request counts**, aging buckets, per-type and per-department breakdowns.

Aggregate these into a KPI dashboard card set reusing existing SC dashboard rendering.

---

## 4. Views & URLs (`supply_chain/views.py`, `supply_chain/urls.py`)

**Access model.** Roles are string profile roles (`admin`, `manager`, `supply_chain`, …) resolved in `core/navigation.py`. Add a new set:
`ITEM_REQUEST_NAV_ROLES = {'admin', 'manager', 'supply_chain', 'planner', 'production_manager', 'production', ...}` — i.e. anyone who may need to raise a request. Expose `can_access_item_request`. Manager-stage actions check `role in {'admin','manager'}`; SC-stage actions reuse `supply_chain_required` / `role in {'admin','supply_chain'}`.

| Route | View | Access |
|---|---|---|
| `item-requests/new/` | create form | anyone with `can_access_item_request` |
| `item-requests/` | list (filter by status/type/dept, my-requests toggle) | all with access |
| `item-requests/<id>/` | detail + full approval log timeline | requester + approvers |
| `item-requests/<id>/review/` | approve / reject / amend / more-info — manager stage then SC stage | `manager` (stage 1), `supply_chain` (stage 2) |
| `item-requests/<id>/resubmit/` | requester edits & resubmits after amend/info | requester |
| `item-requests/<id>/procurement/` | SC enters timeline data | `supply_chain_required` |
| `item-requests/kpis/` | KPI dashboard | anyone with `can_access_item_request` |

### 4.1 Inline "＋ add" for Request Type & Department
On the request form, both `request_type` and `department` render as a dropdown with a small **＋** button beside it. Clicking opens a modal (AJAX POST to `item-requests/type/add/` and `item-requests/department/add/`) that creates the row and appends it to the select without a page reload. Gate creation to `manager`/`admin`/`supply_chain` if you want to keep the lists clean, or allow all — your call.

---

## 5. IR-ID generation

Generate `request_no` only on final approval, in one place (a `generate_request_no()` method or `services.py` helper), format:
`IR-{type.code}-{YYYY}-{zero-padded sequence}` e.g. `IR-MNT-2026-0042`.
Use a per-year, per-type counter inside a DB transaction (`select_for_update`) to avoid race duplicates.

---

## 6. Templates
Under `supply_chain/templates/supply_chain/item_request/`: `form.html`, `list.html`, `detail.html` (with approval-log timeline), `review_modal.html`, `procurement_form.html`, `kpi_dashboard.html`. Match existing SC template styling/theme.

## 7. Admin & signals
- Register all four models in `supply_chain/admin.py`.
- Signal (or view logic) to auto-create the `ItemProcurementTimeline` row and generate the IR-ID when status flips to `APPROVED`.

## 8. Migration & seeding
1. `makemigrations supply_chain` for the four models.
2. Data migration seeding `ItemRequestType` rows.
3. `migrate`.

## 9. Navigation
Add an **"Item Requests"** entry to the **top nav bar** using the global theme (same as other modules), gated by `can_access_item_request`. Add `ITEM_REQUEST_NAV_ROLES` + `can_access_item_request` in `core/navigation.py` and the link in the base/nav template.

---

## 10. Build phases

**Phase 1 — Core loop (MVP):** models, migrations, request form, list, detail, submit → manager/SC review (approve/reject/amend/more-info), approval log, IR-ID on approval.

**Phase 2 — Procurement timeline:** SC data-entry form, timeline fields, auto-created on approval, detail-page timeline display.

**Phase 3 — KPIs:** KPI functions + dashboard, aging/on-time/variance, per-type & per-department filters.

**Phase 4 — Polish:** attachments, Excel export (reuse `excel_io.py` pattern), notifications, admin, nav.

---

## 11. Room for improvement / suggestions

1. **Spend-per-machine reporting** — since the optional machine field links to `core.Machine`, add a KPI/report of request spend grouped by machine, useful for maintenance-cost tracking per press.
2. **Link to existing SKU early** — if the requested item already has a `RawMaterialSku`, skip code-opening and pre-fill price/lead time. Reduces duplicate item codes.
3. **Duplicate-request detection** — warn if an open request already exists for the same item/department (fuzzy match on `item_name`, reuse the normalize helpers already in `models.py`).
4. **Budget / cost centre field** — capture department budget code up front for finance reporting.
5. **Notifications** — email or in-app alert to approver on submit and to requester on decision. Ties into a `tasks` app entry if you want a to-do queue.
6. **SLA timers** — flag requests sitting in a stage beyond N days; feeds directly into SC KPI accountability.
7. **Audit integration** — route state changes through the existing `audit` app rather than a bespoke log, so history is consistent with the rest of the ERP.
8. **Attachments & quotes** — allow multiple supplier quotes on the procurement side to justify PO price (supports price-variance KPI).
9. **Reopen / partial receipt** — support partial deliveries (received_qty < quantity) and keep the request open until fully received.
10. **Digitize the paper form 1:1** — since you're replacing a printed form, add a "print / export PDF" view of the request that mirrors the UTOPIA layout, so a hard copy can still travel with the item if needed.

---

## 12. Decisions locked (from review)
- **Manager approval always required**, then supply chain. ✅
- **Item Request Type** = dropdown with inline "＋ add". ✅
- **Department** = fixed choice list (model) with inline "＋ add". ✅
- **Roles** = existing `manager` and `supply_chain` profile roles. ✅
- **KPIs** visible to anyone with item-request access. ✅
- **Nav** = top nav bar, global theme. ✅
- **Machine** = optional FK to existing `core.Machine`, filled only when item is machine-specific. ✅

Ready for Phase 1 build.
