# Printing Plates ERP Mapping

## 1. Purpose

This document captures the AppSheet `Printing Plates` app structure and maps it to the existing Django ERP entities in `Offset_Printing_ERP`. It is intended as a design reference for a new `Printing Plates` module, not as an implementation.

## 2. AppSheet source and analysis

- App name: `Printing Plates_Update`
- AppSheet data source: Google Sheet `Printing Plates Data`
- Core AppSheet tables extracted:
  - `Plates`
  - `Plate Request`
  - `Plate Received`
- Additional AppSheet process state tables detected:
  - `Process for Send Notification Request Process Table`
  - `Process for Send Notification Received Process Table`

## 3. AppSheet table summaries

### 3.1 Plates

Source: `DocId=1OPRrsZhG3-brAALFnVm731raKhDbfghpLVyYRMpRXgg`

Columns:
- `_RowNumber` (AppSheet key, read-only)
- `S#` (Number, primary set identifier)
- `Layout Date` (Date)
- `Old Set No.` (Number, read-only)
- `Date` (Text)
- `AWC No` (Text)
- `Sku` (Text)
- `Job Name` (Name)
- `Status` (Text)
- `No Of Colors` (Number)
- `Machine` (Enum)
- `Set No` (Number)
- `Source` (Enum)
- `Chalan Sign` (Yes/No)
- `Column_14` (Show)
- `Challan` (Text)
- `Department` (Text)
- `Material` (Text)
- `Column_18` (Show)
- `Box` (Text)

Views:
- `Designer Plates Data`
- `Plates_Detail`
- `Plates_Form`

Actions:
- `Delete`
- `Edit`
- `Add`

### 3.2 Plate Request

Source: `DocId=1OPRrsZhG3-brAALFnVm731raKhDbfghpLVyYRMpRXgg`

Columns:
- `_RowNumber`
- `JC#` (Text)
- `Set #` (Text)
- `New Set #` (Text)
- `Vendor` (Enum)
- `AWC #` (Text)
- `Department` (Enum)
- `SKU` (Text)
- `Machine Name` (Enum)
- `Plate Quantity` (Number)
- `Status` (Enum)
- `Plate Color` (EnumList)
- `Impression` (Text)
- `Remarks` (Text)
- `Requested By` (Enum)
- `Request Date` (DateTime)
- `Progress` (Text)
- `Request By` (Text)
- `Sent By` (Text)
- `Sent At` (DateTime)
- `Received By` (Text)
- `Received At` (DateTime)
- `Image` (Image)
- `Link` (Url)

Views:
- `Production Plate Request Form`
- `Production Plates`
- `Plate Request_Detail`
- `Plate Request_Form`

Actions:
- `Delete`
- `Edit`
- `Add`
- `Marked as Sent`
- `Marked as Received`
- `Received`
- `Back Screen`
- `update set`
- `Sent`
- `Open Url (Link)`

### 3.3 Plate Received

Source: `DocId=1OPRrsZhG3-brAALFnVm731raKhDbfghpLVyYRMpRXgg`

Columns:
- `_RowNumber`
- `Set #` (Text, primary key)
- `AWC#` (Text)
- `Receiver Name` (Enum)
- `Designer` (Enum)
- `Date` (Date)

Views:
- `Old Plate Received`
- `Plate Received_Detail`
- `Plate Received_Form`

Actions:
- `Delete`
- `Edit`
- `Add`

## 4. Key AppSheet workflow and semantics

The AppSheet app is using a request/receipt pattern:

- `Plate Request` captures plate job requests, including job card references, machine, quantity, vendor, status, and send/receive metadata.
- Actions such as `Marked as Sent`, `Marked as Received`, and `Received` are likely AppSheet status transitions or automation triggers.
- `Plate Received` appears to be a separate receipt register for completed plates, possibly for designer or customer acceptance.
- `Plates` is a production register summarizing plate metadata and status for the current print batch.
- The two native process tables are AppSheet-generated workflow state tables and are not source model data.

## 5. Proposed ERP mapping and reuse

### 5.1 Existing ERP entities likely reusable

- `Machine` / `core.models.Machine`
  - `Machine Name` in AppSheet should map to ERP machine master data.
- `Department` / `core.models.Department`
  - AppSheet `Department` values should reuse ERP department definitions.
- `Material` / `core.models.Material`
  - AppSheet material references may map to ERP material or stock master data.
- `JobCard` / `core.models.JobCard`
  - `JC#`, `SKU`, `Job Name`, `Impression`, and machine/set details should link to existing job card entities.
- `Operator` / `UserProfile` or employee identity
  - `Requested By`, `Sent By`, `Received By`, `Designer` should map to ERP users or operators.
- `Vendor`
  - AppSheet `Vendor` may map to existing vendor/supplier entity if one exists or can be introduced as a simple reference.
- `Status` and `Plate Color`
  - These are best modeled as enumerations or lookup tables in ERP.

### 5.2 New module/entities required

A dedicated ERP module should be introduced for printing plates with the following entities:

- `PlateRequest`
  - Fields: `job_card`, `set_number`, `new_set_number`, `vendor`, `awc_number`, `department`, `sku`, `machine`, `quantity`, `status`, `plate_color`, `impression`, `remarks`, `requested_by`, `request_date`, `progress`, `sent_by`, `sent_at`, `received_by`, `received_at`, `image`, `link`
  - Relationships:
    - `job_card` -> `JobCard`
    - `machine` -> `Machine`
    - `department` -> `Department`
    - `requested_by`, `sent_by`, `received_by`, `designer` -> `UserProfile` or `Employee`
- `PlateReceipt`
  - Fields: `set_number`, `awc_number`, `receiver`, `designer`, `receipt_date`, `plate_request` (optional back-link)
- `PlateRegister` or `PlateRecord`
  - Fields: `set_id`, `layout_date`, `old_set_no`, `date`, `awc_no`, `sku`, `job_name`, `status`, `color_count`, `machine`, `set_no`, `source`, `chalan_sign`, `challan`, `department`, `material`, `box`

### 5.3 Reuse matrix

| AppSheet field | ERP reuse candidate | Notes |
|---|---|---|
| `JC#` | `JobCard.job_card_no` or `JobCard.job_card_id` | Link to job card; may require normalization and lookup. |
| `Set #`, `New Set #` | Production set / job card set number | Can often be modeled as part of `JobCard` or a related set record. |
| `Vendor` | Existing supplier/vendor model | If not present, add a simple vendor lookup. |
| `AWC #` | Custom field on request/plate record | Could be stored as text on `PlateRequest` and `PlateRecord`. |
| `Department` | `Department` model | Reuse existing department master data. |
| `SKU` | `PlanningJob` or item code | Could reuse SKU from job card or planning models. |
| `Machine Name` | `Machine` model | Core machine master reuse. |
| `Plate Quantity` | New request quantity field | Numeric field on the new request model. |
| `Status` | New plate status lookup | Use enum choices or dedicated workflow statuses. |
| `Plate Color` | New lookup/table or enum list | Could be normalized to color options. |
| `Impression` | `JobCard.impression` or text field | If job card contains impressions, reuse; else add free-text. |
| `Requested By`, `Sent By`, `Received By`, `Designer` | `UserProfile` / employee | Reuse ERP user identity and permission model. |
| `Request Date`, `Sent At`, `Received At` | Timestamp fields | Standard datetime fields on request/receipt models. |
| `Image`, `Link` | Attachment/URL fields | Reuse existing file/link field patterns in ERP if available. |

## 6. Proposed module structure

### 6.1 Menu and UI placement

- Add a new top-level menu or submenu under `Production` / `Printing`:
  - `Printing Plates`
    - `Plate Requests`
    - `Plate Receipts`
    - `Plate Register`
- Provide list/filter views for:
  - open plate requests
  - sent/received plate requests
  - pending plate receipts
  - historical plates

### 6.2 Core screens/forms

- `Plate Request Form`
  - Capture request details and job card linkage.
  - Allow machine and quantity selection.
  - Set initial status and optional vendor.
- `Plate Request Dashboard`
  - Show summary grouped by status / machine / department.
- `Plate Receipt Form`
  - Record receipt with receiver/designer and date.
- `Plate Register Form`
  - Manage plate details for stock and production tracking.

### 6.3 Workflow actions

- `Submit Request`
- `Mark as Sent`
- `Mark as Received`
- `Receive Plate`
- `Update Set`
- `Open External Link`

## 7. Security and permissions

### 7.1 Permission model

- Use existing ERP roles and permission system.
- Suggested permissions:
  - `can_view_plate_requests`
  - `can_create_plate_requests`
  - `can_edit_plate_requests`
  - `can_mark_plate_sent`
  - `can_receive_plates`
  - `can_manage_plate_register`

### 7.2 Role mapping

- Production staff and supervisors: create/edit requests, mark sent.
- Designers and QC/inspection staff: receive plates and confirm completion.
- Admins/Managers: full access to plate register and status dashboards.

## 8. Integration points

### 8.1 Existing ERP integrations

- Link plate requests to `JobCard` / `PlanningJob` for a single source of truth.
- Use existing `Machine` master and `Department` master to avoid duplicate enums.
- Reuse user identity `Operator` / `UserProfile` for request ownership and receipt confirmation.
- If vendor masters exist, reuse them rather than free-text vendor names.

### 8.2 Notifications and automation

- AppSheet has process tables for notification state; in ERP, this should be implemented using Django signals or workflow services.
- Potential event triggers:
  - request created -> notify production/vendor
  - request sent -> notify plate shop
  - plate received -> notify design / QC

## 9. Gaps and risks

### 9.1 Gaps in the AppSheet model

- AppSheet uses loose text and enums instead of normalized foreign keys.
- `Job Name`, `SKU`, `AWC #`, `Status`, and `Plate Color` are not clearly mapped to existing ERP domains.
- `Vendor` is an enum with unknown value set.
- `Image` and `Link` fields need a storage strategy in Django.
- The AppSheet workflow tables are internal/automation artifacts and should not be modeled directly as business objects.

### 9.2 Risks for ERP implementation

- Data migration from AppSheet/Google Sheets may require cleanup and normalization.
- Unknown AppSheet action logic means some transitions may need manual reverse engineering.
- Existing ERP data models may not support direct one-to-one mapping for plate set history, requiring new intermediate tables.
- User role semantics in AppSheet are implicit; ERP permissions must be explicitly designed.

## 10. Recommended next steps

1. Validate the AppSheet field value sets for `Status`, `Plate Color`, `Vendor`, `Department`, and `Machine Name`.
2. Confirm whether `JC#` and `SKU` should link to existing job cards or be stored as standalone identifiers.
3. Define the new Django models for `PlateRequest`, `PlateReceipt`, and `PlateRecord`.
4. Design a permissions matrix and workflow state machine for plate request lifecycle.
5. Plan a migration path from the AppSheet source data into ERP tables.

---

> Note: This document is based on the AppSheet editor metadata visible in the current browser session and the existing Django ERP repo structure. Actual field value sets, enums, and AppSheet automation rules may require further extraction from the AppSheet editor or source spreadsheet.
