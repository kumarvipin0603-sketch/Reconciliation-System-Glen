<<<<<<< HEAD
# Reconciliation Tower v11 — Persistent, Read-Only Main Reconciliation + Auditable Team Workflow

This build uses the current Amazon and Flipkart reconciliation Python logic supplied with the source workbooks.
Telegram bot credentials are disabled in this Streamlit package.

## Persistent data
Operational records are stored outside the website folder at:

    C:\Users\<WindowsUser>\ReconciliationTowerData\reconciliation_tower.db

The records survive browser refresh, Streamlit restart, closing localhost, Windows shutdown/restart and replacing the website folder with a future version.

## Portal source upload
- Upload Amazon and Flipkart together in one selection, or one at a time.
- Each workbook is auto-detected by sheet structure.
- Re-uploading one marketplace updates that marketplace while the other remains available.
- Source Upload History and Activity Audit Trail are retained.

## Main dashboards (read-only)
1. E-Com Process Reconciliation Dashboard
2. Settlement Process Reconciliation Dashboard

Main reconciliation cannot be edited directly from the website.
All team changes must flow through Pending Task Download -> Team Working Sheet -> Team Update Upload.

## Source-level enrichment
Amazon source:
- Orders: shipment/order qty/value, SKU, branch
- Reverse: return created date, return delivery date, return status/type
- Payments: released net pay, refund/reversal, fees/deductions, settlement reference/date
- Sales: invoice/CN number, date, qty and price
- Replacement_Map: replacement -> original order roll-up
- Reimbursement: amount and reason

Flipkart source:
- ERP Sales Register: invoice/CN number, date, qty, value, branch and item
- Payments: bank settlement, refund, marketplace fee/tax deductions, payment date and NEFT reference
- Returns: approval/completion dates and status
- All Orders: item/status/delivery details
- Sales Report: source order activity used by the supplied reconciliation engine

## Pending Task workflow
Any non-reconciled business remark creates a persistent task. The task records:
- Task Created Date
- Task Status (Pending / Working / Completed)
- Team Remarks
- Task Completed Date
- Aging Days
- Branch/task owner assignment

The team can view the main reconciliation and Pending Task Dashboard but cannot edit the main reconciliation directly.
They download the filtered Pending Task workbook, update the working fields, and upload the same workbook under Team Update Upload.
The main reconciliation immediately reflects the uploaded team status and remarks.

## MIR workflow
Separate persistent workspaces remain for MIR 1600 / 5400 / 5600 / 7800.
MIR/TEI/SR uploads are recorded in the database and reflected back into the E-Com Process Reconciliation Dashboard.

## Filters
Dashboards support:
Market Place | Branch | Pending Task | Task Status | Individual Order ID | Task Assign To

## Audit Trail
The Activity Audit Trail records source uploads, task creation, task-sheet uploads, MIR uploads and owner-master changes.

## Admin vs Team permissions
- On first use, create an Admin PIN in the left sidebar.
- Only an unlocked Admin can upload/replace Amazon or Flipkart source workbooks and change Branch Owner / Microsoft 365 settings.
- Team users can view the read-only reconciliation dashboards, download pending-task sheets, upload completed working sheets, and update the relevant MIR branch sheet.
- The main reconciliation cannot be manually edited from the website.


## v11.1 Header Standardization Fix
Standard headers throughout:
- Net Pay Received
- Reimbursement
- Short Payment Received
- Return Delivery Date

Legacy team upload files are normalized automatically. Dashboard summary
metrics and exports also tolerate missing optional columns without KeyError.


## v11.2 Dashboard corrections
Main dashboard filters:
Market Place | Branch | Pending Task | Task Status | Individual Order ID

Pending Task shows reconciliation issue types such as CN Pending, Billing Pending,
Billing & Payment Pending, Payment Pending, Short Payment Received and Extra Billing.

Task Status is workflow status only: Pending / Working / Completed.
Task Assign To filter was removed from the main dashboard.
Return Value KPI was removed.
Net Pay Received now uses the reconciliation engine's received/settlement amount.
Pending/Working and Completed KPI counts update from the persistent task workflow.


## v11.3 Source + MIR field repair
- Fixes duplicate task/MIR overlay that caused MIR No/Date, TEI No/Date, Team Remarks, Task Completed Date and Pending Task Created Date to appear blank.
- Amazon replacement IDs and replacement-order source activity are rolled to original orders.
- Amazon payment/refund activity is rolled across original + replacement IDs.
- Flipkart ERP Sale/Return is used to classify invoice vs CN.
- Explicit headers: Sale (Invoice No), Sale Date (Invoice Date), Return (CN No), Return Date (CN Date), Return Price (CN Price).
- Return Delivery Date comes from Amazon Reverse / Flipkart Returns completion data.
- MIR branch uploads automatically overlay the main reconciliation, with unique-order fallback if branch formatting differs.
- Re-upload Amazon and Flipkart source workbooks once after upgrade to refresh source-derived blank fields. Task/MIR history is not cleared.


## v11.4 Exact Net Pay Received source
Net Pay Received is now taken from the exact source headers:
- Amazon Payments -> `Rece Amount`
- Flipkart Payments -> `Bank Settlement Value (Rs.)`

Values are grouped by Amazon Order ID / Flipkart PO Number.
Existing persistent rows must be refreshed by re-uploading the current
Amazon and Flipkart source workbook once.


## v11.5 Display Header Change
The dashboard/output header `Net Pay Received` has been renamed to `Received`.

The underlying source mapping is unchanged:
- Amazon Payments -> `Rece Amount`
- Flipkart Payments -> `Bank Settlement Value (Rs.)`

Only the website/downloadable reconciliation display header was changed.


## v11.6 Current FY Source Mapping

Verified against the current uploaded Amazon and Flipkart FY source workbooks.

Amazon:
- Payment Received <- Payments / `Payment Received`
- Replacement Order Id <- Replacement_Map / `Replacement Order Id`
- Return Delivery Date <- Reverse / `Return Delivery Date`
- Sale Invoice/CN <- Sales / Invoice No, Invoice Date, Quantity, Item Price,
  Document Type

Flipkart:
- Payment Received <- Flipkart Payments / `Payment Received`
- Return Delivery Date <- Flipkart Returns / `return_completion_date`
- Sale Invoice/CN <- ERP Sales Register / Invoice No, Invoice Date, Quantity,
  Line Amount/Gross Amount, Document Type

Older payment aliases remain supported for backward compatibility.


## v11.7 Flipkart Return Date Correction
Flipkart `Return Delivery Date` in the reconciliation is now sourced from:
`Flipkart Returns` -> `Return Approval Date`

The earlier return-completion mapping is no longer used for this field.
All other v11.6 source mappings remain unchanged.


## v11.8 Source Join / Task Filter Fix

Source-of-truth hierarchy:

### Amazon
- Replacement Order Id: reconciliation engine / Replacement_Map
- Return Delivery Date: Reverse -> Return Delivery Date
- Payment Received: Amazon reconciliation engine `Rece Amount`
  (already includes original + replacement IDs)
- Sale Invoice No / Date: Sales -> SALE rows
- Return/CN No / Date / Price: Sales -> RETURN rows
- Source invoice/CN and return details on replacement IDs are rolled back
  to the original Amazon order.

### Flipkart
- Return Delivery Date: Flipkart Returns -> Return Approval Date
- Payment Received: Flipkart reconciliation engine `Received Amount`
  sourced from marketplace Payment Received
- Sale/CN No, Date, Qty, Price: ERP Sales Register

### Task workflow
Missing task records from older Tower versions are automatically backfilled
from saved Pending Remarks. Existing task history is never overwritten.
Pending Task filter always contains the defined business task types and also
adds any task types found in the saved reconciliation.


## v11.9 Engine-Level Source Field Fix

Source fields are now produced directly by the reconciliation engines.

Amazon engine:
- Replacement Order Id
- Return Delivery Date / Return Status
- Payment Received (internal `Rece Amount`)
- Sale Invoice No / Sale Invoice Date / Sale Amount
- Return CN No / Return CN Date / Return CN Amount
- replacement activity remains rolled up to the original order

Flipkart engine:
- Branch Code
- Invoice No / Invoice Date
- CN No / CN Date / ERP CN Amount
- Return Approval Date
- Received Amount

The upload screen now reports source coverage counts for Payment, Replacement,
Invoice, CN and Return Date. If Payment Received maps to zero rows, the workbook
is rejected instead of silently overwriting the database with zeros.


## v12 Attached Reconciliation Output Tracking

The E-Com reconciliation dashboard and download now follow the attached
35-column output exactly.

The previously yellow-highlighted source fields are persisted and tracked:
- Replacement Order Id
- Repl Quantity
- Replacement Item Price
- Courier & Customer Return Qty
- Return Delivery Date
- Rece Amount
- Refund
- Reimbursement
- Sale Invoice No / Sale Invoice Date
- Return Invoice No / Return Invoice Date

Amazon derives replacement, return, payment, invoice and CN activity from the
core engine and source sheets. Replacement activity is rolled into the original
Amazon order.

Pending Task Created Date is generated automatically when a non-reconciled task
first appears. MIR / TEI / Team fields remain blank until the corresponding
branch/team workflow is actually updated; they are not fabricated.


## v13 Exact Amazon Column Mapping

Amazon source mapping is now explicitly aligned to the FY workbook:

### Replacement_Map
- F = Quantity
- H = Replacement Order Id
- I = Original Order Id
- Replacement Item Price = Original Unit Item Price × Replacement Quantity

### Sales
- D = Invoice No
- E = Invoice Date
- F = Po Number
- P = Quantity
- U = Item Price
- BU = Sale/Return
- BU=SALE -> Sale Invoice No / Sale Invoice Date
- BU=RETURN -> Return Invoice No / Return Invoice Date

### Payments
- D = Order ID
- AA = Payment Received
- AB = Transaction Status
- AB contains RELEASED -> summed into Rece Amount
- AB contains DEFERRED -> summed into new Deferred Amount

Payment activity for original + replacement order IDs is rolled back to the
original Amazon order.

### Reverse
- D = Order ID
- G = Courier & Customer Return Qty

Each new portal upload replaces only that marketplace's current reconciliation
snapshot. Persistent Task, MIR and audit/history tables are not deleted.


## v14 — First Tower Engine + Persistent Tower Workflow

This build was made by directly comparing the two user-provided Towers.

The first Tower's Amazon and Flipkart reconciliation engines are now the
calculation source because those engines populate the source fields correctly.

The second Tower's features are retained:
- persistent database
- E-Com and Settlement dashboards
- Pending Task workflow
- Team download/upload
- MIR 1600 / 5400 / 5600 / 7800
- Branch Owner / email setup
- audit/history

Amazon source mapping:
- Replacement Order Id / Quantity: Replacement_Map
- Replacement Item Price: Original Unit Item Price × Repl Quantity
- Courier & Customer Return Qty: Reverse
- Rece Amount: Payments RELEASED
- Deferred Amount: Payments DEFERRED
- Sale Invoice / CN details: Sales source enrichment
- Return Delivery Date: Reverse source enrichment

On every marketplace re-upload only that marketplace's current reconciliation
snapshot is replaced. Task/MIR/audit history is preserved.


## v14.1 Dashboard Visibility
The E-Com dashboard now visibly shows and summarizes:
Replacement Order Id, Repl Quantity, Replacement Item Price,
Courier & Customer Return Qty, Rece Amount and Deferred Amount.
The same fields remain in the downloadable reconciliation.


## v14.2 Verified Dashboard Sync

This version verifies the complete pipeline:
Amazon/Flipkart engine -> normalized DataFrame -> SQLite database -> dashboard.

After every source upload, the app compares the expected engine totals against
the values actually saved in SQLite. A mismatch raises an error instead of
showing a successful upload.

Verified fields include:
- Replacement Orders
- Repl Quantity
- Replacement Item Price
- Courier & Customer Return Qty
- Rece Amount
- Deferred Amount

The latest successful marketplace workbook is also stored under:
`C:\Users\<user>\ReconciliationTowerData\source_snapshots`

This allows future one-click dashboard rebuild without asking the user to
upload the same source workbook again.


## v14.3 Return Type Fix

Amazon Order-wise Reconciliation `Return Type` now comes directly from:
- Sheet: Reverse
- Column A: Return Type
- Values normalized to Customer Return / Courier Return
- Match: Order ID (named Order ID column; Column D fallback)
- Replacement order return type can roll back to the original reconciliation row.


## v14.4 — MIR Branch + Team Task Workflow

The separate Team Update Upload workspace has been removed.

Each branch workspace now supports both pending-task workflow and MIR workflow:

- MIR 1600
- MIR 5400
- MIR 5600
- MIR 7800

Inside each branch:
1. View/filter branch pending tasks.
2. Download pending tasks in the same reconciliation-format workbook.
3. Team updates Team Remarks / Task Status / Task Completed Date.
4. Upload the same working sheet back.
5. Main reconciliation updates automatically.
6. Upload MIR sheet for MIR/TEI/SR updates.
7. Saved MIR details remain viewable/downloadable.

The main reconciliation remains read-only for team users.


## v14.5 — Protected MIR Team Template + Ticket Tracking

Main E-Com reconciliation now includes:
`Ticket Raised if Any` between TEI Date and Team Remarks.

Each MIR branch downloads the exact main reconciliation format.
Only these seven fields are unlocked/editable:
- MIR No
- MIR Date
- TEI No
- TEI Date
- Ticket Raised if Any
- Team Remarks
- Task Completed Date

All other reconciliation/source fields are locked in the XLSX.

Branch upload logic ignores all non-editable source columns and rejects rows
belonging to another branch.

Task Status is automatic:
- Task Completed Date entered -> Completed
- Any permitted update without completion date -> Working

Ticket, MIR, TEI, team remarks, completion date and resulting workflow status
are written back to the persistent database and main reconciliation dashboard.


## v14.6 — Adjustment Tracking

Amazon `Adjustment` is now read from the Payments sheet using the source
header `Adjustment` (with `Adjustment Amount` accepted as an alias).

The amount is:
- grouped by Amazon Order ID,
- taken from the same RELEASED payment activity used for settlement tracking,
- rolled up from replacement order IDs to the original Amazon order,
- stored persistently in SQLite,
- shown in Reconciliation Summary,
- shown order-wise in the E-Com reconciliation,
- included in the downloadable reconciliation.

Adjustment is a source/reconciliation field, therefore it remains LOCKED in
the protected MIR/team working template.

## v14.7 — Correct Adjustment Mapping
Verified source mapping:
Payments Column C / Type = Adjustment; Column D = Order ID;
Column AA / Payment Received = Adjustment amount; Status = Released.
The KPI includes all Released Adjustment rows. Order-wise values are assigned
by Order ID and rolled from replacement IDs to the original order.


## v14.7.1 — Adjustment KPI Runtime Fix
Fixed:
`NameError: name 'view' is not defined`

The unfiltered Adjustment KPI now correctly compares the filtered rows against
the E-Com dashboard dataset `display`.

## v14.8 Adjustment Auto Backfill
Exact source:
Payments C = Type (Adjustment), D = Order ID, AA = Payment Received,
AB = Released.

After each Amazon upload the app directly writes matched order-wise Adjustment
into SQLite. If an older DB still has Adjustment=0, the dashboard automatically
repairs it from the saved Amazon source snapshot.

The KPI retains the complete source Adjustment total. Order-wise rows contain
only amounts that can be matched to an Amazon order or replacement-to-original
relationship.

## v14.9
Added `Ticket Raised if Any` and `Raised Date` to Order-wise Reconciliation.
Raised Date is persisted with pending-task/team workflow and included in
downloadable protected working files.

## v14.9.1 — Exact Ticket Headers
Immediately before `Team Remarks`, Order-wise Reconciliation now uses:
1. Ticket Raised If Any
2. Ticket Raised Date

The internal persistent database field remains `raised_date`; only the
user-facing reconciliation/download/upload header is renamed.

## v14.9.2 — Ticket Columns Visible Fix

Fixed the actual E-Com visible and export column lists.

Exact order is now:
TEI No | TEI Date | Ticket Raised If Any | Ticket Raised Date |
Team Remarks | Task Completed Date

No source re-upload is required merely for these columns to appear.

## v14.10 — Return Type from Reverse Column A
Amazon Order-wise Reconciliation `Return Type` is sourced exactly from
Reverse Column A. Values are normalized to `Customer Return` or `Courier Return`.
Order matching uses the Reverse Order ID and replacement-order activity rolls
back to the original Amazon order. Existing blank persistent rows can be
auto-repaired from the saved Amazon source snapshot.

## v14.10.1 — Return Type ID Mapping Fix

Root cause fixed:
the prior Return Type mapper called `normalize_order_id()`, but that function
does not exist in this app. It now consistently uses the app's real order-ID
normalizer: `clean_id()`.

Verified against the supplied Amazon source:
- Reverse rows with Return Type: 10,341
- Courier Return rows: 6,811
- Customer Return rows: 3,530
- Unique Reverse Order IDs with Return Type: 9,956

The E-Com dashboard now forces one Return Type database sync per Streamlit
session from the saved Amazon source snapshot.

## v14.11 — Automatic Sync on Every Upload

Manual repair clicks are no longer part of the normal workflow.

Whenever a new Amazon workbook is uploaded, the app automatically:
1. runs Amazon reconciliation,
2. persists the reconciliation rows,
3. refreshes Return Type from Reverse Column A,
4. refreshes Adjustment from Payments,
5. saves the new workbook as the latest verified source,
6. verifies the persisted dashboard data.

The Repair / Rebuild buttons remain only as emergency/admin recovery tools.


## v14.12 — Preserve Team History on Source Refresh

New Amazon/Flipkart source uploads refresh source-derived reconciliation data
but preserve all operational/team history.

Preserved independently from source refresh:
- Pending task records
- Task Created Date
- Task Status
- Working Date
- Task Completed Date
- Team Remarks
- Ticket Raised If Any
- Ticket Raised Date
- MIR / TEI / SR details
- Activity audit history

Workflow:
1. snapshot operational history for the marketplace,
2. replace only the marketplace source reconciliation snapshot,
3. restore task/team history,
4. operational overlay automatically merges the preserved history back onto
   the refreshed reconciliation by portal/order/task and branch/order for MIR.

A new upload therefore does not erase earlier team working.

## v14.13 — Flipkart Reimbursement

Verified against the supplied Flipkart source workbook:
- Sheet: Flipkart Payments
- Column G: Po Number
- Column P: Reimbursement
- Non-zero reimbursement rows: 317
- Unique POs with non-zero reimbursement: 295
- Source reimbursement total: ₹438,707.67

The value is summed PO-wise and stored in the existing `reimbursement` field.
Therefore it automatically appears in:
- Reconciliation Summary → Reimbursement KPI
- Order-wise Reconciliation → Reimbursement column
- Filtered/downloaded E-Com reconciliation

Every new Flipkart upload auto-syncs reimbursement; no manual repair button is
required.

## v14.14 — Settlement Process Reconciliation Dashboard Filters

The Settlement Process Reconciliation Dashboard now has exactly these four filters:
1. Order Date Range
2. Invoice Date Range
3. Quarter
4. Year

Quarter and Year use Order Date/Shipment Date. Invoice Date Range uses Sale Date
(Invoice Date). All settlement KPIs, the order-wise table, and the filtered
download recalculate from these filters.

## v14.15 — Settlement Marketplace Filter

Settlement Process Reconciliation Dashboard now has five filters:
1. Market Place
2. Order Date Range
3. Invoice Date Range
4. Quarter
5. Year

Market Place supports All / Amazon / Flipkart based on the data currently
available in the reconciliation dashboard.

## v14.16 — MIR Branches Use Order-wise Reconciliation Format

MIR 1600, MIR 5400, MIR 5600 and MIR 7800 now use the same Order-wise
Reconciliation format as the main E-Com dashboard for:
- on-screen view
- filtered download
- protected team working workbook
- upload back into the main reconciliation

Editable team fields:
MIR No, MIR Date, TEI No, TEI Date, Ticket Raised If Any, Ticket Raised Date,
Team Remarks and Task Completed Date.

The separate legacy MIR-format uploader was removed so there is only one branch
working format. All source/reconciliation columns remain protected.

## v14.17 — MIR Pending Tasks Only

MIR 1600 / 5400 / 5600 / 7800 now show only active task rows:
- Pending
- Working

Completed tasks remain preserved in the main reconciliation, database and audit
history, but are hidden from the MIR team workspace.

The MIR table/download continues to use the full Order-wise Reconciliation format.

## v14.18 — Rece Amount Date and Refund Date

Two columns added to the E-Com Order-wise Reconciliation format and therefore
also to MIR 1600 / 5400 / 5600 / 7800 pending-task views/downloads:

- Rece Amount Date — immediately after Rece Amount
- Refund Date — immediately after Refund

Flipkart uses the existing Payment Date / Refund Date enrichment from Flipkart
Payments. Amazon Rece Amount Date uses the payment transaction release date;
Amazon Refund Date uses the latest transaction release date for refund/reversal
rows in Amazon Payments.

These are source-derived, locked team columns.

## v14.19 — Rece/Refund Dates on Both Main Dashboards

E-Com Process Reconciliation Dashboard:
- Rece Amount
- Rece Amount Date
- Deferred Amount
- Refund
- Refund Date

Settlement Process Reconciliation Dashboard:
- Payment Received
- Payment Received Date
- Deferred Amount
- Refund
- Refund Date

The corresponding filtered Excel downloads use the same column sequence.
MIR branch pending-task views continue to inherit the E-Com Order-wise format.

## v14.20 — Supabase Persistent Database

Production persistence now uses `DATABASE_URL` from Streamlit Secrets.

When `DATABASE_URL` exists:
- reconciliation_master -> Supabase PostgreSQL
- pending_tasks -> Supabase PostgreSQL
- mir_details -> Supabase PostgreSQL
- branch_owner_rules -> Supabase PostgreSQL
- upload_history -> Supabase PostgreSQL
- source_kpis -> Supabase PostgreSQL
- activity_audit -> Supabase PostgreSQL

The existing application SQL is supported through a PostgreSQL compatibility
adapter which translates SQLite `?` placeholders, `AUTOINCREMENT`, and
`PRAGMA table_info(...)` migrations.

When DATABASE_URL is absent, local Windows SQLite remains available as a
development fallback.

Important: newer builds also persist the latest verified Amazon/Flipkart source
workbook in PostgreSQL (`source_workbooks`). The container file is only a working
copy and can be restored from the database after restart/redeploy.

## v14.21 — PostgreSQL Cursor Compatibility

Fixed:
`AttributeError: 'PostgresCompat' object has no attribute 'cursor'`

pandas `read_sql_query()` requires a DB-API connection exposing `.cursor()`.
The Supabase/PostgreSQL compatibility adapter now provides a cursor wrapper
with execute/fetch/description/rowcount/close support and qmark-placeholder
translation.

This allows existing reconciliation dashboard reads to operate against
Supabase PostgreSQL without reverting to local SQLite.

---

## v15.0 — Persistent Control Tower + Safe One-Click Deployment

This build separates **application code** from **business data**.

### Permanent data model

When `DATABASE_URL` is configured in Streamlit Secrets, Supabase PostgreSQL is
the production source of truth for:

- current Amazon / Flipkart reconciliation rows
- latest verified Amazon / Flipkart source workbook blobs
- source upload history
- pending-task history
- team Working / Completed status
- Team Remarks
- Ticket Raised / Ticket Raised Date
- MIR / TEI details
- branch-owner rules
- Admin PIN hash
- Microsoft 365 configuration
- activity audit history

A GitHub commit, Streamlit restart or Streamlit redeploy does **not** intentionally
clear any of these tables.

### Source workbook rule

The latest verified source workbook for each marketplace remains active until an
Admin explicitly uploads a newer workbook for that marketplace.

`source_workbooks` stores one current workbook per portal. `saved_source_path()`
re-materialises that workbook from the database when a repair/rebuild needs a
local path. The Streamlit container copy is therefore only a working cache, not
the permanent source of truth.

### Team-history rule

Source refreshes can update source-derived reconciliation fields and create new
business tasks, but they do **not** auto-close existing team tasks and do not
rewrite existing team remarks/completion dates.

The MIR branch upload is change-aware:

- uploading the same unchanged working sheet is a no-op;
- blank cells do not erase previously saved team values;
- a Completed task is not reopened merely because another field is edited;
- MIR / TEI / Ticket / Team Remark / Completion updates persist in PostgreSQL.

### Upload History rule

Only a real Admin source upload is recorded as a source upload. A manual rebuild
from the already-saved source does not create a fake duplicate upload-history
entry.

### One-click deployment

Use:

`DEPLOY_WEBSITE_ONE_CLICK.bat`

The BAT file now uses its own folder rather than a hard-coded Windows path. It:

1. runs `preflight_check.py`,
2. stages and commits local changes when required,
3. rebases on GitHub `main`,
4. runs pre-flight checks again,
5. pushes `main`,
6. opens the live Streamlit URL.

If GitHub and Streamlit Cloud are already connected, the GitHub push triggers the
normal Streamlit Cloud auto-redeployment.

### One-time production setup

One-time only, open Streamlit Cloud -> App -> Settings -> Secrets and configure:

```toml
DATABASE_URL = "postgresql://<user>:<password>@<host>:<port>/<database>"
```

Use the actual PostgreSQL/Supabase connection string. Never commit the live value
to GitHub. `.streamlit/secrets.toml` and `.env` are excluded by `.gitignore`.

After this is configured, the app should display:

`Storage: Supabase PostgreSQL — persistent cloud database`

If it displays the Local SQLite fallback warning on Streamlit Cloud, do not rely
on that deployment for permanent production history until `DATABASE_URL` is
configured.

### Deployment pre-flight

`preflight_check.py` intentionally blocks deployment for known structural
problems including Python syntax errors, missing required project files, missing
persistence guardrails and the earlier `log_activity(..., branch=...)` keyword
mismatch.
=======
# Reconcilation_System_UPDATE_NOTES_v15.12
Glen Appliances Pvt Ltd
>>>>>>> 4ef9c0fda54f904909d1070f33cb4133df9c3e41
