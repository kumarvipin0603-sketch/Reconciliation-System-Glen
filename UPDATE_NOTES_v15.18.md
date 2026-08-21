# Glen Reconciliation Tower v15.18 — Management MIS & Exception Analytics

## New Management MIS workspace
A new `Management MIS` workspace has been added without changing the existing operational dashboards.

### Executive MIS
- Total Orders
- Reconciliation Rate
- Open Tasks
- Overdue Tasks vs configurable SLA
- Return Rate
- Completed Tasks
- Order Value
- Received Amount
- Refund
- Deferred Amount
- Marketplace health table
- Top open exception chart

### Branch Performance
Branch-wise scorecard with:
- Orders
- Reconciliation %
- Return %
- Open Tasks
- Overdue Tasks
- Completed Tasks
- Task Completion %
- SLA Compliance %
- Average Task Aging
- Order Value
- Received Amount
- Open vs Overdue task chart

### Aging & SLA
- User-configurable task SLA target in days
- Aging buckets: 0–2, 3–5, 6–10, 11–20, 21+ days
- Open Tasks / Within SLA / Overdue / SLA Compliance KPIs
- Detailed overdue task table with marketplace, branch, order, task, owner, email and team remarks

### Exception Analytics
- Exception mix from persisted Pending Remarks
- Payment / Short Payment / Billing / MIR / TEI / Partial TEI / CN / Refund / Replacement categories
- Exception counts and financial exposure fields
- High-value exception order list

### Management Filters
- Marketplace
- Branch
- Configurable SLA days

## Architecture
- Uses the same persistent Supabase reconciliation and task data.
- No separate upload or database is required.
- Existing E-Com, Settlement, Pending Task, All Branch Update, MIR, owner setup, Upload History and Audit Trail workspaces are retained.
- No source or team data is deleted or migrated.
