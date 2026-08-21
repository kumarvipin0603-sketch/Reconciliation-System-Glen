# Glen Reconciliation Tower v15.2

## Changes

1. Added **All Branch Pending Task Update** workspace.
   - Shows active Pending/Working tasks across all branch codes.
   - Supports the same protected working-sheet download used by branch teams.
   - One uploaded workbook may contain multiple branch codes.
   - Branch Code, source and reconciliation columns remain locked.
   - Approved editable fields remain: MIR No, MIR Date, TEI No, TEI Date, Ticket Raised If Any, Ticket Raised Date, Team Remarks, Task Completed Date.
   - Task status changes automatically to Working/Completed.
   - Updates are written to the same persistent pending-task/MIR tables and therefore flow automatically to all dashboards and individual MIR branch workspaces.

2. Dashboard names changed:
   - `E-Com Process Reconciliation Dashboard` -> `E-Com Reconciliation Dashboard`
   - `Settlement Process Reconciliation Dashboard` -> `Settlement Dashboard`

3. Previous v15.1 payment-status fix retained.

## Validation

- Python compile: PASS
- Project preflight: PASS
