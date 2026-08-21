# Glen Reconciliation Tower v15.12

## Global Reconciled -> Completed rule
Whenever `Pending Remarks = Reconciled`:
- all active Pending/Working tasks for that order become `Completed`
- `Task Completed Date` is automatically set to the current date if blank
- an existing completion date is preserved
- the task disappears from the active Pending Task dashboard
- the completed record remains available in history

Applied on startup, after source uploads, after team-working uploads, and in the
Pending Task dashboard display guard.

All previous v15.11 and earlier rules are retained.
