# Glen Reconciliation Tower v15.14 — Non-Blocking Startup

## Root cause fixed
The app was executing several full Supabase repair/backfill scans during Python
module import. With 23k+ reconciliation rows and thousands of task rows this could
delay or appear to hang before the dashboard rendered.

## Changes
- Removed heavy import-time calls:
  - payment consistency repair
  - returned/unbilled repair
  - reconciled task auto-completion sweep
  - missing-task backfill
- Added `run_deferred_maintenance()` so those operations run only after a
  successful source/team update, where they are actually required.
- Normal app startup now performs only lightweight database COUNT diagnostics.
- Dashboard rendering from the existing Supabase data starts immediately.
- v15.13 vectorized dashboard improvements are retained.
- v15.12 Reconciled -> Completed business rule remains enforced after updates.

## UI
Build label:
`Build: v15.14 — Non-Blocking Startup + Fast Dashboard`
