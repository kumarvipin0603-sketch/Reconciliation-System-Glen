# Glen Reconciliation Tower v15.8 — Performance & Stability Review

## Team working upload performance
- Rebuilt branch and All Branch working-sheet import as a batched transaction.
- Excel/System Mapping are parsed/indexed once.
- MIR, task and return-quantity state are prefetched in bulk.
- Repeated per-row DB connections/commits were removed.
- MIR updates, task updates and audit history use executemany + one commit.
- v15.7 TEI Qty / partial-return protections remain intact.

## Fast dashboard reflection
- After a successful Amazon/Flipkart source save, the app invalidates read caches and performs one clean Streamlit rerun.
- After branch/All Branch team updates, the app invalidates read caches and reruns immediately.
- A visible completion trigger is shown after rerun: “Upload complete … dashboards refreshed.”
- Source upload completion also reports elapsed processing time.

## Reliability
- Existing row-count verification, reconciliation integrity checks, persistent source snapshot,
  team-history preservation, payment/refund rules, TEI aging and Pending Remarks rules remain enabled.

## Database query speed
- Added idempotent indexes for branch/order task lookups, task status/history,
  reconciliation branch queries, MIR order lookup and audit chronology.
- These improve response time as persistent Supabase history grows.
