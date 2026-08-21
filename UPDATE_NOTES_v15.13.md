# Glen Reconciliation Tower v15.13 — Fast Dashboard Refresh

## Dashboard performance fix
- Confirmed source upload persistence was working; the delay happened after save.
- Vectorized the full-master team/MIR/SR remarks assembly.
- Vectorized Return/TEI/CN aging logic across the master dataset.
- Removed row-wise Pending Remarks -> Task Type conversion; unique remarks are mapped once.
- Cached the final E-Com display for short repeated UI reruns.
- Removed unnecessary SQL ORDER BY work from full bulk master/task/MIR dashboard reads.
- Business rules from v15.12 remain unchanged.

## Build label
- UI now displays: `Build: v15.13 — Fast Dashboard Refresh + v15.12 Reconciled Auto Complete`.

## Existing v15.12 behavior retained
- Pending Remarks = Reconciled -> active tasks Completed.
- Task Completed Date is filled once and preserved.
- Return/TEI/CN quantity-level workflow remains active.
- Supabase persistence and team/MIR history remain intact.
