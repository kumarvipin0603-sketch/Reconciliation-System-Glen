# Glen Reconciliation Tower v15.16 — Fast Task/Owner Mapping

## Exact bottleneck fixed
The E-Com dashboard spinner was spending most of its time in `load_tasks()`.
For every task row, `get_assignees()` reloaded `branch_owner_rules` from
Supabase. With ~3,973 tasks this could create thousands of remote database
queries on one dashboard opening.

## Changes
- Pending tasks are fetched once.
- Branch owner rules are fetched once.
- Exact branch/task and All Branch fallback assignments are mapped in memory.
- Task aging is vectorized instead of DataFrame row-wise `apply()`.
- Owner rules are cached for other screens.
- Owner-rule updates explicitly invalidate dashboard caches.
- v15.15 instant display overlay and v15.14 non-blocking startup are retained.
- All persistent Supabase business data and v15.12 Reconciled -> Completed rules remain unchanged.

Expected result: after loading the 23k reconciliation master, the task/MIR overlay
should complete in seconds rather than issuing thousands of Supabase calls.
