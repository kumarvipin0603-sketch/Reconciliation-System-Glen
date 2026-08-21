# Glen Reconciliation Tower v15.15 — Instant Dashboard Render

## Root cause
Normal E-Com dashboard viewing was still rerunning Return/TEI/CN business-rule
calculations across the entire 23k+ reconciliation master before rendering.

## Fix
- Dashboard now reads the persisted business result directly.
- Normal page rendering only overlays Task and MIR/TEI working tables.
- Return/TEI/CN and reconciliation rules continue to run at source/team update time.
- Removed the large DataFrame cache/hash/serialization from `ecom_process_display`.
- Optimized MIR fallback with indexed mapping.
- Dashboard heading renders immediately with a loading spinner while the small
  operational overlay is attached.

## Retained
- Supabase persistent storage.
- v15.14 non-blocking startup.
- v15.12 Reconciled -> Completed + one-time Task Completed Date.
- All Amazon/Flipkart reconciliation rules and source persistence.
