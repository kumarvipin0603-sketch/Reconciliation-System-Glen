# FBA Reconciliation Module

This branch introduces the reusable database and discrepancy engine for the FBA Stock Reconciliation module.

## Phase 1 implemented

- Persistent FBA discrepancy case register
- Case event/audit history
- Case-to-source match table
- Rule configuration table
- SKU/FNSKU master
- FC-to-plant master
- Shipment, shipment-line, return and reimbursement tables
- Controlled status normalization
- Stable case numbers generated from rule + plant + SKU + source reference
- First automatic discrepancy rules:
  - FBA-003 Inbound shipment short
  - FBA-004 Inbound shipment excess
  - FBA-017 Claim window expired unresolved
- Automatic closure when open quantity becomes zero
- PostgreSQL/Supabase and SQLite compatible SQL through the existing app `db()` abstraction

## Integration contract

In `app.py`:

```python
import fba_reconciliation as fba
```

After the existing `init_db()` is available:

```python
fba.init_fba_db(db)
```

Add `FBA Stock Reconciliation` to the Workspace selector. The UI should call:

```python
cases = fba.load_cases(db)
summary = fba.case_summary(cases)
```

Admin upload handlers should normalize the source workbook into the FBA source tables, then run:

```python
fba.detect_shipment_cases(db)
fba.auto_close_zero_cases(db)
invalidate_read_cache()
```

## Next implementation phase

1. Add `.xlsb` reader dependency (`pyxlsb`) because the current production requirements support xlsx/xlsm but not xlsb.
2. Build source-specific importers for:
   - ASIN vs SKU
   - MSD-Stock
   - Detailed-Stock
   - Stock In
   - Stock Out
   - FBA Transfer
   - FBA Return
   - Shipment
   - Reimbursement
3. Add normalized Amazon inventory-event and ERP-stock tables after importer field names are locked.
4. Add rules FBA-001 through FBA-020 incrementally and test each against the supplied FBA Stock Reconciliation workbook.
5. Add the Control Tower and Action Required Streamlit screens.
6. Add owner assignment and Microsoft 365 owner-wise email distribution using the app's existing Graph integration.

## Safety / deployment

Do not merge this branch into `main` until the importer has been tested with the supplied workbook and the Streamlit UI has been added. The current production reconciliation remains untouched on `main`.
