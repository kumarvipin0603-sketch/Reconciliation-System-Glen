# Glen Reconciliation Tower v15.11

## Returned + no billing + no payment closure

New priority rule:

- Courier & Customer Return Qty > 0
- Return Delivery Date / confirmed return date exists
- Sale Invoice No is blank
- Invoice Qty / ERP Billed Qty = 0
- Rece Amount = 0
- Refund = 0

Result:
- Payment Status remains `Pending` (because no payment was received)
- Pending Remarks -> `Reconciled`
- Transaction Status -> `Reconciled`
- TEI/CN aging is bypassed
- Existing Pending/Working task rows for the order are auto-completed and no
  longer appear in the default Pending Task dashboard.

The rule is applied:
1. In Flipkart source reconciliation.
2. Before master persistence/task creation.
3. In the shared dashboard overlay.
4. As an idempotent startup repair for already-saved data.

All v15.10 and earlier performance, Supabase persistence, payment/refund,
TEI Qty, branch matching and Pending Task filter fixes remain retained.
