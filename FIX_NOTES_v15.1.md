# Glen Reconciliation Tower v15.1 - Payment Status Consistency Repair

## Fixed
- Payment Status is no longer initialized as Received.
- Received is assigned only when actual received amount is greater than zero.
- Zero/blank received amount is Pending, except the existing Short Payment remark remains Short Payment.
- Amazon delivered billed order + zero payment -> Payment Pending.
- Amazon open unbilled order + zero payment -> Billing & Payment Pending (no age-based auto-reconciliation).
- Flipkart delivered billed order + zero payment -> Payment Pending.
- Flipkart delivered unbilled order + zero payment -> Billing & Payment Pending.
- Existing persisted legacy rows are repaired automatically once after deployment/startup.
- Legacy delivered rows incorrectly marked Reconciled with zero payment and no refund/reimbursement are corrected to Payment Pending or Billing & Payment Pending.
- Corrected remarks automatically flow into pending-task creation, dashboard counts, Transaction Status, and Excel exports.

## Validation
- app.py, amazon_engine.py and flipkart_engine.py compile successfully.
- preflight_check.py passes.
- Synthetic Amazon and Flipkart zero-payment cases return the expected pending statuses.
