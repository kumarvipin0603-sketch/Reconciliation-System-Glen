# Glen Reconciliation Tower v15.5

## Refund status correction
- Rece Amount < 0 -> Refund
- Rece Amount = 0 -> Pending / applicable pending workflow
- Rece Amount > 0 -> Received
- Negative receipt takes priority over Short Payment.

The existing startup consistency repair recalculates saved rows through the same
payment-status function, so legacy negative receipts are corrected after deploy.

All previous v15.4 rules and dashboard changes remain included.
