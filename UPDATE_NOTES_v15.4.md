# Glen Reconciliation Tower v15.4

## Return-to-TEI aging rule
- If Return Delivery Date is 2 calendar days old or more and TEI is still not created,
  Pending Remarks becomes **Return Received TEI Pending**.
- Once TEI is created, this return-aging override stops.
- Existing v15.3 rule remains: after 5 calendar days from TEI Date, if CN is still
  pending, Pending Remarks becomes **TEI Generated but CN Pending**.
- Applied in the shared operational master so the status is reflected across
  dashboards, branch/all-branch working views and exports.
