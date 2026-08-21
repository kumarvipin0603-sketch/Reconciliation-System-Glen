# Glen Reconciliation Tower v15.7

## Quantity-level TEI tracking
A new persistent **TEI Qty** field has been added.

### Workflow
- Return Delivery Date + 2 days and no TEI -> `Return Received TEI Pending`
- TEI exists but TEI Qty is less than returned quantity ->
  `Partial TEI Generated - Balance TEI Pending`
- TEI Qty fully covers returned quantity and CN is not yet available -> `CN Pending`
- 5 days from TEI Date with CN still pending -> `TEI Generated but CN Pending`

### VLOOKUP / duplicate protection
A TEI number with blank TEI Qty is conservatively treated as covering **1 unit only**.
Therefore, an order with Return Qty 2 cannot be treated as TEI-complete merely because
the first TEI number was copied against the order. The team must confirm TEI Qty 2
(or the actual full returned quantity).

### Team working
`TEI Qty` is editable in protected Pending Task / branch / All Branch working downloads.
TEI Pending tasks cannot be completed while TEI Qty is below returned quantity.

All v15.6 and earlier fixes are retained.
