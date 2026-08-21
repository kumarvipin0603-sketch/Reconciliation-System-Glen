# Glen Reconciliation Tower v15.6

## Pending Task filter cleanup
- The **Pending Task** filter now lists only the distinct values currently
  present in the **Pending Remarks** column.
- Hard-coded task names and separate Task Type values are no longer injected
  into this filter.
- Selecting a Pending Task option now filters by the exact Pending Remarks value.
- This applies anywhere the shared global filter is used, including Pending Task,
  All Branch and individual branch task workspaces.

All v15.5 functionality remains included.
