# Glen Reconciliation Tower v15.10 — All Branch Upload Matching Fix

## Fixed
- All Branch Excel Branch Codes are now normalized with `clean_id()`.
- Values such as `1600.0`, `5400.0`, `5600.0`, `7800.0` now correctly match
  database branch codes `1600`, `5400`, `5600`, `7800`.
- System Mapping, uploaded rows and DB task/MIR keys now use the same normalized
  Branch Code + Order No key format.
- Unmatched diagnostics now include `BranchCode:OrderNo` for faster troubleshooting.

## Completion message correction
- A green success notification is shown only when at least one task record is actually updated.
- If the workbook processes but contains no new changes, the app shows a warning instead of
  the misleading “0 task record(s) updated” success message.

All v15.9/v15.8 speed and stability improvements and all earlier business rules remain retained.
