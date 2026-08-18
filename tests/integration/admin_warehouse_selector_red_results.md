# Admin Warehouse Selector RED

- Date: 2026-08-17
- Focused test: `tests/tooling/Test-Plan022Slice4lRibbonSessionControls.ps1`
- Result: **RED**, 9 passed / 1 failed

| Check | Result | Expected behavioral gap |
|---|---|---|
| `Ribbon.AdminWarehouseSelector` | FAIL | The Admin ribbon had no live `Send To` warehouse dropdown and did not invalidate an Admin selector after target selection. |

The other nine session, sign-in/out, target-refresh, and fail-closed contracts
remained GREEN. This is meaningful RED for the requested Admin RibbonX control,
not a compile, fixture, workbook, or harness failure.
