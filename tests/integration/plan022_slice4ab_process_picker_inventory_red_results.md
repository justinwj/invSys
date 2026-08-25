# Plan 022 Slice 4ab Process picker inventory RED

Date: 2026-08-25

Operator-visible RED: **Production Item Search** opens from a numbered
acceptable-item cell but reports no managed inventory while the same seeded
inventory is visible in Inventory Viewer.

Focused command:

```powershell
tests/tooling/Test-Plan022Slice4abProcessPickerInventory.ps1 -RepoRoot .
```

Result: meaningful behavioral contract RED, `1/4` passed and `3/4` RED.

- D15, Plan 022 Slice 4ab, and controls v1 agree on exact `System_Key`
  inventory identity.
- `cDynItemSearch.BuildInventoryPickerItemsFromTable` still requires prohibited
  legacy `ROW`, so current-schema managed rows are discarded.
- The packaged public selection-event proof verifies only that the picker opens;
  it does not yet require a nonempty managed-inventory result.

The picker opened successfully in the packaged operator workflow, so this is
not a form, workbook, fixture, compile, or harness failure.
