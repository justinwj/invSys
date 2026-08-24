# Plan 022 Slice 4aa Process bulk-import RED

Date: 2026-08-24

Focused command:

```powershell
tests/tooling/Test-Plan022Slice4aaProcessBulkImport.ps1 -RepoRoot .
```

Result: expected behavioral contract RED, `1/8` passed and `7/8` RED.

- Normative D15, Plan 022 Slice 4aa, and controls v1 are reconciled.
- Generated worksheet identities are not yet text-safe and INPUT Requirement ID
  is not generated into its managed column.
- UOM has no Recipe UOM Catalog validation.
- The worksheet has one unnumbered acceptable-item pair and no add-pair action.
- Core commits only to the unnumbered pair.
- the packaged public callback has no Ctrl+click multi-table DRAFT-import proof
  and no actual picker-open proof.

This is meaningful RED against the already compiling, packaged-GREEN Slice 4z
baseline; it is not a harness, fixture, workbook, or compile failure.
