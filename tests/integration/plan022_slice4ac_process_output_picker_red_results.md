# Plan 022 Slice 4ac Process OUTPUT picker RED

Date: 2026-08-25

Operator-visible RED: current managed inventory appears in **Production Item
Search** for INPUT acceptable-item cells, but entering an OUTPUT Name cell does
not open the search tool even though every OUTPUT must identify a managed item.

Focused command:

```powershell
tests/tooling/Test-Plan022Slice4acProcessOutputPicker.ps1 -RepoRoot .
```

Result: meaningful behavioral contract RED, `1/6` passed and `5/6` RED.

- Architecture v4.11, Plan 022 Slice 4ac, and controls v1 agree on a
  picker-selected output SKU distinct from generated Design identity.
- OUTPUT Name is not a worksheet picker target.
- Core picker commit only resolves numbered INPUT alternative pairs.
- No hidden Output SKU round-trip or packaged public-handler assertion exists.

The existing INPUT picker and managed-inventory projection were GREEN, so this
was not a form, inventory, workbook, compile, fixture, or harness failure.
