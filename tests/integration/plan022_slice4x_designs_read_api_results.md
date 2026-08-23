# Plan 022 Slice 4x Designs Read API Results

- Date: 2026-08-23
- Active slice: Plan 022 Slice 4x, reusable Production Processes and Recipe graphs
- Package boundary: Designs Domain -> Core -> Operations primitive bridge

## Focused RED and GREEN

| Boundary | RED | GREEN | Protected contract |
|---|---:|---:|---|
| Designs Domain public bridge, tests 276-278 | 0/3 | 3/3 | Released-only Process/Recipe lists, exact-version serialized Process/Recipe envelopes, released-Recipe validation, and read-only projection counts. |
| Core cross-XLAM compatibility bridge, test 279 | 0/1 | 1/1 | Primitive arrays/strings round-trip through one declared Designs Domain query dispatcher. |
| Operations primitive bridge, test 280 | 0/1 | 1/1 | Role packages consume the Core boundary through direct typed calls and never call the Designs Domain XLAM directly. |

The RED harnesses compiled and completed their schema/lifecycle fixtures. They
returned 0 only when the requested public bridge function was unavailable;
there was no compile failure, missing workbook, or broken fixture.

## Regression and package evidence

- Reusable Production range 264-280: 17/17 GREEN.
- Designs/Core/Inventory neighborhood 262-300: 38/39. The sole failure is the
  pre-existing `TestInventoryQueries_PickerPublishesEverySkuLocation` legacy
  `ROW` assertion recorded in the 2026-08-17 baseline; no changed source or test
  uses that prohibited identity path.
- Published five-XLAM compile/load/Ribbon validation: 74/74 GREEN.
- Clean-state packaged `mProduction.BtnOpenProductionForm` inspection preserved
  the expected surface RED: second launch opened zero additional workbooks,
  `.001%`/`100%`/`1000%` List scaling remained GREEN, and the only contract
  failure was the unchanged four-page legacy Recipe Builder surface without
  Process Designer or Recipe Designer.
- Static maintenance: 967 candidates versus 968 baseline; duplicate-body
  groups 185/185; literal `Application.Run` targets 8/8; unresolved dynamic
  calls improved from 47 to 45.
- The existing `modDesignsApply` module/procedure split candidates remain open
  and unchanged. No bloat or dynamic-call exception was accepted.

The read APIs return arrays containing primitive list values and JSON strings
for exact Process/Recipe definitions. `ValidateReleasedRecipe` returns a
tab-delimited primitive status envelope. All same-project calls are direct and
typed; Core consolidates the declared cross-XLAM reads through one
`Application.Run` call site.

Generated range reports, harness workbooks, and static runtime reports are
ignored machine evidence. This sanitized record contains no operational row,
user, credential, or warehouse data.
