# Plan 022 Slice 4y Process Worksheet GREEN

- Date: 2026-08-23
- Entry boundary: packaged `mProduction.BtnOpenProductionForm`
- Operator action boundary: actual Process Designer New/line-add/worksheet
  toggle handlers in the captured saved Production workbook
- Runtime: isolated generated test warehouse; row values and machine paths are
  omitted from this committed record

| Check | Result | Evidence |
|---|---|---|
| Generated identities | GREEN | New Process, Recipe, Requirement, and Output actions produced locked three-character Base-36 IDs. |
| Formula formulation | GREEN | 100 + 200 + 11.2 + 300 lb produced basis 611.2 and displayed percentages 16.4, 32.7, 1.8, and 49.1, totaling 100.0%. |
| Mixed UOM safety | GREEN | A deliberate LB/KG edit was rejected; the bound table remained available for correction. |
| Successful retrieval | GREEN | The corrected table replaced the form draft and only its uniquely named temporary table was removed. |
| Repeat editing | GREEN | The retrieved Process was sent to the sheet and retrieved again through the same toggle handler. |
| Saved-workbook restart | GREEN | A third outstanding table was saved, rediscovered after a clean Excel restart against the same operator workbook, retrieved, and removed. |
| Preserved reusable workflow | GREEN | Process/Recipe save, release, obsolete, reuse, assignment, graph/order, two exact-key multi-output batches, co-product/yield basis, and clean recipe reload remained GREEN. |

Focused source contract: 6/6 GREEN. Focused packaged callback/restart evidence:
2/2 GREEN. The package report also preserved `0.001%`, `100%`, and `1000%`
scaling, insufficiency/stale-key rejection, routed intermediate consumption,
distinct output keys, and saved-workbook reuse.
