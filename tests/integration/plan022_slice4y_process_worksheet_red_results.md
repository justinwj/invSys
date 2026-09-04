# Plan 022 Slice 4y Process Worksheet RED

- Date: 2026-08-23
- Entry boundary: packaged `mProduction.BtnOpenProductionForm`
- Operator action boundary: actual Process Designer New/line-add/worksheet
  toggle handlers in the captured saved Production workbook
- Runtime: isolated generated test warehouse; row values and machine paths are
  omitted from this committed record

| Check | Result | Expected behavioral RED |
|---|---|---|
| Process ID generation | RED | New Process still received the prior GUID-style identity instead of a locked three-character Base-36 ID. |
| Recipe ID generation | RED | New Recipe still received the prior GUID-style identity instead of a locked three-character Base-36 ID. |
| Requirement ID generation | RED | The actual Add requirement handler required a manually typed ID and did not stage the row. |
| Output ID generation | RED | The actual Add output handler required a manually typed ID and did not stage the row. |
| Identity controls | RED | Process/Recipe identity controls remained operator-editable. |
| Process worksheet toggle | RED | The operator-visible action reached its real handler, which reported `Process worksheet round-trip is not implemented.` |

The focused packaged report was 1 pass / 1 RED. The same run preserved the
five-page Process/Recipe surface, batch-scale bounds, reusable Process/Recipe
lifecycle actions, two exact-key multi-output batches, and clean-session saved
workbook restart. This is meaningful behavioral RED; the harness compiled,
opened the packaged form, bound the expected saved operator workbook, and
continued through the existing GREEN handlers.
