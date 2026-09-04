# Plan 022 Slice 4z Process Workbench RED

- Date: 2026-08-24
- Entry boundary: packaged `mProduction.BtnOpenProductionForm`
- Operator boundary: the current Process Designer worksheet action handler
- Runtime: isolated generated warehouse; machine paths and row data omitted

The source contract was 1/7 (documentation only). The packaged reusable
Production run was 1 pass / 1 behavioral RED. The clean-restart Recipe and
Slice 4y selected-workbook round trip remained GREEN before the new probe.

| Check | Result | Expected behavioral RED |
|---|---|---|
| Separate actions | RED | The form still exposes one Edit/Retrieve toggle. |
| Multiple tables | RED | The second action attempts retrieval and only one table remains. |
| Selected-table import | RED | Retrieval is bound to one global outstanding-table name. |
| Record Type dropdown | RED | Cells have no validation choices. |
| Formula-owned Percent | RED | A quantity entered after creation receives no Percent/basis formula. |
| Generated output design | RED | Output Design ID is blank/operator-authored. |
| Output Item Code removal | RED | The worksheet still exposes Item Code. |
| Ingredient Assignment | RED | Requirement alternatives and acceptable managed items are not exposed. |
| Existing item search | RED | No Process-table acceptable-item cell routes to the Core picker. |

The exact packaged result was:
`SeparateActions=False|MultipleTables=False|SelectedOnly=False|RecordTypeDropdown=False|CalculatedPercent=False|GeneratedDesign=False|ItemCodeRemoved=False|Assignments=False|ItemSearch=False|Tables=1`.
