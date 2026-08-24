# Plan 022 Slice 4z Process Workbench GREEN

- Date: 2026-08-24
- Entry boundary: packaged `mProduction.BtnOpenProductionForm`
- Operator boundary: `Create Process Table` and `Retrieve Selected Process`
  form handlers
- Runtime: isolated generated warehouse; machine paths and row data omitted

The focused source contract is 7/7 GREEN. The packaged reusable Production
launcher and clean-restart gate is 2/2 GREEN.

| Check | Result | Evidence |
|---|---|---|
| Separate actions | GREEN | Both operator controls and their distinct form handlers were exercised. |
| Multiple tables | GREEN | Three uniquely named Process tables coexisted in the captured workbook. |
| Selected-table import | GREEN | Retrieving table one left tables two and three unchanged. |
| Save/reopen | GREEN | The two remaining tables were rediscovered after a clean Excel restart. |
| Record Type dropdown | GREEN | Generated rows use list validation for `INPUT`, `OUTPUT`, `INSTRUCTION`, and `ALTERNATIVE`. |
| Formula-owned Percent | GREEN | 611.2 lb produced 16.4%, 32.7%, 1.8%, and 49.1%, totaling 100.0%. |
| Generated output design | GREEN | Output Design ID/version are formula generated from Process/Output identity. |
| Output Item Code removal | GREEN | No operator-visible Item Code column remains. |
| Ingredient Assignment | GREEN | Requirement ID, acceptable managed item, and managed SKU round-trip columns are present. |
| Existing item search | GREEN | Acceptable-item selection routes to the Core Production picker and writes item/SKU. |

Packaged result:
`SeparateActions=True|MultipleTables=True|SelectedOnly=True|RecordTypeDropdown=True|CalculatedPercent=True|GeneratedDesign=True|ItemCodeRemoved=True|Assignments=True|ItemSearch=True`.

Clean-restart result:
`WorksheetRediscovered=True|WorksheetRetrieved=True|MultipleTablesRediscovered=True|SelectedOnly=True|AllRetrieved=True`.

Regression results:

- packaged XLAM: 74/74;
- Ribbon/compile: 142/142;
- live role workflows: 47/47;
- ordered Release 1 chain: 30/30;
- Plan 022 launcher contracts: 24/24;
- dedicated NAS safety across two clean sessions: 16/16;
- deterministic static baseline: 19/19; and
- reviewed cleanup/ratchets: 13/13.

Visible layout RED found the new Retrieve control overlapping Description at the
expanded size. After relocation and widening, the packaged form is GREEN across
all five pages at minimum/default/expanded sizes and through native
minimize/restore/maximize transitions. Operator-visible workbench UAT remains
pending.
