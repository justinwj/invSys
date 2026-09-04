# Plan 022 Slice 4aa Process bulk-import GREEN

Date: 2026-08-24

The focused behavioral RED was `1/8` passed and `7/8` RED against the compiling,
packaged-GREEN Slice 4z baseline. After implementation, the same focused source
contract is `8/8` GREEN and the historical Slice 4z contract remains `7/7`
GREEN.

The packaged public-handler proof is `2/2` GREEN through
`mProduction.BtnOpenProductionForm`, the actual Process worksheet selection
event, the actual Retrieve Selected Process form handler, and a clean Excel
restart. It records:

- `TextSafeIds=True` and `RequirementIds=True`; numeric-only Base-36 identity
  `001` remains text-safe through staging, processor ingress, Designs Domain,
  projection, and clean restart;
- `UomCatalog=True`, with worksheet validation sourced from Settings' Recipe
  UOM Catalog;
- `NumberedAlternatives=True` and `AddedAlternative=True`, including four
  initial managed-item/SKU pairs and an appended fifth pair;
- `PickerOpened=True`, proving that entering a numbered acceptable-item cell
  invokes the existing Core item-search UI and commits to the matching pair;
- `MultiAreaSelection=True` and `MultiTableDrafts=True`, proving Ctrl+click
  selection resolves two Process tables, persists both through the public
  Process save path, deletes only successful selected tables, and leaves the
  unselected historical workbench tables intact; and
- the accepted formula example remains GREEN at 611.2 lb total and
  16.4/32.7/1.8/49.1 percent, totaling 100.0%.

Regression evidence:

| Gate | Result |
|---|---:|
| Focused Slice 4aa source | 8/8 |
| Historical Slice 4z source | 7/7 |
| Packaged Production public callback + clean restart | 2/2 |
| Packaged XLAM | 74/74 |
| Ribbon/compile | 142/142 |
| Live role workflows | 47/47 |
| Ordered Release 1 chain | 30/30 |
| Launcher contracts | 24/24 |
| Dedicated NAS runtime | 16/16 |
| Deterministic static baseline | 19/19 |
| Reviewed cleanup/growth | 13/13 |
| Production layout | 3 sizes x 5 pages + native window transitions |

Visible operator acceptance remains pending; Production Run - Tree remains
experimental and outside Release 1 acceptance.
