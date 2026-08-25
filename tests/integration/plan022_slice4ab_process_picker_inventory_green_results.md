# Plan 022 Slice 4ab Process picker inventory GREEN

Date: 2026-08-25

Operator-visible RED: **Production Item Search** opened from a numbered
acceptable-item cell but showed zero rows while seeded managed inventory was
visible in Inventory Viewer.

Focused RED was `1/4` passed and `3/4` RED. The same focused contract is now
`4/4` GREEN. The Core picker's Process-only path queries active exact Inventory
Domain entities using `System_Key`, deduplicates them to managed SKU
alternatives, and writes only managed item/SKU identity into the Process table;
it does not allocate an exact entity during design.

The packaged public-handler proof is `2/2` GREEN through
`mProduction.BtnOpenProductionForm`, the actual Production worksheet selection
event, the Core-owned search form, and a clean Excel restart. It records both
`PickerOpened=True` and `PickerInventoryRows=True` while retaining every Slice
4aa bulk-import field as true.

| Gate | Result |
|---|---:|
| Focused Slice 4ab source | 4/4 |
| Historical Slice 4aa source | 8/8 |
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

Visible confirmation against the user's seeded warehouse remains pending.
