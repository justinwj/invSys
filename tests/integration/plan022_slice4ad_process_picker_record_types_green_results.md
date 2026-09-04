# Plan 022 Slice 4ad Process picker record-type GREEN

Date: 2026-08-25

Entering **Acceptable Managed Item 1** now opens the same Core Production Item
Search for INPUT and OUTPUT rows. INPUT continues to fill its exact numbered
alternative pair. OUTPUT fills the visible managed-item selector, hidden
Accepted SKU mirror, and canonical hidden Output SKU; a nonblank descriptive
Output Name is retained. Retrieval persists the SKU without a physical
`System_Key`.

| Gate | Result |
|---|---:|
| Focused Slice 4ad source | 6/6 |
| Historical Slice 4ac / 4ab / 4aa / 4z source | 6/6; 4/4; 8/8; 7/7 |
| Packaged Production public callback + clean restart | 2/2 |
| Packaged XLAM | 74/74 |
| Ribbon/compile | 142/142 |
| Live role workflows | 47/47 |
| Ordered Release 1 chain | 30/30 |
| Launcher contracts | 24/24 |
| Dedicated NAS runtime | 16/16 |
| Deterministic static baseline | 19/19 |
| Reviewed cleanup/growth | 13/13 |

The packaged result records `OutputPickerOpened=True`,
`OutputPickerCommitted=True`, `OutputSkuRoundTrip=True`,
`OutputNameRetained=True`, and `NoPhysicalKey=True`. Visible confirmation in the
user's saved Production workbook remains pending.
