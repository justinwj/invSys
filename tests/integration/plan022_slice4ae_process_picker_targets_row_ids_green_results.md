# Plan 022 Slice 4ae Process picker targets and row IDs GREEN

Date: 2026-08-25

Production Item Search now opens only from valid **Acceptable Managed Item n**
cells. OUTPUT Name is descriptive only. INPUT, REQUIREMENT, OUTPUT, and
INSTRUCTION rows use one table-wide three-character Base-36 namespace; the
two-pass allocator retains existing valid unique IDs while repairing blanks,
invalid values, and duplicates.

| Gate | Result |
|---|---:|
| Focused Slice 4ae source | 6/6 |
| Packaged Production public callback + clean restart | 2/2 |
| Historical Slice 4ad / 4ac / 4ab / 4aa / 4z / 4y source | 6/6; 6/6; 4/4; 8/8; 7/7; 6/6 |
| Packaged XLAM | 74/74 |
| Ribbon/compile | 142/142 |
| Live role workflows | 47/47 |
| Ordered Release 1 chain | 30/30 |
| Launcher contracts | 24/24 |
| Dedicated NAS runtime | 16/16 |
| Deterministic static baseline | 19/19 |
| Reviewed cleanup/growth | 13/13 |

The packaged callback records `OutputPickerOpened=True`,
`OutputPickerCommitted=True`, `OutputSkuRoundTrip=True`,
`OutputNamePickerSuppressed=True`, `UniqueRowIds=True`,
`FirstAssignedIdRetained=True`, and `NoPhysicalKey=True`. Visible confirmation
in the user's saved Production workbook remains pending.
