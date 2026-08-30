# Slice 9 Production Layout Runtime Results

- Packaged geometry: PASS (3 sizes x 5 pages; zero out-of-bounds and zero interactive-control overlaps)
- Native window behavior: PASS (minimize, restore, maximize, restore)
- Screenshots: PASS (minimum/default/expanded Ingredients Assignment page)

| Case | Requested points | Pages | Screenshot |
|---|---:|---:|---|
| minimum | 900x600 | 5 | slice9-layout/production-minimum.png |
| default | 1110x690 | 5 | slice9-layout/production-default.png |
| expanded | 1350x750 | 5 | slice9-layout/production-expanded.png |

| Native action | Result |
|---|---|
| Minimize | PASS |
| Restore | PASS |
| Maximize | PASS |
| MaximizedContentFill | PASS |
| RestoreAfterMaximize | PASS |

Representative packaged reports:

- minimum: `OK|Requested=900.0x600.0|Actual=1110.0x690.0|Page=2|Zoom=100|Anchors=45|OutOfBounds=0|Overlap=0|WindowStyle=Handle=True|Resizable=True|Minimize=True|Maximize=True|Detail=`
- default: `OK|Requested=1110.0x690.0|Actual=1110.0x690.0|Page=2|Zoom=100|Anchors=45|OutOfBounds=0|Overlap=0|WindowStyle=Handle=True|Resizable=True|Minimize=True|Maximize=True|Detail=`
- expanded: `OK|Requested=1350.0x750.0|Actual=1350.0x750.0|Page=2|Zoom=100|Anchors=45|OutOfBounds=0|Overlap=0|WindowStyle=Handle=True|Resizable=True|Minimize=True|Maximize=True|Detail=`
