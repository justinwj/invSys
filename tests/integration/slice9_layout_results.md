# Slice 9 Production Layout Runtime Results

- Packaged geometry: PASS (3 sizes x 4 pages; zero out-of-bounds and zero interactive-control overlaps)
- Native window behavior: PASS (minimize, restore, maximize, restore)
- Screenshots: PASS (minimum/default/expanded production-run page)

| Case | Requested points | Pages | Screenshot |
|---|---:|---:|---|
| minimum | 900x600 | 4 | slice9-layout/production-minimum.png |
| default | 1110x690 | 4 | slice9-layout/production-default.png |
| expanded | 1350x750 | 4 | slice9-layout/production-expanded.png |

| Native action | Result |
|---|---|
| Minimize | PASS |
| Restore | PASS |
| Maximize | PASS |
| RestoreAfterMaximize | PASS |

Representative packaged reports:

- minimum: `OK|Requested=900.0x600.0|Actual=1110.0x690.0|Page=2|Zoom=60|Anchors=44|OutOfBounds=0|Overlap=0|WindowStyle=Handle=True|Resizable=True|Minimize=True|Maximize=True|Detail=`
- default: `OK|Requested=1110.0x690.0|Actual=1110.0x690.0|Page=2|Zoom=60|Anchors=44|OutOfBounds=0|Overlap=0|WindowStyle=Handle=True|Resizable=True|Minimize=True|Maximize=True|Detail=`
- expanded: `OK|Requested=1350.0x750.0|Actual=1350.0x750.0|Page=2|Zoom=60|Anchors=44|OutOfBounds=0|Overlap=0|WindowStyle=Handle=True|Resizable=True|Minimize=True|Maximize=True|Detail=`
