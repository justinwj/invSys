# Plan 022 Slice 4k Target Binding RED

- Date: 2026-08-17
- Active slice: Plan 022 Slice 4k — selected warehouse/session binding
- Package: Core, consumed by Operations and Admin
- Public contract: selecting a warehouse from Operations **Send To** must bind
  that warehouse to the current Windows-computer station before sign-in.
- Protecting test:
  `TestNasSelectWarehouseTarget_AutoRegistersCurrentComputerStationAndSignsIn`
- Expected RED: a warehouse whose config still contains only a legacy station
  rejects this computer and leaves the prior warehouse bound.
- Observed RED: 0/1 passed. Selection failed before the current target changed,
  so authentication continued against the prior warehouse.

This was a meaningful behavioral RED: the packaged workbook, fixture, and Excel
harness loaded successfully and the failure was the reported target-binding
behavior.
