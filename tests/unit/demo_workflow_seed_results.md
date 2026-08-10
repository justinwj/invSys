# Demo Workflow Seed Results

- Passed: 4
- Failed: 0

| Check | Result | Contract |
|---|---|---|
| Seed.CompleteKitCount | PASS | One Admin seed event carries exactly 19 R1 workflow inventory entities. |
| Seed.MaterialCoverage | PASS | The kit covers raw inputs, WIP, shippable goods, and shipping packaging. |
| Seed.D14Identity | PASS | Every seed entity receives a new System_Key and GOOD condition without a ROW identity path. |
| Seed.CatalogMetadata | PASS | The event carries the catalog metadata needed by Receiving, Production, Shipping, and Viewer projections. |
