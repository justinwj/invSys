# Demo Workflow Seed Results

- Passed: 12
- Failed: 0

| Check | Result | Contract |
|---|---|---|
| Seed.CompleteKitCount | PASS | One Admin seed event carries exactly 24 R1 workflow inventory entities, including box-making consumables. |
| Seed.MaterialCoverage | PASS | The kit covers raw inputs, WIP, shippable goods, and shipping packaging. |
| Seed.D14Identity | PASS | Every seed entity receives a new System_Key and GOOD condition without a ROW identity path. |
| Seed.CatalogMetadata | PASS | The event carries the catalog metadata needed by Receiving, Production, Shipping, and Viewer projections. |
| DemoLifecycle.FormActions | PASS | The Admin Demo Inventory form exposes separate inventory and data-set actions. |
| DemoLifecycle.IdempotentSeed | PASS | Repeated seed skips active demo item/location/condition groups instead of creating duplicate entities. |
| DemoLifecycle.ExactKeyDelete | PASS | Delete depletes demo entities through exact-System_Key audited adjustments. |
| DemoLifecycle.ValidatedCsvUpload | PASS | Upload accepts a validated demo CSV and preserves generated entity identity. |
| DemoLifecycle.DatasetSelection | PASS | The form lets the administrator choose the built-in kit or an uploaded CSV before Seed mutates inventory. |
| DemoLifecycle.PersistentDatasetLibrary | PASS | Uploaded CSV definitions persist in the selected warehouse library and remain selectable on later form launches. |
| DemoLifecycle.DeleteDatasetDefinition | PASS | The form can delete a selected uploaded dataset definition independently from active demo inventory. |
| DemoLifecycle.R1DatasetImmutable | PASS | The built-in R1 workflow dataset cannot be deleted. |
