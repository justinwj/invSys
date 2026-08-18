# Demo Data Set Library RED

- Date: 2026-08-17
- Focused test: `tests/tooling/Test-DemoWorkflowSeed.ps1`
- Result: **RED**, 9 passed / 3 failed

| Check | Result | Expected behavioral gap |
|---|---|---|
| `DemoLifecycle.PersistentDatasetLibrary` | FAIL | Uploaded CSV definitions were temporary file selections and did not persist in a selectable warehouse library. |
| `DemoLifecycle.DeleteDatasetDefinition` | FAIL | No action existed to delete a stored dataset independently from demo inventory. |
| `DemoLifecycle.R1DatasetImmutable` | FAIL | No permanent R1 dataset identity or deletion guard existed. |

The prior seed, inventory-delete, CSV-validation, identity, and selection
contracts remained GREEN. This is behavioral RED for the new dataset-library
contract, not a missing workbook, compile, fixture, or harness failure.
