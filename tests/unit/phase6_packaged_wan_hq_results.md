# Phase 6 Packaged WAN HQ Validation Results

- Date: 2026-03-25 19:30:05
- Deploy root: C:\Users\Justin\repos\invSys_fork\deploy\current
- Session root: C:\Users\Justin\AppData\Local\Temp\invsys-phase6-wanhq-97a725a50c334b36a9d9cc331351ee49
- Passed: 9
- Failed: 1

| Check | Result | Detail |
|---|---|---|
| Setup.RuntimeRoots | PASS | SessionRoot=C:\Users\Justin\AppData\Local\Temp\invsys-phase6-wanhq-97a725a50c334b36a9d9cc331351ee49 |
| Packaged.OpenA | PASS | Core+Inventory.Domain |
| Packaged.OpenB | PASS | Core+Inventory.Domain |
| Packaged.OpenHQ | PASS | Core+Inventory.Domain |
| Packaged.RuntimeOverrides | PASS | WH97=C:\Users\Justin\AppData\Local\Temp\invsys-phase6-wanhq-97a725a50c334b36a9d9cc331351ee49\WH97; WH98=C:\Users\Justin\AppData\Local\Temp\invsys-phase6-wanhq-97a725a50c334b36a9d9cc331351ee49\WH98 |
| Publish.WH97.Initial | PASS | EventID=EVT-WH97-20260325192956948; Processed=1; Report=Applied=1; SkipDup=0; Poison=0; RunId=RUN-WH97-INVENTORY-20260325192957-684750; C:\Users\Justin\AppData\Local\Temp\invsys-phase6-wanhq-97a725a50c334b36a9d9cc331351ee49\Share\Snapshots\WH97.invSys.Snapshot.Inventory.xlsb |
| Publish.WH98.Initial | PASS | EventID=EVT-WH98-20260325192957349; Processed=1; Report=Applied=1; SkipDup=0; Poison=0; RunId=RUN-WH98-INVENTORY-20260325192958-362713; C:\Users\Justin\AppData\Local\Temp\invsys-phase6-wanhq-97a725a50c334b36a9d9cc331351ee49\Share\Snapshots\WH98.invSys.Snapshot.Inventory.xlsb |
| Aggregate.Initial | FAIL | QtyA=5; QtyB=8 |
| Publish.WH98.Catchup | PASS | EventID=EVT-WH98-20260325193003600; Processed=1; Report=Applied=1; SkipDup=0; Poison=0; RunId=RUN-WH98-INVENTORY-20260325193004-429191; C:\Users\Justin\AppData\Local\Temp\invsys-phase6-wanhq-97a725a50c334b36a9d9cc331351ee49\Share\Snapshots\WH98.invSys.Snapshot.Inventory.xlsb |
| Aggregate.Catchup | PASS | QtyA=5; QtyB=11 |
