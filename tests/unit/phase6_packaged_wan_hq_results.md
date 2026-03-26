# Phase 6 Packaged WAN HQ Validation Results

- Date: 2026-03-25 19:39:29
- Deploy root: C:\Users\Justin\repos\invSys_fork\deploy\current
- Session root: C:\Users\Justin\AppData\Local\Temp\invsys-phase6-wanhq-2ad11126d8a14e66acc661b18d62414a
- Passed: 10
- Failed: 0

| Check | Result | Detail |
|---|---|---|
| Setup.RuntimeRoots | PASS | SessionRoot=C:\Users\Justin\AppData\Local\Temp\invsys-phase6-wanhq-2ad11126d8a14e66acc661b18d62414a |
| Packaged.OpenA | PASS | Core+Inventory.Domain |
| Packaged.OpenB | PASS | Core+Inventory.Domain |
| Packaged.OpenHQ | PASS | Core+Inventory.Domain |
| Packaged.RuntimeOverrides | PASS | WH97=C:\Users\Justin\AppData\Local\Temp\invsys-phase6-wanhq-2ad11126d8a14e66acc661b18d62414a\WH97; WH98=C:\Users\Justin\AppData\Local\Temp\invsys-phase6-wanhq-2ad11126d8a14e66acc661b18d62414a\WH98 |
| Publish.WH97.Initial | PASS | EventID=EVT-WH97-20260325193918723; Processed=1; Report=Applied=1; SkipDup=0; Poison=0; RunId=RUN-WH97-INVENTORY-20260325193919-510159; C:\Users\Justin\AppData\Local\Temp\invsys-phase6-wanhq-2ad11126d8a14e66acc661b18d62414a\Share\Snapshots\WH97.invSys.Snapshot.Inventory.xlsb |
| Publish.WH98.Initial | PASS | EventID=EVT-WH98-20260325193919121; Processed=1; Report=Applied=1; SkipDup=0; Poison=0; RunId=RUN-WH98-INVENTORY-20260325193920-453701; C:\Users\Justin\AppData\Local\Temp\invsys-phase6-wanhq-2ad11126d8a14e66acc661b18d62414a\Share\Snapshots\WH98.invSys.Snapshot.Inventory.xlsb |
| Aggregate.Initial | PASS | QtyA=5; QtyB=8 |
| Publish.WH98.Catchup | PASS | EventID=EVT-WH98-20260325193928145; Processed=1; Report=Applied=1; SkipDup=0; Poison=0; RunId=RUN-WH98-INVENTORY-20260325193928-512107; C:\Users\Justin\AppData\Local\Temp\invsys-phase6-wanhq-2ad11126d8a14e66acc661b18d62414a\Share\Snapshots\WH98.invSys.Snapshot.Inventory.xlsb |
| Aggregate.Catchup | PASS | QtyA=5; QtyB=11 |
