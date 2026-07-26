# Phase 6 Packaged WAN HQ Validation Results

- Date: 2026-07-25 15:57:13
- Deploy root: C:\Users\justu\source\repos\invSys_fork\deploy\current
- Session root: C:\Users\justu\AppData\Local\Temp\invsys-phase6-wanhq-6e6aa25e73b24eb58dc717d08a4ce183
- Passed: 10
- Failed: 0

| Check | Result | Detail |
|---|---|---|
| Setup.RuntimeRoots | PASS | SessionRoot=C:\Users\justu\AppData\Local\Temp\invsys-phase6-wanhq-6e6aa25e73b24eb58dc717d08a4ce183 |
| Packaged.OpenA | PASS | Core+Inventory.Domain |
| Packaged.OpenB | PASS | Core+Inventory.Domain |
| Packaged.OpenHQ | PASS | Core+Inventory.Domain |
| Packaged.RuntimeOverrides | PASS | WH97=C:\Users\justu\AppData\Local\Temp\invsys-phase6-wanhq-6e6aa25e73b24eb58dc717d08a4ce183\WH97; WH98=C:\Users\justu\AppData\Local\Temp\invsys-phase6-wanhq-6e6aa25e73b24eb58dc717d08a4ce183\WH98 |
| Publish.WH97.Initial | PASS | EventID=EVT-WH97-20260725155653186; Processed=1; Report=Applied=1; SkipDup=0; Poison=0; RunId=RUN-WH97-INVENTORY-20260725155656-631436; C:\Users\justu\AppData\Local\Temp\invsys-phase6-wanhq-6e6aa25e73b24eb58dc717d08a4ce183\Share\Snapshots\WH97.invSys.Snapshot.Inventory.xlsb |
| Publish.WH98.Initial | PASS | EventID=EVT-WH98-20260725155653770; Processed=1; Report=Applied=1; SkipDup=0; Poison=0; RunId=RUN-WH98-INVENTORY-20260725155700-826779; C:\Users\justu\AppData\Local\Temp\invsys-phase6-wanhq-6e6aa25e73b24eb58dc717d08a4ce183\Share\Snapshots\WH98.invSys.Snapshot.Inventory.xlsb |
| Aggregate.Initial | PASS | QtyA=5; QtyB=8 |
| Publish.WH98.Catchup | PASS | EventID=EVT-WH98-20260725155708459; Processed=1; Report=Applied=1; SkipDup=0; Poison=0; RunId=RUN-WH98-INVENTORY-20260725155709-274604; C:\Users\justu\AppData\Local\Temp\invsys-phase6-wanhq-6e6aa25e73b24eb58dc717d08a4ce183\Share\Snapshots\WH98.invSys.Snapshot.Inventory.xlsb |
| Aggregate.Catchup | PASS | QtyA=5; QtyB=11 |
