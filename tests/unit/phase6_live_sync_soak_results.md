# Phase 6 Bidirectional Live Sync Soak Results

- Date: 2026-03-28 21:27:05
- AutoRefreshIntervalSeconds: 1
- AverageBatchMs: 1136
- AverageCatchupMs: 3642
- AverageCatchupMs_S1toS2: 5340
- AverageCatchupMs_S2toS1: 1945
- CanonicalRoot: C:\Users\Justin\repos\invSys_fork\tests\fixtures\phase6_lan_boundary_20260328_212622_004\runtime
- IterationsRequested: 1
- LegsExecuted: 2
- LegsFailed: 0
- LegsPassed: 2
- MaxCatchupMs: 5340
- OperatorWorkbook: C:\Users\Justin\repos\invSys_fork\tests\fixtures\phase6_lan_boundary_20260328_212622_004\WH89_S2_Receiving_Operator.xlsb
- PollIntervalMs: 250
- PollTimeoutSeconds: 20
- SessionRoot: C:\Users\Justin\repos\invSys_fork\tests\fixtures\phase6_lan_boundary_20260328_212622_004
- SourceWorkbook: C:\Users\Justin\repos\invSys_fork\tests\fixtures\phase6_lan_boundary_20260328_212622_004\WH89_FRODECO.inventory_management.xlsb
- SyncLogPath: C:\Users\Justin\AppData\Local\Temp\invSys.Inventory.Sync.log

| Iteration | Direction | Qty | ExpectedTotal | Processed | BatchMs | CatchupMs | RuntimeReadOnly | ObservedTotal | ObservedReceived | Result | Detail |
|---|---|---:|---:|---:|---:|---:|---|---:|---:|---|---|
| 1 | S2->S1 | 1 | 8 | 1 | 1234 | 1945 | True | 8 | 1 | PASS | 2026-03-28 21:26:52 / CANARY / SchedulerFired=2026-03-28 21:26:52 // 2026-03-28 21:26:52 / DETECTION / OpenWbs=1/WH89_FRODECO.inventory_management.xlsb=True; // 2026-03-28 21:26:52 / TRACE / SrcWb=WH89_FRODECO.inventory_ |
| 1 | S1->S2 | 1 | 9 | 1 | 1037 | 5340 |  | 9 | 1 | PASS | Processed=1; Report=Applied=1; SkipDup=0; Poison=0; RunId=RUN-WH89-INVENTORY-20260328212659-890846 |
