# Phase 6 Live Role Workflow Validation Results

- Date: 2026-03-22 12:49:48
- Deploy root: C:\Users\Justin\repos\invSys_fork\deploy\current
- Runtime root override: C:\Users\Justin\AppData\Local\Temp\invsys-phase6-live-ced45a601ab64f43b89c3f190d9ecd38
- Passed: 10
- Failed: 1

| Check | Result | Detail |
|---|---|---|
| Core.RuntimeRootOverride | PASS | C:\Users\Justin\AppData\Local\Temp\invsys-phase6-live-ced45a601ab64f43b89c3f190d9ecd38 |
| Core.AuthDiagnostic.User | PASS | ResolvedUser=Justin; SeededUsers=Justin,user1,svc_processor |
| Core.AuthDiagnostic.Config | PASS | WarehouseId=WH1; StationId=S1; PathDataRoot=C:\Users\Justin\AppData\Local\Temp\invsys-phase6-live-ced45a601ab64f43b89c3f190d9ecd38 |
| Core.AuthDiagnostic.AuthLoad | PASS |  |
| Core.AuthDiagnostic.ReceiveCapability | PASS | User=Justin; WarehouseId=WH1; StationId=S1 |
| Core.AuthDiagnostic.ShipCapability | PASS | User=Justin; WarehouseId=WH1; StationId=S1 |
| Core.AuthDiagnostic.ProdCapability | PASS | User=Justin; WarehouseId=WH1; StationId=S1 |
| Receiving.ConfirmWrites.QueueDiagnostic | PASS | OK |
| Receiving.ConfirmWrites.Local | PASS | RECEIVED=7; LogRows=1 |
| Receiving.ConfirmWrites.Queue | PASS | InboxRows=3; Row=2 |
| Harness.Exception | FAIL | Step=Run Receiving ConfirmWrites; Exception calling "Run" with "3" argument(s): "Exception from HRESULT: 0x800A9C68" |
