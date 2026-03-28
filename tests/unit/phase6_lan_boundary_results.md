# Phase 6 LAN Boundary Validation Results

- Date: 2026-03-28 12:34:01
- Passed: 10
- Failed: 0

| Check | Result |
|---|---|
| Setup.SharedRoot | PASS - OK|Root=C:\Users\Justin\repos\invSys_fork\tests\fixtures\phase6_lan_boundary_20260328_123348_302\runtime|Published=C:\Users\Justin\repos\invSys_fork\tests\fixtures\phase6_lan_boundary_20260328_123348_302\published |
| Attach.SessionA | PASS - OK|Warehouse=WH89|Station=S1 |
| Attach.SessionB | PASS - OK|Warehouse=WH89|Station=S2 |
| Lock.SessionAHold | PASS - OK|Path=C:\Users\Justin\repos\invSys_fork\tests\fixtures\phase6_lan_boundary_20260328_123348_302\runtime\WH89.invSys.Data.Inventory.xlsb|ReadOnly=False |
| Lock.SessionBDeniedByFileBoundary | PASS - OK|EventID=256345FA-C879-463E-9AB4-D06AF2D50D0Al|Processed=0|Report=Inventory workbook is read-only or locked by another Excel session.|Status=NEW|ErrorCode=|ErrorMessage= |
| Lock.SessionARelease | PASS - OK|Closed |
| Lock.SessionARetryAfterRelease | PASS - OK|EventID=4E587530-1DBB-46D7-839F-7B7882242A54l|Processed=1|Report=Applied=1; SkipDup=0; Poison=0; RunId=RUN-WH89-INVENTORY-20260328123359-021481|Status=PROCESSED|ErrorCode=|ErrorMessage= |
| Publish.SessionAToSharedSnapshot | PASS - OK|PublishedPath=C:\Users\Justin\repos\invSys_fork\tests\fixtures\phase6_lan_boundary_20260328_123348_302\published\WH89.invSys.Snapshot.Inventory.xlsb |
| Operator.BuildStationB | PASS - OK|OperatorPath=C:\Users\Justin\repos\invSys_fork\tests\fixtures\phase6_lan_boundary_20260328_123348_302\stationB_operator.xlsb |
| Refresh.SessionBReadsPublishedSnapshot | PASS - OK|TotalInv=7|QtyAvailable=7|SnapshotId=WH89.invSys.Snapshot.Inventory.xlsb/20260328123401|SourceType=SHAREPOINT|IsStale=False|Path=C:\Users\Justin\repos\invSys_fork\tests\fixtures\phase6_lan_boundary_20260328_123348_302\stationB_operator.xlsb |
