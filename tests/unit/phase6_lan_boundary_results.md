# Phase 6 LAN Boundary Validation Results

- Date: 2026-03-28 21:26:32
- Passed: 9
- Failed: 1

| Check | Result |
|---|---|
| Setup.SharedRoot | FAIL - ERR|That name is already taken. Try a different one. |
| Attach.SessionA | PASS - OK|Warehouse=WH89|Station=S1 |
| Attach.SessionB | PASS - OK|Warehouse=WH89|Station=S2 |
| Lock.SessionAHold | PASS - OK|Path=C:\Users\Justin\repos\invSys_fork\tests\fixtures\phase6_lan_boundary_20260328_212622_004\runtime\WH89.invSys.Data.Inventory.xlsb|ReadOnly=False |
| Lock.SessionBDeniedByFileBoundary | PASS - OK|EventID=9ABFCB92-1905-45BD-AD03-A84EBFE6234Dl|Processed=0|Report=Inventory workbook is read-only or locked by another Excel session.|Status=NEW|ErrorCode=|ErrorMessage= |
| Lock.SessionARelease | PASS - OK|Closed |
| Lock.SessionARetryAfterRelease | PASS - OK|EventID=275EEB06-830E-4CA0-9571-762357634749l|Processed=1|Report=Applied=1; SkipDup=0; Poison=0; RunId=RUN-WH89-INVENTORY-20260328212631-292965|Status=PROCESSED|ErrorCode=|ErrorMessage= |
| Publish.SessionAToSharedSnapshot | PASS - OK|PublishedPath=C:\Users\Justin\repos\invSys_fork\tests\fixtures\phase6_lan_boundary_20260328_212622_004\published\WH89.invSys.Snapshot.Inventory.xlsb |
| Operator.BuildStationB | PASS - OK|OperatorPath=C:\Users\Justin\repos\invSys_fork\tests\fixtures\phase6_lan_boundary_20260328_212622_004\stationB_operator.xlsb |
| Refresh.SessionBReadsPublishedSnapshot | PASS - OK|TotalInv=7|QtyAvailable=7|SnapshotId=WH89.invSys.Snapshot.Inventory.xlsb/20260328212632|SourceType=SHAREPOINT|IsStale=False|Path=C:\Users\Justin\repos\invSys_fork\tests\fixtures\phase6_lan_boundary_20260328_212622_004\stationB_operator.xlsb |
