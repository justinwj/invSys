# Create Warehouse Integration Results

- Date: 2026-07-27 14:35:13
- Overall: PASS
- Harness: tests/fixtures/<generated-create-warehouse-harness>.xlsm
- Warehouse: WHBOOT-E2E_01
- Station: ADM1
- Local root: <temporary-test-root>
- SharePoint root: <temporary-test-root>
- Summary: Create warehouse lifecycle completed, SharePoint artifacts were published, and duplicate rejection was proven.
- Passed checks: 15
- Failed checks: 0

| Check | Result | Detail |
|---|---|---|
| WarehouseSpec.Valid | PASS | OK |
| CollisionCheck.InitialClear | PASS | WarehouseIdExists=False |
| Bootstrap.Local | PASS | OK ; Hub=<redacted-path> ; Inbox=<redacted-path> ; Seed=SEEDED ; Operator=<redacted-path> |
| LocalStructure.Exists | PASS | All required runtime folders and seeded artifacts were created under <redacted-path> |
| D14.GeneratedSchemas.ManagedHeadersNoROW | PASS | Inventory, snapshot, and operator tables contain required managed headers and no ROW header. |
| D14.Seed.UniqueKeysConditionGood | PASS | tblInventoryEntities contains 3 unique nonblank System_Key values with Condition=GOOD. |
| D14.RoundTrip.PreservesSystemKey | PASS | Inventory entity keys must survive processor application, snapshot publication, and operator refresh. |
| D14.AdminSeed.RepeatedCreatesUniqueKeysNoMigration | PASS | Repeated Admin seed created six new collision-free System_Key values with Condition=GOOD and no migration event. |
| D14.OperatorRefresh.PreservesCustomColumn | PASS | Custom_Local_Note survived refresh for System_Key 8C536538-D98E-42D0-9E54-27BE0941010D. |
| D14.OperatorReopen.PreservesCustomColumn | PASS | Custom_Local_Note remained associated with its System_Key after save and reopen. |
| ConfigSeeded.Correctly | PASS | Config workbook seeded WarehouseId, WarehouseName, StationId, PathDataRoot, PathSharePointRoot, and RECEIVE defaults. |
| SharePointPublish.Initial | PASS | OK ; Config=COPIED:<redacted-path> ; Discovery=COPIED:<redacted-path> |
| SharePointArtifacts.Exists | PASS | Discovery artifact and published config workbook exist under <redacted-path> |
| CollisionCheck.DuplicateVisible | PASS | WarehouseIdExists=True |
| DuplicateRun.Rejected | PASS | WarehouseId already exists in the configured warehouse catalog: WHBOOT-E2E_01 |
