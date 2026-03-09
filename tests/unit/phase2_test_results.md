# Phase 2 VBA Test Results

- Date: 2026-03-08 22:20:08
- Passed: 21
- Failed: 0

| Test | Result |
|---|---|
| TestCoreConfig.TestLoad_ValidConfig | PASS |
| TestCoreConfig.TestLoad_MissingRequiredKey | PASS |
| TestCoreConfig.TestPrecedence_StationOverridesWarehouse | PASS |
| TestCoreConfig.TestGetRequired_MissingKey | PASS |
| TestCoreConfig.TestGetBool_TypeConversion | PASS |
| TestCoreConfig.TestReload_UpdatedValue | PASS |
| TestCoreAuth.TestCanPerform_Allow | PASS |
| TestCoreAuth.TestCanPerform_Deny_MissingCapability | PASS |
| TestCoreAuth.TestCanPerform_WildcardStation | PASS |
| TestCoreAuth.TestCanPerform_DisabledUser | PASS |
| TestCoreAuth.TestCanPerform_ExpiredCapability | PASS |
| TestCoreAuth.TestRequire_RaisesOnDeny | PASS |
| TestInventorySchema.TestEnsureInventorySchema_RecreatesTables | PASS |
| TestInventorySchema.TestEnsureInventorySchema_AddsMissingColumns | PASS |
| TestCoreLockManager.TestAcquireReleaseLock_Lifecycle | PASS |
| TestCoreLockManager.TestHeartbeat_ExtendsExpiry | PASS |
| TestInventoryApply.TestApplyReceive_ValidEvent | PASS |
| TestInventoryApply.TestApplyReceive_InvalidSKU | PASS |
| TestInventoryApply.TestApplyReceive_Duplicate | PASS |
| TestCoreProcessor.TestRunBatch_ProcessesInboxRow | PASS |
| TestCoreProcessor.TestRunBatch_DuplicateMarkedSkipDup | PASS |
