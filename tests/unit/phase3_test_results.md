# Phase 3 VBA Test Results

- Date: 2026-03-09 01:13:51
- Passed: 12
- Failed: 0

| Test | Result |
|---|---|
| TestCoreRoleEventWriter.TestQueueReceiveEvent_WritesInboxRow | PASS |
| TestCoreRoleEventWriter.TestQueueShipEvent_WritesInboxRow | PASS |
| TestCoreRoleEventWriter.TestQueuePayloadEvent_DeniedWithoutCapability | PASS |
| TestCoreRoleEventWriter.TestBuildPayloadJson_WithObjectItems | PASS |
| TestCoreRoleUiAccess.TestCanCurrentUserPerformCapability_Allow | PASS |
| TestCoreRoleUiAccess.TestCanCurrentUserPerformCapability_Deny | PASS |
| TestCoreRoleUiAccess.TestApplyShapeCapability_TogglesVisibility | PASS |
| TestCoreItemSearch.TestNormalizeSearchText_CollapsesWhitespace | PASS |
| TestCoreItemSearch.TestAnyTextMatchesSearch_MatchesAcrossFields | PASS |
| TestCoreItemSearch.TestIdentifiersMatch_UsesTokenOverlap | PASS |
| TestCoreItemSearch.TestResolveSearchCaption_ReturnsRoleSpecificText | PASS |
| TestCoreItemSearch.TestShouldDefaultShippableForRole_UsesRoleDefaults | PASS |
