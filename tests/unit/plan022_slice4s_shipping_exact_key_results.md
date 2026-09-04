# Plan 022 Slice 4s Shipping Exact-Key Results

- Passed: 6
- Failed: 0

| Check | Result | Contract |
|---|---|---|
| Shipping.Add.PublicAction | PASS | The real Shipping Add callback reaches the public ShipmentsFormCommitLine action and its local reserve boundary. |
| Shipping.Add.SystemKeyApply | PASS | Shipping Add reserves current-schema inventory by immutable System_Key when managed ROW is absent. |
| BoxDesigner.ActiveExactEntities | PASS | Box Designer excludes nonpositive balances and removes only duplicate projections of the same exact System_Key. |
| BoxDesigner.PreservesDistinctIdentity | PASS | Positive entities with different System_Key values remain separate selectable component identities. |
| SavingUi.ActionBoundaries | PASS | Shipping Add, Box Designer save, and Box Maker post retain one nested quiet UI boundary across required persistence. |
| SavingUi.StatusBarRestored | PASS | Quiet UI hides Excel save-status churn and restores the operator's previous status-bar setting. |
