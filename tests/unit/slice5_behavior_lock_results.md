# Slice 5 Packaged Behavior Lock Results

- Passed: 8
- Failed: 5

| Check | Result | Contract |
|---|---|---|
| Receiving.FormAction.ConfirmWrites.Handler | PASS | The form button must call the operator Confirm Writes handler. |
| Production.FormActions.RequiredHandlers | PASS | Selection/Apply, Check In, Complete Run, and Next Batch must remain wired to the operator handlers. |
| Shipping.FormActions.RequiredHandlers | PASS | To Shipments and Shipments Sent must remain wired to the operator handlers. |
| Receiving.Form.ModelessLauncher | PASS | Receiving launcher must open the main form modelessly. |
| Production.Form.ModelessLauncher | FAIL | Production launcher must open the main form modelessly. |
| Shipping.Form.ModelessLauncher | FAIL | Shipping launcher must open the main form modelessly. |
| Receiving.Navigation.PurchasingStub | FAIL | Receiving must expose a selectable, visibly non-operational Purchasing tab. |
| Shipping.Navigation.SingleTabbedShell | FAIL | The main Shipping form must contain Box Builder and Box Maker tabs. |
| Operations.Package.Exists | PASS | The build map must define invSys.Operations.xlam. |
| Operations.Ribbon.SingleShippingLauncher | FAIL | Operations Ribbon must expose one Shipping launcher and no separate Box Builder or Box Maker buttons. |
| Receiving.Form.CapturedWorkbookState | PASS | Receiving form must retain explicit operator-workbook state. |
| Production.Form.CapturedWorkbookState | PASS | Production form must retain explicit operator-workbook state. |
| Shipping.Form.CapturedWorkbookState | PASS | Shipping form must retain explicit operator-workbook state. |
