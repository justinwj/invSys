# Phase 6 Packaged XLAM Validation Results

- Date: 2026-07-27 15:43:38
- Deploy root: deploy/current
- Passed: 59
- Failed: 0

| Check | Result | Detail |
|---|---|---|
| invSys.Core.xlam.Open | PASS | Opened from <redacted-path> |
| invSys.Core.xlam.IsAddin | PASS | IsAddin=True |
| invSys.Inventory.Domain.xlam.Open | PASS | Opened from <redacted-path> |
| invSys.Inventory.Domain.xlam.IsAddin | PASS | IsAddin=True |
| invSys.Designs.Domain.xlam.Open | PASS | Opened from <redacted-path> |
| invSys.Designs.Domain.xlam.IsAddin | PASS | IsAddin=True |
| DomainStartup.OperatorIsolation | PASS | Core, Inventory Domain, and Designs Domain left the active operator sentinel unchanged. |
| invSys.Receiving.xlam.Open | PASS | Opened from <redacted-path> |
| invSys.Receiving.xlam.IsAddin | PASS | IsAddin=True |
| invSys.Shipping.xlam.Open | PASS | Opened from <redacted-path> |
| invSys.Shipping.xlam.IsAddin | PASS | IsAddin=True |
| invSys.Production.xlam.Open | PASS | Opened from <redacted-path> |
| invSys.Production.xlam.IsAddin | PASS | IsAddin=True |
| invSys.Admin.xlam.Open | PASS | Opened from <redacted-path> |
| invSys.Admin.xlam.IsAddin | PASS | IsAddin=True |
| XLAMStartup.ExplicitAutoOpenOperatorIsolation | PASS | All Domain/Role/Admin Auto_Open entry points left the active operator workbook unchanged. |
| RoleEvents.WorkbookOpenOperatorIsolation | PASS | WorkbookOpen/NewWorkbook handlers left an unrelated operator workbook unchanged. |
| RoleEvents.NamedSheetOperatorIsolation | PASS | Selection/change handlers ignored unrelated workbooks whose sheet names resembled role sheets but lacked role-owned tables. |
| invSys.Core.xlam.modInventoryDomainBridge | PASS | OK |
| invSys.Core.xlam.modDesignsDomainBridge | PASS | OK |
| invSys.Inventory.Domain.xlam.modInventoryApply | PASS | OK |
| invSys.Inventory.Domain.xlam.modInventoryQueries | PASS | OK |
| invSys.Inventory.Domain.xlam.modInvMan | PASS | OK |
| invSys.Inventory.Domain.xlam.cInventoryAppEvents | PASS | OK |
| invSys.Designs.Domain.xlam.modDesignsApply | PASS | OK |
| invSys.Designs.Domain.xlam.modDesignsQueries | PASS | OK |
| invSys.Designs.Domain.xlam.modDesignsSchema | PASS | OK |
| invSys.Admin.xlam.modAdminConsole | PASS | OK |
| invSys.Admin.xlam.modAdminDesignLifecycle | PASS | OK |
| Receiving.Init | PASS | modReceivingInit.InitReceivingAddin |
| Receiving.SafeMacro | PASS | modTS_Received.EnsureGeneratedButtons |
| Receiving.Surface | PASS | OK |
| Shipping.frmShipmentsTally.Code | PASS | OK |
| Shipping.Init | PASS | modShippingInit.InitShippingAddin |
| Shipping.SafeMacro | PASS | modTS_Shipments.InitializeShipmentsUI |
| Shipping.Surface | PASS | OK |
| Production.Init | PASS | modProductionInit.InitProductionAddin |
| Production.SafeMacro | PASS | mProduction.InitializeProductionUI |
| Production.FormInitialize | PASS | OK/Pages=4/WindowStyle=Handle=True/Resizable=True/Minimize=True/Maximize=True/Status=Production form loaded for WH1.Production.Operator.xlsx. Inventory: ContractVersion=R1-INVENTORY-1/Workbook=invSys.Inventory.Domain.xlam/IsAddin=True/StartupOperatorMutation=False/LegacyDirectWrites=False/UndoModel=CompensatingEvent/Authority=WHx.invSys.Data.Inventory.xlsb Designs: legacy recipe fallback (disabled in warehouse config). |
| Production.Surface | PASS | OK |
| Admin.Init | PASS | modAdminInit.InitAdminAddin |
| Admin.FormInitialize | PASS | OK/Rows=28/Workbook=WHL4BF29D.invSys.Config.xlsb/ManualServerCredentials=FALSE/Uoms=12 |
| Admin.Surface | PASS | OK |
| Admin.PoisonReissue.PackagedSurface | PASS | FAIL/Report=Source workbook not open: __PACKAGED_SMOKE_MISSING__.xlsb |
| Admin.DesignLifecycle.LegacyMigrationControl | PASS | LayoutReady=1 |
| InventoryDomain.PeerAutoLoad | PASS | ContractVersion=R1-INVENTORY-1/Workbook=invSys.Inventory.Domain.xlam/IsAddin=True/StartupOperatorMutation=False/LegacyDirectWrites=False/UndoModel=CompensatingEvent/Authority=WHx.invSys.Data.Inventory.xlsb |
| DesignsDomain.PeerAutoLoad | PASS | ContractVersion=R1-DESIGNS-1/Workbook=invSys.Designs.Domain.xlam/IsAddin=True/StartupMutation=False/Authority=WHx.invSys.Data.Designs.xlsb; WorkbookOpen=False |
| Restart.invSys.Core.xlam | PASS | IsAddin=True; FullName=<redacted-path> |
| Restart.invSys.Inventory.Domain.xlam | PASS | IsAddin=True; FullName=<redacted-path> |
| Restart.invSys.Designs.Domain.xlam | PASS | IsAddin=True; FullName=<redacted-path> |
| Restart.invSys.Receiving.xlam | PASS | IsAddin=True; FullName=<redacted-path> |
| Restart.invSys.Shipping.xlam | PASS | IsAddin=True; FullName=<redacted-path> |
| Restart.invSys.Production.xlam | PASS | IsAddin=True; FullName=<redacted-path> |
| Restart.invSys.Admin.xlam | PASS | IsAddin=True; FullName=<redacted-path> |
| Restart.Receiving.SavedWorkbook | PASS | FullName=<redacted-path>; Surface=OK |
| Restart.Shipping.SavedWorkbook | PASS | FullName=<redacted-path>; Surface=OK |
| Restart.Production.SavedWorkbook | PASS | FullName=<redacted-path>; Surface=OK |
| Restart.Admin.SavedWorkbook | PASS | FullName=<redacted-path>; Surface=OK |
| Restart.DomainBridges | PASS | Inventory=ContractVersion=R1-INVENTORY-1/Workbook=invSys.Inventory.Domain.xlam/IsAddin=True/StartupOperatorMutation=False/LegacyDirectWrites=False/UndoModel=CompensatingEvent/Authority=WHx.invSys.Data.Inventory.xlsb; Designs=ContractVersion=R1-DESIGNS-1/Workbook=invSys.Designs.Domain.xlam/IsAddin=True/StartupMutation=False/Authority=WHx.invSys.Data.Designs.xlsb |
