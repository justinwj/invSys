# Phase 6 Packaged XLAM Validation Results

- Date: 2026-07-26 00:11:33
- Deploy root: C:\Users\justu\source\repos\invSys_fork\deploy\current
- Passed: 59
- Failed: 0

| Check | Result | Detail |
|---|---|---|
| invSys.Core.xlam.Open | PASS | Opened from C:\Users\justu\source\repos\invSys_fork\deploy\current\invSys.Core.xlam |
| invSys.Core.xlam.IsAddin | PASS | IsAddin=True |
| invSys.Inventory.Domain.xlam.Open | PASS | Opened from C:\Users\justu\source\repos\invSys_fork\deploy\current\invSys.Inventory.Domain.xlam |
| invSys.Inventory.Domain.xlam.IsAddin | PASS | IsAddin=True |
| invSys.Designs.Domain.xlam.Open | PASS | Opened from C:\Users\justu\source\repos\invSys_fork\deploy\current\invSys.Designs.Domain.xlam |
| invSys.Designs.Domain.xlam.IsAddin | PASS | IsAddin=True |
| DomainStartup.OperatorIsolation | PASS | Core, Inventory Domain, and Designs Domain left the active operator sentinel unchanged. |
| invSys.Receiving.xlam.Open | PASS | Opened from C:\Users\justu\source\repos\invSys_fork\deploy\current\invSys.Receiving.xlam |
| invSys.Receiving.xlam.IsAddin | PASS | IsAddin=True |
| invSys.Shipping.xlam.Open | PASS | Opened from C:\Users\justu\source\repos\invSys_fork\deploy\current\invSys.Shipping.xlam |
| invSys.Shipping.xlam.IsAddin | PASS | IsAddin=True |
| invSys.Production.xlam.Open | PASS | Opened from C:\Users\justu\source\repos\invSys_fork\deploy\current\invSys.Production.xlam |
| invSys.Production.xlam.IsAddin | PASS | IsAddin=True |
| invSys.Admin.xlam.Open | PASS | Opened from C:\Users\justu\source\repos\invSys_fork\deploy\current\invSys.Admin.xlam |
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
| Production.FormInitialize | PASS | OK/Pages=4/WindowStyle=Handle=True/Resizable=True/Minimize=True/Maximize=True/Status=Production form loaded for WH1.Production.Operator.xlsx. Inventory: ContractVersion=R1-INVENTORY-1/Workbook=invSys.Inventory.Domain.xlam/IsAddin=True/StartupOperatorMutation=False/LegacyDirectWrites=False/UndoModel=CompensatingEvent/Authority=WHx.invSys.Data.Inventory.xlsb Designs: ContractVersion=R1-DESIGNS-1/Workbook=invSys.Designs.Domain.xlam/IsAddin=True/StartupMutation=False/Authority=WHx.invSys.Data.Designs.xlsb |
| Production.Surface | PASS | OK |
| Admin.Init | PASS | modAdminInit.InitAdminAddin |
| Admin.FormInitialize | PASS | OK/Rows=28/Workbook=invsys_Zenbook_WH.invSys.Config.xlsb/ManualServerCredentials=FALSE/Uoms=12 |
| Admin.Surface | PASS | OK |
| Admin.PoisonReissue.PackagedSurface | PASS | FAIL/Report=Source workbook not open: __PACKAGED_SMOKE_MISSING__.xlsb |
| Admin.DesignLifecycle.LegacyMigrationControl | PASS | LayoutReady=1 |
| InventoryDomain.PeerAutoLoad | PASS | ContractVersion=R1-INVENTORY-1/Workbook=invSys.Inventory.Domain.xlam/IsAddin=True/StartupOperatorMutation=False/LegacyDirectWrites=False/UndoModel=CompensatingEvent/Authority=WHx.invSys.Data.Inventory.xlsb |
| DesignsDomain.PeerAutoLoad | PASS | ContractVersion=R1-DESIGNS-1/Workbook=invSys.Designs.Domain.xlam/IsAddin=True/StartupMutation=False/Authority=WHx.invSys.Data.Designs.xlsb; WorkbookOpen=False |
| Restart.invSys.Core.xlam | PASS | IsAddin=True; FullName=C:\Users\justu\source\repos\invSys_fork\deploy\current\invSys.Core.xlam |
| Restart.invSys.Inventory.Domain.xlam | PASS | IsAddin=True; FullName=C:\Users\justu\source\repos\invSys_fork\deploy\current\invSys.Inventory.Domain.xlam |
| Restart.invSys.Designs.Domain.xlam | PASS | IsAddin=True; FullName=C:\Users\justu\source\repos\invSys_fork\deploy\current\invSys.Designs.Domain.xlam |
| Restart.invSys.Receiving.xlam | PASS | IsAddin=True; FullName=C:\Users\justu\source\repos\invSys_fork\deploy\current\invSys.Receiving.xlam |
| Restart.invSys.Shipping.xlam | PASS | IsAddin=True; FullName=C:\Users\justu\source\repos\invSys_fork\deploy\current\invSys.Shipping.xlam |
| Restart.invSys.Production.xlam | PASS | IsAddin=True; FullName=C:\Users\justu\source\repos\invSys_fork\deploy\current\invSys.Production.xlam |
| Restart.invSys.Admin.xlam | PASS | IsAddin=True; FullName=C:\Users\justu\source\repos\invSys_fork\deploy\current\invSys.Admin.xlam |
| Restart.Receiving.SavedWorkbook | PASS | FullName=C:\Users\justu\AppData\Local\Temp\invsys-packaged-surfaces-cc59cfefa3074d4fb638ea7d6a414f70\WH1.Receiving.Operator.xlsx; Surface=OK |
| Restart.Shipping.SavedWorkbook | PASS | FullName=C:\Users\justu\AppData\Local\Temp\invsys-packaged-surfaces-cc59cfefa3074d4fb638ea7d6a414f70\WH1.Shipping.Operator.xlsx; Surface=OK |
| Restart.Production.SavedWorkbook | PASS | FullName=C:\Users\justu\AppData\Local\Temp\invsys-packaged-surfaces-cc59cfefa3074d4fb638ea7d6a414f70\WH1.Production.Operator.xlsx; Surface=OK |
| Restart.Admin.SavedWorkbook | PASS | FullName=C:\Users\justu\AppData\Local\Temp\invsys-packaged-surfaces-cc59cfefa3074d4fb638ea7d6a414f70\WH1.Admin.Console.xlsx; Surface=OK |
| Restart.DomainBridges | PASS | Inventory=ContractVersion=R1-INVENTORY-1/Workbook=invSys.Inventory.Domain.xlam/IsAddin=True/StartupOperatorMutation=False/LegacyDirectWrites=False/UndoModel=CompensatingEvent/Authority=WHx.invSys.Data.Inventory.xlsb; Designs=ContractVersion=R1-DESIGNS-1/Workbook=invSys.Designs.Domain.xlam/IsAddin=True/StartupMutation=False/Authority=WHx.invSys.Data.Designs.xlsb |
