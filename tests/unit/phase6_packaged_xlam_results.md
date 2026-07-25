# Phase 6 Packaged XLAM Validation Results

- Date: 2026-07-25 12:58:07
- Deploy root: C:\Users\justu\source\repos\invSys_fork\deploy\current
- Passed: 43
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
| Receiving.Init | PASS | modReceivingInit.InitReceivingAddin |
| Receiving.SafeMacro | PASS | modTS_Received.EnsureGeneratedButtons |
| Receiving.Surface | PASS | OK |
| Shipping.frmShipmentsTally.Code | PASS | OK |
| Shipping.Init | PASS | modShippingInit.InitShippingAddin |
| Shipping.SafeMacro | PASS | modTS_Shipments.InitializeShipmentsUI |
| Shipping.Surface | PASS | OK |
| Production.Init | PASS | modProductionInit.InitProductionAddin |
| Production.SafeMacro | PASS | mProduction.InitializeProductionUI |
| Production.FormInitialize | PASS | OK/Pages=4/Status=Production form loaded for WH1.Production.Operator.xlsx. Inventory: ContractVersion=R1-INVENTORY-1/Workbook=invSys.Inventory.Domain.xlam/IsAddin=True/StartupOperatorMutation=False/LegacyDirectWrites=False/UndoModel=CompensatingEvent/Authority=WHx.invSys.Data.Inventory.xlsb Designs: legacy recipe fallback (disabled in warehouse config). |
| Production.Surface | PASS | OK |
| Admin.Init | PASS | modAdminInit.InitAdminAddin |
| Admin.FormInitialize | PASS | OK/Rows=27/Workbook=WH1.invSys.Config.xlsb |
| Admin.Surface | PASS | OK |
| InventoryDomain.PeerAutoLoad | PASS | ContractVersion=R1-INVENTORY-1/Workbook=invSys.Inventory.Domain.xlam/IsAddin=True/StartupOperatorMutation=False/LegacyDirectWrites=False/UndoModel=CompensatingEvent/Authority=WHx.invSys.Data.Inventory.xlsb |
| DesignsDomain.PeerAutoLoad | PASS | ContractVersion=R1-DESIGNS-1/Workbook=invSys.Designs.Domain.xlam/IsAddin=True/StartupMutation=False/Authority=WHx.invSys.Data.Designs.xlsb; WorkbookOpen=False |
