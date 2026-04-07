# Phase 6 Packaged Ribbon Validation Results

- Date: 2026-04-06 13:37:35
- Deploy root: C:\Users\Justin\repos\invSys_fork\deploy\current
- Runtime root override: C:\Users\Justin\AppData\Local\Temp\invsys-phase6-ribbon-10977d5c3021455dbaec230ecc1b9469
- Passed: 38
- Failed: 44

| Check | Result | Detail |
|---|---|---|
| invSys.Core.xlam.Open | PASS | Opened from C:\Users\Justin\repos\invSys_fork\deploy\current\invSys.Core.xlam |
| invSys.Inventory.Domain.xlam.Open | PASS | Opened from C:\Users\Justin\repos\invSys_fork\deploy\current\invSys.Inventory.Domain.xlam |
| invSys.Designs.Domain.xlam.Open | PASS | Opened from C:\Users\Justin\repos\invSys_fork\deploy\current\invSys.Designs.Domain.xlam |
| invSys.Receiving.xlam.Open | PASS | Opened from C:\Users\Justin\repos\invSys_fork\deploy\current\invSys.Receiving.xlam |
| invSys.Shipping.xlam.Open | PASS | Opened from C:\Users\Justin\repos\invSys_fork\deploy\current\invSys.Shipping.xlam |
| invSys.Production.xlam.Open | PASS | Opened from C:\Users\Justin\repos\invSys_fork\deploy\current\invSys.Production.xlam |
| invSys.Admin.xlam.Open | PASS | Opened from C:\Users\Justin\repos\invSys_fork\deploy\current\invSys.Admin.xlam |
| Core.RuntimeRootOverride | PASS | C:\Users\Justin\AppData\Local\Temp\invsys-phase6-ribbon-10977d5c3021455dbaec230ecc1b9469 |
| Receiving.RibbonXml | PASS | customUI/customUI.xml present. |
| Receiving.CallbackModule | PASS | modRibbonGenerated |
| Receiving.RibbonButton.btnReceivingSetup | PASS | Label=Setup UI; OnAction=RibbonOnActionReceiving; Screentip= |
| Receiving.MacroExists.btnReceivingSetup | PASS | modTS_Received.EnsureGeneratedButtons |
| Receiving.CallbackMap.btnReceivingSetup | PASS | btnReceivingSetup -> modTS_Received.EnsureGeneratedButtons |
| Receiving.SafeExec.btnReceivingSetup | FAIL | Exception calling "Run" with "1" argument(s): "The remote procedure call failed. (Exception from HRESULT: 0x800706BE)" |
| Receiving.RibbonButton.btnReceivingConfirm | PASS | Label=Confirm Writes; OnAction=RibbonOnActionReceiving; Screentip= |
| Receiving.MacroExists.btnReceivingConfirm | FAIL | modTS_Received.ConfirmWrites |
| Receiving.CallbackMap.btnReceivingConfirm | PASS | btnReceivingConfirm -> modTS_Received.ConfirmWrites |
| Receiving.RibbonButton.btnReceivingUndo | PASS | Label=Undo; OnAction=RibbonOnActionReceiving; Screentip= |
| Receiving.MacroExists.btnReceivingUndo | FAIL | modTS_Received.MacroUndo |
| Receiving.CallbackMap.btnReceivingUndo | PASS | btnReceivingUndo -> modTS_Received.MacroUndo |
| Receiving.RibbonButton.btnReceivingRedo | PASS | Label=Redo; OnAction=RibbonOnActionReceiving; Screentip= |
| Receiving.MacroExists.btnReceivingRedo | FAIL | modTS_Received.MacroRedo |
| Receiving.CallbackMap.btnReceivingRedo | PASS | btnReceivingRedo -> modTS_Received.MacroRedo |
| Shipping.RibbonXml | PASS | customUI/customUI.xml present. |
| Shipping.CallbackModule | FAIL | modRibbonGenerated |
| Shipping.RibbonButton.btnShippingSetup | PASS | Label=Setup UI; OnAction=RibbonOnActionShipping; Screentip= |
| Shipping.MacroExists.btnShippingSetup | FAIL | modTS_Shipments.InitializeShipmentsUI |
| Shipping.CallbackMap.btnShippingSetup | FAIL | btnShippingSetup -> modTS_Shipments.InitializeShipmentsUI |
| Shipping.SafeExec.btnShippingSetup | FAIL | Exception calling "Run" with "1" argument(s): "The RPC server is unavailable. (Exception from HRESULT: 0x800706BA)" |
| Shipping.RibbonButton.btnShippingConfirm | PASS | Label=Confirm Inventory; OnAction=RibbonOnActionShipping; Screentip= |
| Shipping.MacroExists.btnShippingConfirm | FAIL | modTS_Shipments.BtnConfirmInventory |
| Shipping.CallbackMap.btnShippingConfirm | FAIL | btnShippingConfirm -> modTS_Shipments.BtnConfirmInventory |
| Shipping.RibbonButton.btnShippingStage | PASS | Label=To Shipments; OnAction=RibbonOnActionShipping; Screentip= |
| Shipping.MacroExists.btnShippingStage | FAIL | modTS_Shipments.BtnToShipments |
| Shipping.CallbackMap.btnShippingStage | FAIL | btnShippingStage -> modTS_Shipments.BtnToShipments |
| Shipping.RibbonButton.btnShippingSend | PASS | Label=Shipments Sent; OnAction=RibbonOnActionShipping; Screentip= |
| Shipping.MacroExists.btnShippingSend | FAIL | modTS_Shipments.BtnShipmentsSent |
| Shipping.CallbackMap.btnShippingSend | FAIL | btnShippingSend -> modTS_Shipments.BtnShipmentsSent |
| Production.RibbonXml | PASS | customUI/customUI.xml present. |
| Production.CallbackModule | FAIL | modRibbonGenerated |
| Production.RibbonButton.btnProductionSetup | PASS | Label=Setup UI; OnAction=RibbonOnActionProduction; Screentip= |
| Production.MacroExists.btnProductionSetup | FAIL | mProduction.InitializeProductionUI |
| Production.CallbackMap.btnProductionSetup | FAIL | btnProductionSetup -> mProduction.InitializeProductionUI |
| Production.SafeExec.btnProductionSetup | FAIL | Exception calling "Run" with "1" argument(s): "The RPC server is unavailable. (Exception from HRESULT: 0x800706BA)" |
| Production.RibbonButton.btnProductionLoad | PASS | Label=Load Recipe; OnAction=RibbonOnActionProduction; Screentip= |
| Production.MacroExists.btnProductionLoad | FAIL | mProduction.BtnLoadRecipe |
| Production.CallbackMap.btnProductionLoad | FAIL | btnProductionLoad -> mProduction.BtnLoadRecipe |
| Production.RibbonButton.btnProductionUsed | PASS | Label=To Used; OnAction=RibbonOnActionProduction; Screentip= |
| Production.MacroExists.btnProductionUsed | FAIL | mProduction.BtnToUsed |
| Production.CallbackMap.btnProductionUsed | FAIL | btnProductionUsed -> mProduction.BtnToUsed |
| Production.RibbonButton.btnProductionMade | PASS | Label=To Made; OnAction=RibbonOnActionProduction; Screentip= |
| Production.MacroExists.btnProductionMade | FAIL | mProduction.BtnToMade |
| Production.CallbackMap.btnProductionMade | FAIL | btnProductionMade -> mProduction.BtnToMade |
| Production.RibbonButton.btnProductionTotal | PASS | Label=To Total Inv; OnAction=RibbonOnActionProduction; Screentip= |
| Production.MacroExists.btnProductionTotal | FAIL | mProduction.BtnToTotalInv |
| Production.CallbackMap.btnProductionTotal | FAIL | btnProductionTotal -> mProduction.BtnToTotalInv |
| Production.RibbonButton.btnProductionPrintCodes | PASS | Label=Print Recall Codes; OnAction=RibbonOnActionProduction; Screentip= |
| Production.MacroExists.btnProductionPrintCodes | FAIL | mProduction.BtnPrintRecallCodes |
| Production.CallbackMap.btnProductionPrintCodes | FAIL | btnProductionPrintCodes -> mProduction.BtnPrintRecallCodes |
| Admin.RibbonXml | PASS | customUI/customUI.xml present. |
| Admin.CallbackModule | FAIL | modRibbonGenerated |
| Admin.RibbonButton.btnAdminOpen | PASS | Label=Admin Console; OnAction=RibbonOnActionAdmin; Screentip= |
| Admin.MacroExists.btnAdminOpen | FAIL | modAdmin.Admin_Click |
| Admin.CallbackMap.btnAdminOpen | FAIL | btnAdminOpen -> modAdmin.Admin_Click |
| Admin.SafeExec.btnAdminOpen | FAIL | Exception calling "Run" with "1" argument(s): "The RPC server is unavailable. (Exception from HRESULT: 0x800706BA)" |
| Admin.RibbonButton.btnAdminUsers | PASS | Label=Users and Roles; OnAction=RibbonOnActionAdmin; Screentip= |
| Admin.MacroExists.btnAdminUsers | FAIL | modAdmin.Open_CreateDeleteUser |
| Admin.CallbackMap.btnAdminUsers | FAIL | btnAdminUsers -> modAdmin.Open_CreateDeleteUser |
| Admin.SafeExec.btnAdminUsers | FAIL | Exception calling "Run" with "1" argument(s): "The RPC server is unavailable. (Exception from HRESULT: 0x800706BA)" |
| Admin.RibbonButton.btnAdminCreateWarehouse | PASS | Label=Create New Warehouse; OnAction=RibbonOnActionAdmin; Screentip= |
| Admin.MacroExists.btnAdminCreateWarehouse | FAIL | modAdmin.Open_CreateWarehouse |
| Admin.CallbackMap.btnAdminCreateWarehouse | FAIL | btnAdminCreateWarehouse -> modAdmin.Open_CreateWarehouse |
| Admin.RibbonButton.btnAdminSetupTesterStation | FAIL | Button missing from Ribbon XML. |
| Admin.MacroExists.btnAdminSetupTesterStation | FAIL | modAdmin.Admin_SetupTesterStation_Click |
| Admin.CallbackMap.btnAdminSetupTesterStation | FAIL | btnAdminSetupTesterStation -> modAdmin.Admin_SetupTesterStation_Click |
| Admin.RibbonButton.btnAdminVerifyAddinsPublished | PASS | Label=Verify Add-ins Published; OnAction=RibbonOnActionAdmin; Screentip= |
| Admin.MacroExists.btnAdminVerifyAddinsPublished | FAIL | modAdmin.Verify_AddinsPublished |
| Admin.CallbackMap.btnAdminVerifyAddinsPublished | FAIL | btnAdminVerifyAddinsPublished -> modAdmin.Verify_AddinsPublished |
| Admin.RibbonButton.btnAdminRetireMigrateWarehouse | PASS | Label=Retire / Migrate Warehouse; OnAction=RibbonOnActionAdmin; Screentip=Archive, migrate, retire, or delete a warehouse runtime |
| Admin.RibbonButtonScreentip.btnAdminRetireMigrateWarehouse | PASS | Archive, migrate, retire, or delete a warehouse runtime |
| Admin.MacroExists.btnAdminRetireMigrateWarehouse | FAIL | modAdmin.Admin_RetireMigrateWarehouse_Click |
| Admin.CallbackMap.btnAdminRetireMigrateWarehouse | FAIL | btnAdminRetireMigrateWarehouse -> modAdmin.Admin_RetireMigrateWarehouse_Click |
