# Phase 6 Live Role Workflow Validation Results

- Date: 2026-07-25 15:56:00
- Deploy root: C:\Users\justu\source\repos\invSys_fork\deploy\current
- Runtime root override: C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-1a20f31482114b879dd8a219f0d62ef1
- Passed: 44
- Failed: 0

| Check | Result | Detail |
|---|---|---|
| Core.RuntimeRootOverride | PASS | C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-1a20f31482114b879dd8a219f0d62ef1 |
| Core.AuthDiagnostic.User | PASS | ResolvedUser=justu; SeededUsers=justu,Justin Jahn,user1,svc_processor |
| Core.AuthDiagnostic.Config | PASS | WarehouseId=WHL0A93DB; StationId=S1; PathDataRoot=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-1a20f31482114b879dd8a219f0d62ef1 |
| Core.AuthDiagnostic.AuthLoad | PASS |  |
| Core.AuthDiagnostic.TargetSelect | PASS | OK/Connected - WHL0A93DB (Main Warehouse) at C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-1a20f31482114b879dd8a219f0d62ef1 |
| Core.AuthDiagnostic.SignIn | PASS | OK/User=justu/DisplayName=justu |
| Core.AuthDiagnostic.ReceiveCapability | PASS | User=justu; WarehouseId=WHL0A93DB; StationId=S1 |
| Core.AuthDiagnostic.ShipCapability | PASS | User=justu; WarehouseId=WHL0A93DB; StationId=S1 |
| Core.AuthDiagnostic.ProdCapability | PASS | User=justu; WarehouseId=WHL0A93DB; StationId=S1 |
| Core.ConfigBootstrap.CleanSurface | PASS | Load=True; Validate=WARN CONFIG_TABLE_CREATED: tblWarehouseConfig created.; WARN CONFIG_TABLE_CREATED: tblStationConfig created.; Sheets=2; WHTables=System.String[]; STTables=System.String[] |
| Core.RuntimeInventoryDiagnostic | PASS | Override=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-1a20f31482114b879dd8a219f0d62ef1; PathDataRoot=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-1a20f31482114b879dd8a219f0d62ef1; InventoryPath=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-1a20f31482114b879dd8a219f0d62ef1\WHL0A93DB.invSys.Data.Inventory.xlsb; FileExists=True; OpenFullName=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-1a20f31482114b879dd8a219f0d62ef1\WHL0A93DB.invSys.Data.Inventory.xlsb |
| Receiving.Form.Stage | PASS | StageResult=True; ReceivedTallyRows=1; AggregateReceivedRows=1 |
| Receiving.Capability.BeforeConfirm | PASS | Allowed=True/SignedIn=True/User=justu/Warehouse=WHL0A93DB/Station=S1/Auth=WHL0A93DB.invSys.Auth.xlsb/Error= |
| Receiving.ConfirmWrites.Local | PASS | Succeeded=True; Status=Confirm Writes succeeded.; CapabilityAfter=Allowed=True/SignedIn=True/User=justu/Warehouse=WHL0A93DB/Station=S1/Auth=WHL0A93DB.invSys.Auth.xlsb/Error=; ReceivedTallyRows=0; AggregateReceivedRows=0; RECEIVED=0; TOTAL_INV=7; QtyOnHand=; SourceType=LOCAL; IsStale=False; LogRows=1 |
| Receiving.ConfirmWrites.Queue | PASS | InboxRows=1; Row=1 |
| Receiving.ConfirmWrites.Process | PASS | StatusBeforeRun=PROCESSED; RunBatch=0; Status=PROCESSED; OutboxRow=0; ErrorCode=; ErrorMessage=; Processed=0; Report=Applied=0; SkipDup=0; Poison=0; RunId=RUN-WHL0A93DB-INVENTORY-20260725155512-631191; OpenBooks=WHL0A93DB.invSys.Auth.xlsb=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-1a20f31482114b879dd8a219f0d62ef1\WHL0A93DB.invSys.Auth.xlsb; WHL0A93DB.invSys.Data.Inventory.xlsb=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-1a20f31482114b879dd8a219f0d62ef1\WHL0A93DB.invSys.Data.Inventory.xlsb; invSys.Inbox.Receiving.S1.xlsb=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-1a20f31482114b879dd8a219f0d62ef1\invSys.Inbox.Receiving.S1.xlsb; invSys.Inbox.Shipping.S1.xlsb=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-1a20f31482114b879dd8a219f0d62ef1\invSys.Inbox.Shipping.S1.xlsb; invSys.Inbox.Production.S1.xlsb=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-1a20f31482114b879dd8a219f0d62ef1\invSys.Inbox.Production.S1.xlsb; WHL0A93DB.invSys.Config.xlsb=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-1a20f31482114b879dd8a219f0d62ef1\WHL0A93DB.invSys.Config.xlsb; WHL0A93DB.S1.Receiving.Operator.xlsb=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-1a20f31482114b879dd8a219f0d62ef1\WHL0A93DB.S1.Receiving.Operator.xlsb; WHL0A93DB.S1.Shipping.Operator.xlsb=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-1a20f31482114b879dd8a219f0d62ef1\WHL0A93DB.S1.Shipping.Operator.xlsb; WHL0A93DB.S1.Production.Operator.xlsb=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-1a20f31482114b879dd8a219f0d62ef1\WHL0A93DB.S1.Production.Operator.xlsb |
| Receiving.ConfirmWrites.InventoryLog | PASS | InventoryLogRowsBefore=3; Row=4; OutboxRow=0 |
| Shipping.Form.Stage | PASS | ShipmentRows=1; ShipROW=201; ShipQty=5; InvROW=201; InvCode=SKU-SHIP; InvTOTAL_INV=20 |
| Shipping.Capability.BeforeSent | PASS | Allowed=True/SignedIn=True/User=justu/Warehouse=WHL0A93DB/Station=S1/Auth=WHL0A93DB.invSys.Auth.xlsb/Error= |
| Shipping.Form.ShipmentsSent.Local | PASS | Report=OK/Shipments sent: 5. package(s).
Boxes sent:
- 5. Ship Widget vShip Widget
Carrier: UPS
Inbox EventID: A74F276C-CCE6-4E27-9D62-1C2195A12DC8s
Server inventory processed SHIP event: Processed=1; StagingReport=No local staged inbox rows.; BatchReport=Applied=1; SkipDup=0; Poison=0; RunId=RUN-WHL0A93DB-INVENTORY-20260725155518-611209; PublishWarning=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-1a20f31482114b879dd8a219f0d62ef1\WHL0A93DB.invSys.Snapshot.Inventory.xlsb; TimingMs=Total:5719;Batch:4180;Refresh:0; ShipmentRows=0; SHIPMENTS=0 |
| Shipping.BtnShipmentsSent.Queue | PASS | InboxRows=1; Row=1; Payload=[{"Row":201,"SKU":"SKU-SHIP","Qty":5,"Location":"","Note":"Ship Widget; VERSION=vShip Widget; REF=REF-SHIP-LIVE-001; CARRIER=UPS; ROW=201","Version":"vShip Widget","BomVersionLabel":"vShip Widget"}] |
| Shipping.BtnShipmentsSent.Process | PASS | StatusBeforeRun=PROCESSED; RunBatch=0; Status=PROCESSED; ErrorCode=; ErrorMessage=; Processed=0; Report=Applied=0; SkipDup=0; Poison=0; RunId=RUN-WHL0A93DB-INVENTORY-20260725155526-854113 |
| Shipping.BtnShipmentsSent.InventoryLog | PASS | InventoryLogRow=5 |
| Shipping.Hold.ToggleNotShipped | PASS | InitialHidden=False; AfterFirst=True; AfterSecond=False |
| Shipping.Hold.Send | PASS | Result=OK/Moved=4/SourceQty=6/TargetQty=4; ShipQty=6; HoldQty=4; HoldROW=250 |
| Shipping.Hold.Return | PASS | Result=OK/Moved=4/SourceQty=0/TargetQty=10; ShipQty=10; HoldQty=0 |
| Shipping.BtnBoxesMade.Local | PASS | ComponentUSED=0; ComponentTOTAL_INV=7; PackageMADE=2; AggregatePackagesRows=1 |
| Shipping.BtnToTotalInv.Local | PASS | PackageMADE=0; PackageTOTAL_INV=2 |
| Production.BtnSavePalette | PASS | PaletteRow=1; Before=ProdWb=WHL0A93DB.S1.Production.Operator.xlsb; RecipesSheet=Recipes; PaletteSheet=IngredientsPalette; RecipeId=R-001; IngredientId=ING-001; ChooseRecipeRows=1; ChooseIngredientRows=1; ChooseItemRows=1; FirstItem=Sugar Bin; PaletteRows=0; FirstPaletteRecipe=; After=ProdWb=WHL0A93DB.S1.Production.Operator.xlsb; RecipesSheet=Recipes; PaletteSheet=IngredientsPalette; RecipeId=R-001; IngredientId=ING-001; ChooseRecipeRows=1; ChooseIngredientRows=1; ChooseItemRows=1; FirstItem=Sugar Bin; PaletteRows=1; FirstPaletteRecipe=R-001 |
| Production.BtnPrintRecallCodes | PASS | Diag=OK; Sheet=RecallCodesPrint; Rows=1; RecallRows=1; RecallCode=RC-001 |
| Production.BtnToMade.Preflight | PASS | ProcessTables=RecipeChooser_generated:Rows=0,Process=,IO=,Ingredient=,Amount=; proc_1_rchooser:Rows=1,Process=Mix,IO=MADE,Ingredient=Finished Good,Amount=8; ProcessCheckboxes=0; OutputROW=401; RealOutput=8; InvRow2Code=SKU-FG |
| Production.Capability.BeforeComplete | PASS | Allowed=True/SignedIn=True/User=justu/Warehouse=WHL0A93DB/Station=S1/Auth=WHL0A93DB.invSys.Auth.xlsb/Error= |
| Production.Form.CheckIn | PASS | Report=OK/StagedUsed=2.; CheckedQty=2 |
| Production.Form.CheckIn.NoPrematureQueue | PASS | InboxBefore=0; InboxAfter=0 |
| Production.Form.CheckIn.OutputPreserved | PASS | RealOutput=8 |
| Production.Form.CheckIn.NoPrematureLog | PASS | InventoryLogRow=0 |
| Production.Form.CompleteRun.Local | PASS | Report=OK	ConsumeEvent=C8CDEACB-1588-4102-9FAD-0A5870FC3094; CompleteEvent=EC72CD2D-CFCD-42CB-B3A3-6AEF947A6C0D㾀; Processed=2; StagingReport=LocalStagingMerged=2; LocalStagingFailed=0; BatchReport=Applied=2; SkipDup=0; Poison=0; RunId=RUN-WHL0A93DB-INVENTORY-20260725155541-988399; PublishWarning=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-1a20f31482114b879dd8a219f0d62ef1\WHL0A93DB.invSys.Snapshot.Inventory.xlsb; RefreshReport=OK; TimingMs=Total:7504;Batch:5547;Surface:0;Refresh:434; MADE=0; TOTAL_INV=8; ProductionOutputRows=1; RealOutput=8 |
| Production.Form.CompleteRun.Queue | PASS | InboxRows=2; ConsumeRow=1; CompleteRow=2 |
| Production.Form.CompleteRun.Process | PASS | CompleteStatusBeforeRun=PROCESSED; RunBatch=0; CompleteStatus=PROCESSED; ConsumeStatus=PROCESSED; ErrorCode=; ErrorMessage=; Processed=0; Report=Applied=0; SkipDup=0; Poison=0; RunId=RUN-WHL0A93DB-INVENTORY-20260725155550-932201 |
| Production.Form.CompleteRun.InventoryLog | PASS | ConsumeLogRow=6; CompleteLogRow=7 |
| InventoryDomain.ProjectionRecovery.Delete | PASS | InventoryLogRows=7; AppliedRows=7 |
| InventoryDomain.ProjectionRecovery.RunBatch | PASS | Processed=1; Report=Applied=1; SkipDup=0; Poison=0; RunId=RUN-WHL0A93DB-INVENTORY-20260725155555-770028; Status=PROCESSED; ErrorCode=; ErrorMessage= |
| InventoryDomain.ProjectionRecovery.Balances | PASS | SKU-REC=8; SKU-SHIP=15; SKU-SUGAR=98; SKU-FG=8 |
| InventoryDomain.ProjectionRecovery.NonAuthoritative | PASS | InventoryLogRows=8; AppliedRows=8; RecoveryEvent=EVT-PROJECTION-RECOVERY-1E325630 |
