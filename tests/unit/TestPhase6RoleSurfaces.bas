Attribute VB_Name = "TestPhase6RoleSurfaces"
Option Explicit

Private mLastTestFailure As String

Public Sub ClearLastTestFailure()
    mLastTestFailure = vbNullString
End Sub

Public Function GetLastTestFailure() As String
    GetLastTestFailure = mLastTestFailure
End Function

Public Function TestEnsureReceivingWorkbookSurface_CreatesExpectedTables() As Long
    Dim wb As Workbook
    Dim report As String

    Set wb = Application.Workbooks.Add

    On Error GoTo CleanFail
    If Not modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface(wb, report) Then GoTo CleanExit
    If HasTable(wb, "ReceivedTally") _
       And HasTable(wb, "AggregateReceived") _
       And HasTable(wb, "invSysData_Receiving") _
       And HasTable(wb, "ReceivedLog") _
       And HasTable(wb, "invSys") _
       And TableHasColumns(wb, "ReceivedTally", Array("REF_NUMBER", "ITEMS", "QUANTITY", "ROW")) _
       And TableHasColumns(wb, "AggregateReceived", Array("REF_NUMBER", "ITEM_CODE", "VENDORS", "VENDOR_CODE", "DESCRIPTION", "ITEM", "UOM", "QUANTITY", "LOCATION", "ROW")) _
       And TableHasColumns(wb, "invSysData_Receiving", Array("ROW", "ITEM_CODE", "ITEM", "UOM", "LOCATION", "DESCRIPTION")) _
       And TableHasColumns(wb, "ReceivedLog", Array("SNAPSHOT_ID", "ENTRY_DATE", "REF_NUMBER", "ITEMS", "QUANTITY", "UOM", "VENDOR", "LOCATION", "ITEM_CODE", "ROW")) _
       And TableHasColumns(wb, "invSys", Array("ROW", "ITEM_CODE", "ITEM", "UOM", "LOCATION", "DESCRIPTION", "TOTAL INV", "QtyAvailable", "LocationSummary", "LastRefreshUTC", "SnapshotId", "SourceType", "IsStale")) Then
        TestEnsureReceivingWorkbookSurface_CreatesExpectedTables = 1
    End If

CleanExit:
    CloseNoSavePhase6 wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestEnsureInventoryManagementSurface_RemovesDuplicateAliasColumns() As Long
    Dim wb As Workbook
    Dim report As String
    Dim lo As ListObject

    Set wb = Application.Workbooks.Add

    On Error GoTo CleanFail
    If Not modRoleWorkbookSurfaces.EnsureInventoryManagementSurface(wb, report) Then GoTo CleanExit
    If Not HasTable(wb, "invSys") Then GoTo CleanExit
    Set lo = wb.Worksheets("InventoryManagement").ListObjects("invSys")

    lo.ListColumns.Add.Name = "SKU"
    lo.ListColumns.Add.Name = "ItemName"
    lo.ListColumns.Add.Name = "QtyOnHand"
    lo.ListColumns.Add.Name = "LastAppliedUTC"
    lo.ListColumns.Add.Name = "TIMESTAMP"

    If Not modRoleWorkbookSurfaces.EnsureInventoryManagementSurface(wb, report) Then GoTo CleanExit

    If TableColumnHidden(wb, "invSys", "ROW") _
       And TableColumnHidden(wb, "invSys", "TOTAL INV LAST EDIT") _
       And Not TableColumnHidden(wb, "invSys", "ITEM_CODE") _
       And Not TableColumnHidden(wb, "invSys", "TOTAL INV") _
       And Not TableColumnHidden(wb, "invSys", "QtyAvailable") _
       And Not TableColumnHidden(wb, "invSys", "LocationSummary") _
       And Not TableColumnHidden(wb, "invSys", "LastRefreshUTC") _
       And Not TableColumnHidden(wb, "invSys", "SnapshotId") _
       And Not TableColumnHidden(wb, "invSys", "SourceType") _
       And Not TableColumnHidden(wb, "invSys", "IsStale") _
       And Not TableHasColumns(wb, "invSys", Array("SKU", "ItemName", "QtyOnHand", "LastAppliedUTC", "TIMESTAMP")) Then
        TestEnsureInventoryManagementSurface_RemovesDuplicateAliasColumns = 1
    End If

CleanExit:
    CloseNoSavePhase6 wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestEnsureShippingWorkbookSurface_CreatesExpectedTables() As Long
    Dim wb As Workbook
    Dim report As String

    Set wb = Application.Workbooks.Add

    On Error GoTo CleanFail
    If Not modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface(wb, report) Then GoTo CleanExit
    If HasTable(wb, "ShipmentsTally") _
       And HasTable(wb, "BoxBuilder") _
       And HasTable(wb, "BoxBOM") _
       And HasTable(wb, "AggregatePackages") _
       And HasTable(wb, "invSysData_Shipping") _
       And HasTable(wb, "AggregateBoxBOM_Log") _
       And HasTable(wb, "AggregatePackages_Log") _
       And HasTable(wb, "Check_invSys") _
       And HasTable(wb, "invSys") _
       And HasTable(wb, "ShippingBOMView") _
       And Not WorksheetExists(wb, "ShippingBOM") _
       And TableHasColumns(wb, "ShipmentsTally", Array("LINE_ID", "SERVER_RESERVE_EVENT_ID", "REF_NUMBER", "ITEMS", "QUANTITY", "System_Key", "UOM", "LOCATION", "DESCRIPTION")) _
       And TableHasColumns(wb, "BoxBuilder", Array("Box Name", "UOM", "LOCATION", "DESCRIPTION")) _
       And Not TableHasColumns(wb, "BoxBuilder", Array("ROW")) _
       And TableHasColumns(wb, "BoxBOM", Array("ITEM", "ROW", "QUANTITY", "UOM", "LOCATION", "DESCRIPTION")) _
       And TableHasColumns(wb, "AggregatePackages", Array("ROW", "ITEM_CODE", "ITEM", "QUANTITY", "UOM", "LOCATION")) _
       And TableHasColumns(wb, "invSysData_Shipping", Array("ROW", "ITEM_CODE", "ITEM", "UOM", "LOCATION", "DESCRIPTION")) _
       And TableHasColumns(wb, "ShippingBOMView", Array("PackageRow", "PackageItem", "ComponentRow", "ComponentQty", "UpdatedAtUTC", "UpdatedBy")) _
       And TableHasColumns(wb, "AggregateBoxBOM_Log", Array("GUID", "USER", "ACTION", "ROW", "ITEM_CODE", "ITEM", "QTY_DELTA", "NEW_VALUE", "TIMESTAMP")) _
       And TableHasColumns(wb, "AggregatePackages_Log", Array("GUID", "USER", "ACTION", "System_Key", "ITEM_CODE", "ITEM", "QTY_DELTA", "NEW_VALUE", "TIMESTAMP")) Then
        TestEnsureShippingWorkbookSurface_CreatesExpectedTables = 1
    End If

CleanExit:
    CloseNoSavePhase6 wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestEnsureReceivingWorkbookSurface_RecreatesDeletedArtifacts() As Long
    Dim wb As Workbook
    Dim report As String

    Set wb = Application.Workbooks.Add

    On Error GoTo CleanFail
    If Not modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface(wb, report) Then GoTo CleanExit

    DeleteTablePhase6 wb, "AggregateReceived"
    DeleteTablePhase6 wb, "invSys"
    DeleteWorksheetPhase6 wb, "ReceivedLog"

    If Not modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface(wb, report) Then GoTo CleanExit
    If HasTable(wb, "AggregateReceived") _
       And HasTable(wb, "invSys") _
       And HasTable(wb, "ReceivedLog") _
       And WorksheetExists(wb, "ReceivedLog") Then
        TestEnsureReceivingWorkbookSurface_RecreatesDeletedArtifacts = 1
    End If

CleanExit:
    CloseNoSavePhase6 wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestReceivingForm_SearchFiltersInventoryAndKeepsRefExternal() As Long
    Dim values(1 To 3, 1 To 8) As Variant
    Dim matchCount As Long

    On Error GoTo CleanFail
    values(1, 1) = 95
    values(1, 2) = "ITM-TEA-001"
    values(1, 3) = "Malawi Black Tea"
    values(1, 4) = "LB"
    values(1, 5) = 100
    values(1, 6) = "CLEARVIEW"
    values(1, 7) = "Fine cut black tea"
    values(1, 8) = "Henry"
    values(2, 1) = 96
    values(2, 2) = "ITM-WTR-001"
    values(2, 3) = "Filtered Water"
    values(2, 4) = "GAL"
    values(2, 5) = 0
    values(2, 6) = "A1"
    values(2, 7) = "Utility water"
    values(2, 8) = "Utility"
    values(3, 1) = 97
    values(3, 2) = "ITM-TEA-002"
    values(3, 3) = "Assam Black Tea"
    values(3, 4) = "LB"
    values(3, 5) = 40
    values(3, 6) = "NAS-A1"
    values(3, 7) = "Coarse leaf tea"
    values(3, 8) = "Vendor"

    matchCount = frmReceiving.TestSearchInventoryCount(values, "black tea")
    If matchCount = 2 _
       And frmReceiving.TestSearchInventoryCount(values, "clearview") = 1 _
       And frmReceiving.TestReceiptIdSeparatedFromReference() = 1 Then
        TestReceivingForm_SearchFiltersInventoryAndKeepsRefExternal = 1
    End If

CleanExit:
    On Error Resume Next
    Unload frmReceiving
    On Error GoTo 0
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestReceivingForm_InventoryLoaderUsesRawRowValue() As Long
    Dim wb As Workbook
    Dim report As String
    Dim loInv As ListObject
    Dim lr As ListRow
    Dim rowsOut As Variant
    Dim rowCol As Long

    Set wb = Application.Workbooks.Add

    On Error GoTo CleanFail
    If Not modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface(wb, report) Then GoTo CleanExit
    Set loInv = FindTable(wb, "invSys")
    If loInv Is Nothing Then GoTo CleanExit

    If loInv.DataBodyRange Is Nothing Then
        Set lr = loInv.ListRows.Add
    Else
        Set lr = loInv.ListRows(1)
    End If
    SetTableValueByColumn loInv, lr.Index, "ROW", 14
    SetTableValueByColumn loInv, lr.Index, "ITEM_CODE", "ITM-ROW-014"
    SetTableValueByColumn loInv, lr.Index, "ITEM", "Row Format Test"
    SetTableValueByColumn loInv, lr.Index, "UOM", "EA"
    SetTableValueByColumn loInv, lr.Index, "LOCATION", "A1"
    SetTableValueByColumn loInv, lr.Index, "DESCRIPTION", "Date formatted row"
    SetTableValueByColumn loInv, lr.Index, "TOTAL INV", 5
    SetTableValueByColumn loInv, lr.Index, "QtyAvailable", 5

    rowCol = TableColumnIndex(loInv, "ROW")
    If rowCol = 0 Then GoTo CleanExit
    loInv.ListColumns(rowCol).DataBodyRange.NumberFormat = "m/d/yyyy"
    wb.Activate

    rowsOut = modTS_Received.LoadReceivingFormInventory("")
    If IsEmpty(rowsOut) Or Not IsArray(rowsOut) Then GoTo CleanExit
    If CStr(rowsOut(1, 1)) = "14" Then
        TestReceivingForm_InventoryLoaderUsesRawRowValue = 1
    End If

CleanExit:
    CloseNoSavePhase6 wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestReceivingForm_HidesSupportSheetsAfterFormRefresh() As Long
    Dim wb As Workbook
    Dim report As String
    Dim wsOperator As Worksheet
    Dim wsRt As Worksheet
    Dim wsInv As Worksheet
    Dim wsLog As Worksheet

    Set wb = Application.Workbooks.Add(xlWBATWorksheet)

    On Error GoTo CleanFail
    Set wsOperator = wb.Worksheets(1)
    wsOperator.Name = "OperatorVisible"
    If Not modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface(wb, report) Then GoTo CleanExit
    Set wsRt = wb.Worksheets("ReceivedTally")
    Set wsInv = wb.Worksheets("InventoryManagement")
    Set wsLog = wb.Worksheets("ReceivedLog")
    wsRt.Visible = xlSheetVisible
    wsInv.Visible = xlSheetVisible
    wsLog.Visible = xlSheetVisible

    modTS_Received.EnforceReceivingSupportSheetsHidden wb
    If wsRt.Visible <> xlSheetVeryHidden Then GoTo CleanExit
    If wsInv.Visible <> xlSheetVeryHidden Then GoTo CleanExit
    If wsLog.Visible <> xlSheetVeryHidden Then GoTo CleanExit
    If wsOperator.Visible <> xlSheetVisible Then GoTo CleanExit

    TestReceivingForm_HidesSupportSheetsAfterFormRefresh = 1

CleanExit:
    CloseNoSavePhase6 wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestReceivingForm_AddStagesSelectedInventoryForConfirm() As Long
    Dim wb As Workbook
    Dim report As String
    Dim loInv As ListObject
    Dim loRt As ListObject
    Dim loAgg As ListObject
    Dim lr As ListRow
    Dim stageReport As String

    Set wb = Application.Workbooks.Add

    On Error GoTo CleanFail
    If Not modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface(wb, report) Then GoTo CleanExit
    Set loInv = FindTable(wb, "invSys")
    Set loRt = FindTable(wb, "ReceivedTally")
    Set loAgg = FindTable(wb, "AggregateReceived")
    If loInv Is Nothing Or loRt Is Nothing Or loAgg Is Nothing Then GoTo CleanExit

    If loInv.DataBodyRange Is Nothing Then
        Set lr = loInv.ListRows.Add
    Else
        Set lr = loInv.ListRows(1)
    End If
    SetTableValueByColumn loInv, lr.Index, "ROW", 701
    SetTableValueByColumn loInv, lr.Index, "ITEM_CODE", "ITM-RECV-701"
    SetTableValueByColumn loInv, lr.Index, "ITEM", "Received Test Tea"
    SetTableValueByColumn loInv, lr.Index, "UOM", "LB"
    SetTableValueByColumn loInv, lr.Index, "LOCATION", "DOCK"
    SetTableValueByColumn loInv, lr.Index, "DESCRIPTION", "Receiving form add test"
    SetTableValueByColumn loInv, lr.Index, "VENDOR(s)", "Test Vendor"
    SetTableValueByColumn loInv, lr.Index, "VENDOR_CODE", "TV-701"
    SetTableValueByColumn loInv, lr.Index, "TOTAL INV", 0
    SetTableValueByColumn loInv, lr.Index, "QtyAvailable", 0

    If Not modTS_Received.StageReceivingFormLineForWorkbook(wb, "PO-701", 701, 12, stageReport) Then GoTo CleanExit
    If loRt.DataBodyRange Is Nothing Or loAgg.DataBodyRange Is Nothing Then GoTo CleanExit

    If CStr(GetTableValueByColumn(loRt, 1, "REF_NUMBER")) = "PO-701" _
       And CStr(GetTableValueByColumn(loRt, 1, "ITEMS")) = "Received Test Tea" _
       And CDbl(GetTableValueByColumn(loRt, 1, "QUANTITY")) = 12 _
       And CLng(GetTableValueByColumn(loRt, 1, "ROW")) = 701 _
       And CStr(GetTableValueByColumn(loAgg, 1, "ITEM_CODE")) = "ITM-RECV-701" _
       And CStr(GetTableValueByColumn(loAgg, 1, "ITEM")) = "Received Test Tea" _
       And CDbl(GetTableValueByColumn(loAgg, 1, "QUANTITY")) = 12 _
       And CLng(GetTableValueByColumn(loAgg, 1, "ROW")) = 701 Then
        TestReceivingForm_AddStagesSelectedInventoryForConfirm = 1
    End If

CleanExit:
    CloseNoSavePhase6 wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestReceivingForm_AddMergesSameRefItemAndSeparatesDifferentRef() As Long
    Dim wb As Workbook
    Dim report As String
    Dim loInv As ListObject
    Dim loRt As ListObject
    Dim loAgg As ListObject
    Dim lr As ListRow
    Dim stageReport As String

    Set wb = Application.Workbooks.Add

    On Error GoTo CleanFail
    If Not modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface(wb, report) Then GoTo CleanExit
    Set loInv = FindTable(wb, "invSys")
    Set loRt = FindTable(wb, "ReceivedTally")
    Set loAgg = FindTable(wb, "AggregateReceived")
    If loInv Is Nothing Or loRt Is Nothing Or loAgg Is Nothing Then GoTo CleanExit

    If loInv.DataBodyRange Is Nothing Then
        Set lr = loInv.ListRows.Add
    Else
        Set lr = loInv.ListRows(1)
    End If
    SetTableValueByColumn loInv, lr.Index, "ROW", 706
    SetTableValueByColumn loInv, lr.Index, "ITEM_CODE", "ITM-RECV-706"
    SetTableValueByColumn loInv, lr.Index, "ITEM", "Cardamom Pods"
    SetTableValueByColumn loInv, lr.Index, "UOM", "LB"
    SetTableValueByColumn loInv, lr.Index, "LOCATION", "DOCK"
    SetTableValueByColumn loInv, lr.Index, "DESCRIPTION", "Receiving merge test"
    SetTableValueByColumn loInv, lr.Index, "VENDOR(s)", "Test Vendor"
    SetTableValueByColumn loInv, lr.Index, "VENDOR_CODE", "TV-706"

    If Not modTS_Received.StageReceivingFormLineForWorkbook(wb, "PO-706", 706, 500, stageReport) Then GoTo CleanExit
    If Not modTS_Received.StageReceivingFormLineForWorkbook(wb, "PO-706", 706, 100, stageReport) Then GoTo CleanExit
    If Not modTS_Received.StageReceivingFormLineForWorkbook(wb, "PO-707", 706, 25, stageReport) Then GoTo CleanExit
    If loRt.DataBodyRange Is Nothing Or loAgg.DataBodyRange Is Nothing Then GoTo CleanExit

    If loRt.ListRows.Count >= 2 _
       And CStr(GetTableValueByColumn(loRt, 1, "REF_NUMBER")) = "PO-706" _
       And CDbl(GetTableValueByColumn(loRt, 1, "QUANTITY")) = 600 _
       And CStr(GetTableValueByColumn(loRt, 2, "REF_NUMBER")) = "PO-707" _
       And CDbl(GetTableValueByColumn(loRt, 2, "QUANTITY")) = 25 _
       And CDbl(GetTableValueByColumn(loAgg, 1, "QUANTITY")) = 625 _
       And CStr(GetTableValueByColumn(loAgg, 1, "REF_NUMBER")) = "PO-706,PO-707" Then
        TestReceivingForm_AddMergesSameRefItemAndSeparatesDifferentRef = 1
    End If

CleanExit:
    CloseNoSavePhase6 wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestReceivingForm_AddStagesByItemCodeWhenRowsCollide() As Long
    Dim wb As Workbook
    Dim report As String
    Dim loInv As ListObject
    Dim loAgg As ListObject
    Dim lr As ListRow
    Dim stageReport As String

    Set wb = Application.Workbooks.Add

    On Error GoTo CleanFail
    If Not modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface(wb, report) Then GoTo CleanExit
    Set loInv = FindTable(wb, "invSys")
    Set loAgg = FindTable(wb, "AggregateReceived")
    If loInv Is Nothing Or loAgg Is Nothing Then GoTo CleanExit

    Set lr = loInv.ListRows.Add
    SetTableValueByColumn loInv, lr.Index, "ROW", 1
    SetTableValueByColumn loInv, lr.Index, "ITEM_CODE", "DEMO-RAW-BLACK-TEA"
    SetTableValueByColumn loInv, lr.Index, "ITEM", "Demo Black Tea"
    SetTableValueByColumn loInv, lr.Index, "UOM", "LB"
    SetTableValueByColumn loInv, lr.Index, "LOCATION", "NAS-A1"

    Set lr = loInv.ListRows.Add
    SetTableValueByColumn loInv, lr.Index, "ROW", 1
    SetTableValueByColumn loInv, lr.Index, "ITEM_CODE", "ITEM-0002"
    SetTableValueByColumn loInv, lr.Index, "ITEM", "Black Tea Base"
    SetTableValueByColumn loInv, lr.Index, "UOM", "LB"
    SetTableValueByColumn loInv, lr.Index, "LOCATION", "CLEARVIEW"

    If Not modTS_Received.StageReceivingFormItemForWorkbook(wb, "BOL-2", 1, "ITEM-0002", 7, stageReport) Then GoTo CleanExit
    If loAgg.DataBodyRange Is Nothing Then GoTo CleanExit

    If CStr(GetTableValueByColumn(loAgg, 1, "ITEM_CODE")) = "ITEM-0002" _
       And CStr(GetTableValueByColumn(loAgg, 1, "ITEM")) = "Black Tea Base" _
       And CDbl(GetTableValueByColumn(loAgg, 1, "QUANTITY")) = 7 Then
        TestReceivingForm_AddStagesByItemCodeWhenRowsCollide = 1
    End If

CleanExit:
    CloseNoSavePhase6 wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestEnsureProductionWorkbookSurface_CreatesExpectedTables() As Long
    Dim wb As Workbook
    Dim report As String

    Set wb = Application.Workbooks.Add

    On Error GoTo CleanFail
    If Not modRoleWorkbookSurfaces.EnsureProductionWorkbookSurface(wb, report) Then GoTo CleanExit
    If HasTable(wb, "RB_AddRecipeName") _
       And HasTable(wb, "RecipeBuilder") _
       And HasTable(wb, "IP_ChooseRecipe") _
       And HasTable(wb, "IP_ChooseIngredient") _
       And HasTable(wb, "IP_ChooseItem") _
       And HasTable(wb, "RC_RecipeChoose") _
       And HasTable(wb, "ProductionOutput") _
       And HasTable(wb, "Prod_invSys_Check") _
       And HasTable(wb, "Recipes") _
       And HasTable(wb, "IngredientPalette") _
       And HasTable(wb, "TemplatesTable") _
       And HasTable(wb, "ProductionLog") _
       And HasTable(wb, "BatchCodesLog") _
       And HasTable(wb, "invSys") _
       And WorksheetExistsAny(wb, Array("IngredientPalette", "IngredientsPalette")) _
       And TableHasColumns(wb, "IP_ChooseRecipe", Array("RECIPE_NAME", "DESCRIPTION", "GUID", "RECIPE_ID")) _
       And TableHasColumns(wb, "IP_ChooseIngredient", Array("INGREDIENT", "UOM", "QUANTITY", "DESCRIPTION", "GUID", "RECIPE_ID", "INGREDIENT_ID", "PROCESS")) _
       And TableHasColumns(wb, "IP_ChooseItem", Array("ITEMS", "UOM", "DESCRIPTION", "System_Key", "RECIPE_ID", "INGREDIENT_ID")) _
       And TableHasColumns(wb, "InventoryPalette_generated", Array("INGREDIENT", "INGREDIENT_ID", "ITEM", "SPLIT %", "QUANTITY", "BASE QUANTITY", "PROCESS", "System_Key")) _
       And TableHasColumns(wb, "IngredientPalette", Array("RECIPE_ID", "INGREDIENT_ID", "INPUT/OUTPUT", "ITEM", "PERCENT", "UOM", "AMOUNT", "System_Key", "GUID")) _
       And TableHasColumns(wb, "TemplatesTable", Array("TEMPLATE_SCOPE", "RECIPE_ID", "INGREDIENT_ID", "PROCESS", "TARGET_TABLE", "TARGET_COLUMN", "FORMULA", "GUID", "NOTES", "ACTIVE", "CREATED_AT", "UPDATED_AT")) _
       And TableHasColumns(wb, "ProductionLog", Array("TIMESTAMP", "RECIPE", "RECIPE_ID", "DEPARTMENT", "DESCRIPTION", "PROCESS", "OUTPUT", "PREDICTED OUTPUT", "REAL OUTPUT", "BATCH", "BATCH_ID", "RECALL CODE", "ITEM_CODE", "VENDORS", "VENDOR_CODE", "ITEM", "UOM", "QUANTITY", "LOCATION", "System_Key", "INPUT/OUTPUT", "INGREDIENT_ID", "GUID")) _
       And TableHasColumns(wb, "BatchCodesLog", Array("RECIPE", "RECIPE_ID", "PROCESS", "OUTPUT", "UOM", "REAL OUTPUT", "BATCH", "RECALL CODE", "TIMESTAMP", "LOCATION", "USER", "GUID")) Then
        TestEnsureProductionWorkbookSurface_CreatesExpectedTables = 1
    End If

CleanExit:
    CloseNoSavePhase6 wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestProductionForm_InitializeCreatesTabbedSurface() As Long
    Dim wb As Workbook
    Dim report As String
    Dim pageCount As Long

    Set wb = Application.Workbooks.Add

    On Error GoTo CleanFail
    If Not modRoleWorkbookSurfaces.EnsureProductionWorkbookSurface(wb, report) Then GoTo CleanExit
    wb.Activate

    pageCount = frmProduction.TestPageCount()

    If pageCount = 4 _
       And frmProduction.TestRunPaletteCanonicalItemCodeStorage() = 1 _
       And frmProduction.TestAssignmentItemRowsGrowWithoutTableCollision(wb) = 1 _
       And frmProduction.TestProductionCheckRowsRecognizeSkuIdentity(wb) = 1 Then
        TestProductionForm_InitializeCreatesTabbedSurface = 1
    End If

CleanExit:
    On Error Resume Next
    Unload frmProduction
    On Error GoTo 0
    CloseNoSavePhase6 wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestProductionForm_OutputSelectionMapsPastBlankTableRows() As Long
    Dim wb As Workbook
    Dim report As String
    Dim loOutput As ListObject
    Dim lr As ListRow

    Set wb = Application.Workbooks.Add

    On Error GoTo CleanFail
    If Not modRoleWorkbookSurfaces.EnsureProductionWorkbookSurface(wb, report) Then GoTo CleanExit
    Set loOutput = FindTable(wb, "ProductionOutput")
    If loOutput Is Nothing Then GoTo CleanExit

    If loOutput.DataBodyRange Is Nothing Then
        Set lr = loOutput.ListRows.Add
    Else
        loOutput.DataBodyRange.ClearContents
    End If
    Set lr = loOutput.ListRows.Add
    SetTableValueByColumn loOutput, lr.Index, "PROCESS", "BLEND"
    SetTableValueByColumn loOutput, lr.Index, "OUTPUT", "Finished Tea"
    SetTableValueByColumn loOutput, lr.Index, "ITEM_CODE", "SKU-PROD-OUT"
    SetTableValueByColumn loOutput, lr.Index, "System_Key", "SYS-PROD-OUT-1202"

    wb.Activate
    If frmProduction.TestSelectedProductionOutputTableRow(wb, 0) = 2 Then
        TestProductionForm_OutputSelectionMapsPastBlankTableRows = 1
    End If

CleanExit:
    On Error Resume Next
    Unload frmProduction
    On Error GoTo 0
    CloseNoSavePhase6 wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestProductionCompleteRun_BuildsDeltasFromStagedRowsWithoutInvSysData() As Long
    Dim wb As Workbook
    Dim report As String
    Dim loInv As ListObject
    Dim loOutput As ListObject
    Dim loCheck As ListObject
    Dim lr As ListRow
    Dim result As String

    Set wb = Application.Workbooks.Add

    On Error GoTo CleanFail
    If Not modRoleWorkbookSurfaces.EnsureProductionWorkbookSurface(wb, report) Then GoTo CleanExit
    Set loInv = FindTable(wb, "invSys")
    Set loOutput = FindTable(wb, "ProductionOutput")
    Set loCheck = FindTable(wb, "Prod_invSys_Check")
    If loInv Is Nothing Or loOutput Is Nothing Or loCheck Is Nothing Then GoTo CleanExit

    If Not loInv.DataBodyRange Is Nothing Then loInv.DataBodyRange.Delete
    If Not loOutput.DataBodyRange Is Nothing Then loOutput.DataBodyRange.Delete
    If Not loCheck.DataBodyRange Is Nothing Then loCheck.DataBodyRange.Delete

    Set lr = loOutput.ListRows.Add
    SetTableValueByColumn loOutput, lr.Index, "PROCESS", "BLEND"
    SetTableValueByColumn loOutput, lr.Index, "OUTPUT", "Finished Tea"
    SetTableValueByColumn loOutput, lr.Index, "REAL OUTPUT", 8
    SetTableValueByColumn loOutput, lr.Index, "BATCH", 1
    SetTableValueByColumn loOutput, lr.Index, "System_Key", "SYS-PROD-OUT-1202"
    SetTableValueByColumn loOutput, lr.Index, "ITEM_CODE", "SKU-PROD-OUT"

    Set lr = loCheck.ListRows.Add
    SetTableValueByColumn loCheck, lr.Index, "System_Key", "SYS-PROD-IN-1201"
    SetTableValueByColumn loCheck, lr.Index, "ITEM_CODE", "SKU-TEA-IN"
    SetTableValueByColumn loCheck, lr.Index, "ITEM", "Tea Input"
    SetTableValueByColumn loCheck, lr.Index, "USED", 12

    result = mProduction.TestCompletionDeltasFromStagedRows(loOutput, loCheck)
    If result <> "OK|MadeSystemKey=SYS-PROD-OUT-1202;MadeQty=8;UsedSystemKey=SYS-PROD-IN-1201;UsedQty=12" Then
        mLastTestFailure = result
        GoTo CleanExit
    End If

    SetTableValueByColumn loOutput, 1, "System_Key", ""
    SetTableValueByColumn loOutput, 1, "ITEM_CODE", "SKU-PROD-OUT"
    result = mProduction.TestSelectedMadeDeltaSkuIdentity(loOutput)
    If Left$(result, 3) = "OK|" _
       And InStr(1, result, "|SKU-PROD-OUT|8", vbTextCompare) > 0 _
       And result <> "OK||SKU-PROD-OUT|8" Then
        TestProductionCompleteRun_BuildsDeltasFromStagedRowsWithoutInvSysData = 1
    Else
        mLastTestFailure = result
    End If

CleanExit:
    CloseNoSavePhase6 wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestProductionRun_CheckInStagesOutsideInvSysReadModel() As Long
    Dim wb As Workbook
    Dim report As String
    Dim loInv As ListObject
    Dim loCheck As ListObject
    Dim lr As ListRow
    Dim result As String

    Set wb = Application.Workbooks.Add
    On Error GoTo CleanFail
    If Not modRoleWorkbookSurfaces.EnsureProductionWorkbookSurface(wb, report) Then GoTo CleanExit
    Set loInv = FindTable(wb, "invSys")
    Set loCheck = FindTable(wb, "Prod_invSys_Check")
    If loInv Is Nothing Or loCheck Is Nothing Then GoTo CleanExit
    If Not loInv.DataBodyRange Is Nothing Then loInv.DataBodyRange.Delete
    If Not loCheck.DataBodyRange Is Nothing Then loCheck.DataBodyRange.Delete

    Set lr = loInv.ListRows.Add
    SetTableValueByColumn loInv, lr.Index, "System_Key", "SYS-MALAWI-96"
    SetTableValueByColumn loInv, lr.Index, "ITEM_CODE", "SKU-MALAWI-FINE-CUT"
    SetTableValueByColumn loInv, lr.Index, "ITEM", "Malawi Fine Cut Black Tea"
    SetTableValueByColumn loInv, lr.Index, "USED", 0
    SetTableValueByColumn loInv, lr.Index, "TOTAL INV", 3175

    result = mProduction.TestProductionUsedStagingDoesNotMutateInvSys(loInv, loCheck, "SYS-MALAWI-96", 32.5)
    If Left$(result, 3) = "OK|" _
       And CDbl(GetTableValueByColumn(loInv, 1, "USED")) = 0 _
       And CDbl(GetTableValueByColumn(loInv, 1, "TOTAL INV")) = 3175 _
       And CDbl(GetTableValueByColumn(loCheck, 1, "USED")) = 32.5 Then
        TestProductionRun_CheckInStagesOutsideInvSysReadModel = 1
    Else
        mLastTestFailure = result
    End If

CleanExit:
    CloseNoSavePhase6 wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestProductionCompleteRun_ResolvesLooseOutputNameFromCanonicalPicker() As Long
    Dim pickerItems(1 To 2, 1 To 7) As Variant
    Dim result As String

    On Error GoTo CleanFail
    pickerItems(1, 1) = "SYS-BREWED-1301"
    pickerItems(1, 2) = "Brewed Black Tea"
    pickerItems(1, 3) = "LBS"
    pickerItems(1, 6) = "Finished concentrated black tea"
    pickerItems(1, 7) = "SKU-BREWED-BLACK-TEA"
    pickerItems(2, 1) = "SYS-GREEN-1302"
    pickerItems(2, 2) = "Green Tea"
    pickerItems(2, 3) = "LBS"
    pickerItems(2, 7) = "SKU-GREEN-TEA"

    result = mProduction.TestLookupOutputSystemKeyFromPicker(pickerItems, "Brew Black Tea")
    If Left$(result, Len("SYS-BREWED-1301|")) = "SYS-BREWED-1301|" Then
        TestProductionCompleteRun_ResolvesLooseOutputNameFromCanonicalPicker = 1
    End If
    Exit Function

CleanFail:
End Function

Public Function TestProductionCompleteRun_LogsOutputIdempotently() As Long
    Dim wb As Workbook
    Dim report As String
    Dim loOutput As ListObject
    Dim loLog As ListObject
    Dim lr As ListRow
    Dim result As String

    Set wb = Application.Workbooks.Add

    On Error GoTo CleanFail
    If Not modRoleWorkbookSurfaces.EnsureProductionWorkbookSurface(wb, report) Then GoTo CleanExit
    Set loOutput = FindTable(wb, "ProductionOutput")
    Set loLog = FindTable(wb, "ProductionLog")
    If loOutput Is Nothing Or loLog Is Nothing Then GoTo CleanExit

    If Not loOutput.DataBodyRange Is Nothing Then loOutput.DataBodyRange.Delete
    If Not loLog.DataBodyRange Is Nothing Then loLog.DataBodyRange.Delete

    Set lr = loOutput.ListRows.Add
    SetTableValueByColumn loOutput, lr.Index, "PROCESS", "BREW"
    SetTableValueByColumn loOutput, lr.Index, "OUTPUT", "Brew Black Tea"
    SetTableValueByColumn loOutput, lr.Index, "UOM", "LBS"
    SetTableValueByColumn loOutput, lr.Index, "REAL OUTPUT", 400
    SetTableValueByColumn loOutput, lr.Index, "BATCH", 1
    SetTableValueByColumn loOutput, lr.Index, "System_Key", "SYS-BREWED-1301"

    result = mProduction.TestLogProductionOutputRow(wb.Worksheets("Production"), loOutput, 1)
    If Left$(result, 3) = "OK|" And loLog.ListRows.Count = 1 Then
        If CDbl(GetTableValueByColumn(loLog, 1, "REAL OUTPUT")) = 400 _
           And CStr(GetTableValueByColumn(loLog, 1, "System_Key")) = "SYS-BREWED-1301" Then
            TestProductionCompleteRun_LogsOutputIdempotently = 1
        End If
    End If

CleanExit:
    CloseNoSavePhase6 wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestProductionCompleteRun_PreservesImmutableOutputSystemKey() As Long
    Dim wb As Workbook
    Dim report As String
    Dim loInv As ListObject
    Dim loOutput As ListObject
    Dim lr As ListRow
    Dim result As String

    Set wb = Application.Workbooks.Add

    On Error GoTo CleanFail
    If Not modRoleWorkbookSurfaces.EnsureProductionWorkbookSurface(wb, report) Then GoTo CleanExit
    Set loInv = FindTable(wb, "invSys")
    Set loOutput = FindTable(wb, "ProductionOutput")
    If loInv Is Nothing Or loOutput Is Nothing Then GoTo CleanExit

    If Not loInv.DataBodyRange Is Nothing Then loInv.DataBodyRange.Delete
    If Not loOutput.DataBodyRange Is Nothing Then loOutput.DataBodyRange.Delete

    Set lr = loInv.ListRows.Add
    SetTableValueByColumn loInv, lr.Index, "System_Key", "SYS-BREWED-1301"
    SetTableValueByColumn loInv, lr.Index, "ITEM_CODE", "SKU-BREWED-BLACK-TEA"
    SetTableValueByColumn loInv, lr.Index, "ITEM", "Brewed Black Tea"
    SetTableValueByColumn loInv, lr.Index, "DESCRIPTION", "Finished concentrated black tea"

    Set lr = loOutput.ListRows.Add
    SetTableValueByColumn loOutput, lr.Index, "PROCESS", "BREW"
    SetTableValueByColumn loOutput, lr.Index, "OUTPUT", "Brew Black Tea"
    SetTableValueByColumn loOutput, lr.Index, "REAL OUTPUT", 400
    SetTableValueByColumn loOutput, lr.Index, "BATCH", 2
    SetTableValueByColumn loOutput, lr.Index, "System_Key", "SYS-OUTPUT-67"
    SetTableValueByColumn loOutput, lr.Index, "ITEM_CODE", "SKU-BREWED-BLACK-TEA"

    result = mProduction.TestSelectedMadeDeltaSystemKey(loOutput, loInv)
    If Left$(result, Len("OK|SYS-OUTPUT-67|")) = "OK|SYS-OUTPUT-67|" Then
        If CStr(GetTableValueByColumn(loOutput, 1, "System_Key")) = "SYS-OUTPUT-67" Then
            TestProductionCompleteRun_PreservesImmutableOutputSystemKey = 1
        End If
    End If

CleanExit:
    CloseNoSavePhase6 wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestProductionCompleteRun_UsesCatalogIdentityOutsideInvSysProjection() As Long
    Dim pickerItems(1 To 2, 1 To 7) As Variant
    Dim result As String

    On Error GoTo CleanFail
    pickerItems(1, 1) = "SYS-BREWED-67"
    pickerItems(1, 2) = "Brewed Black Tea"
    pickerItems(1, 3) = "LBS"
    pickerItems(1, 6) = "Finished concentrated black tea"
    pickerItems(1, 7) = "SKU-BREWED-BLACK-TEA"
    pickerItems(2, 1) = "SYS-MALAWI-96"
    pickerItems(2, 2) = "Malawi Fine Cut Black Tea"
    pickerItems(2, 3) = "LB"
    pickerItems(2, 7) = "SKU-MALAWI-FINE-CUT"

    result = mProduction.TestOutputIdentityFromPicker(pickerItems, "SYS-BREWED-67", "Brew Black Tea")
    If result = "SYS-BREWED-67|SKU-BREWED-BLACK-TEA|Brewed Black Tea" Then
        TestProductionCompleteRun_UsesCatalogIdentityOutsideInvSysProjection = 1
    End If
    Exit Function

CleanFail:
End Function

Public Function TestProductionRunInventory_RejectsIdentitylessDomainRowsWithoutOperatorFallback() As Long
    Dim wb As Workbook
    Dim report As String
    Dim loInv As ListObject
    Dim lr As ListRow
    Dim items As Variant

    Set wb = Application.Workbooks.Add

    On Error GoTo CleanFail
    If Not modRoleWorkbookSurfaces.EnsureProductionWorkbookSurface(wb, report) Then GoTo CleanExit
    Set loInv = FindTable(wb, "invSys")
    If loInv Is Nothing Then GoTo CleanExit
    If Not loInv.DataBodyRange Is Nothing Then loInv.DataBodyRange.Delete

    Set lr = loInv.ListRows.Add
    SetTableValueByColumn loInv, lr.Index, "System_Key", "SYS-MALAWI-96"
    SetTableValueByColumn loInv, lr.Index, "ITEM_CODE", "SKU-MALAWI-FINE-CUT"
    SetTableValueByColumn loInv, lr.Index, "ITEM", "Malawi Fine Cut Black Tea"
    SetTableValueByColumn loInv, lr.Index, "UOM", "LB"
    SetTableValueByColumn loInv, lr.Index, "TOTAL INV", 3175
    SetTableValueByColumn loInv, lr.Index, "LOCATION", "CLEARVIEW"

    wb.Activate
    items = mProduction.LoadProductionRunInventoryPickerItems("")
    If IsEmpty(items) Then
        TestProductionRunInventory_RejectsIdentitylessDomainRowsWithoutOperatorFallback = 1
    Else
        mLastTestFailure = "Production run inventory picker accepted an identityless Domain row or fell back to operator-workbook authority."
    End If

CleanExit:
    CloseNoSavePhase6 wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestProductionForm_ClosedCapturedWorkbookDoesNotRebindToActiveWorkbook() As Long
    Dim capturedWb As Workbook
    Dim decoyWb As Workbook
    Dim productionForm As frmProduction
    Dim report As String
    Dim statusText As String

    Set capturedWb = Application.Workbooks.Add
    Set decoyWb = Application.Workbooks.Add

    On Error GoTo CleanFail
    If Not modRoleWorkbookSurfaces.EnsureProductionWorkbookSurface(capturedWb, report) Then GoTo CleanExit
    If Not modRoleWorkbookSurfaces.EnsureProductionWorkbookSurface(decoyWb, report) Then GoTo CleanExit

    Set productionForm = New frmProduction
    productionForm.SetOperatorWorkbook capturedWb
    capturedWb.Close SaveChanges:=False
    Set capturedWb = Nothing
    decoyWb.Activate

    productionForm.InitializeFromProduction
    statusText = productionForm.TestStatusText()
    If InStr(1, statusText, "captured", vbTextCompare) > 0 _
       And InStr(1, statusText, decoyWb.Name, vbTextCompare) = 0 Then
        TestProductionForm_ClosedCapturedWorkbookDoesNotRebindToActiveWorkbook = 1
    Else
        mLastTestFailure = "Production form silently rebound after its captured workbook closed. Status=" & statusText
    End If

CleanExit:
    On Error Resume Next
    If Not productionForm Is Nothing Then Unload productionForm
    CloseNoSavePhase6 capturedWb
    CloseNoSavePhase6 decoyWb
    On Error GoTo 0
    Exit Function
CleanFail:
    mLastTestFailure = "Captured-workbook regression raised: " & Err.Description
    Resume CleanExit
End Function

Public Function TestProductionDesignStaging_DoesNotMutateOperatorRecipes() As Long
    Dim operatorWb As Workbook
    Dim stagingWb As Workbook
    Dim wsOperator As Worksheet
    Dim loOperator As ListObject
    Dim loStaging As ListObject
    Dim bom(1 To 2, 1 To 10) As Variant
    Dim report As String

    Set operatorWb = Application.Workbooks.Add(xlWBATWorksheet)
    Set wsOperator = operatorWb.Worksheets(1)
    wsOperator.Name = "Recipes"
    wsOperator.Range("A1").Value = "RECIPE_ID"
    wsOperator.Range("B1").Value = "RECIPE"
    wsOperator.Range("A2").Value = "LOCAL-1"
    wsOperator.Range("B2").Value = "Local Draft"
    Set loOperator = wsOperator.ListObjects.Add(xlSrcRange, wsOperator.Range("A1:B2"), , xlYes)
    loOperator.Name = "Recipes"

    bom(1, 1) = 1
    bom(1, 2) = 1
    bom(1, 3) = "USED"
    bom(1, 4) = "SKU-TEA"
    bom(1, 7) = 2.5
    bom(1, 8) = "LB"
    bom(1, 9) = 100
    bom(1, 10) = "Tea"
    bom(2, 1) = 2
    bom(2, 2) = 1
    bom(2, 3) = "OUTPUT"
    bom(2, 4) = "SKU-BREW"
    bom(2, 7) = 10
    bom(2, 8) = "LB"
    bom(2, 10) = "Brew"

    On Error GoTo CleanFail
    Set stagingWb = mProduction.BuildDesignRecipeStagingWorkbookFromData( _
        "DES-1", "Released Brew", "canonical", bom, report)
    If stagingWb Is Nothing Then mLastTestFailure = report: GoTo CleanExit
    If stagingWb Is operatorWb Then mLastTestFailure = "Staging reused the operator workbook.": GoTo CleanExit
    Set loStaging = stagingWb.Worksheets("Recipes").ListObjects("Recipes")
    If loStaging.ListRows.Count <> 2 Then mLastTestFailure = "Expected two staged BOM rows.": GoTo CleanExit
    If loOperator.ListRows.Count <> 1 Then mLastTestFailure = "Operator Recipes row count changed.": GoTo CleanExit
    If CStr(loOperator.DataBodyRange.Cells(1, 1).Value) <> "LOCAL-1" Then mLastTestFailure = "Operator recipe identity changed.": GoTo CleanExit
    If CStr(loOperator.DataBodyRange.Cells(1, 2).Value) <> "Local Draft" Then mLastTestFailure = "Operator recipe content changed.": GoTo CleanExit
    If CStr(loStaging.DataBodyRange.Cells(1, loStaging.ListColumns("RECIPE_ID").Index).Value) <> "DES-1" Then _
        mLastTestFailure = "Canonical DesignId was not staged.": GoTo CleanExit
    TestProductionDesignStaging_DoesNotMutateOperatorRecipes = 1

CleanExit:
    CloseNoSavePhase6 stagingWb
    CloseNoSavePhase6 operatorWb
    Exit Function
CleanFail:
    mLastTestFailure = Err.Description
    Resume CleanExit
End Function

Public Function TestAdminDesignLifecycleForm_BuildsVersionPicker() As Long
    Dim frm As frmAdminDesignLifecycle

    On Error GoTo CleanFail
    Set frm = New frmAdminDesignLifecycle
    If frm.TestLayoutReady() = 1 Then TestAdminDesignLifecycleForm_BuildsVersionPicker = 1
CleanExit:
    On Error Resume Next
    If Not frm Is Nothing Then Unload frm
    Set frm = Nothing
    On Error GoTo 0
    Exit Function
CleanFail:
    mLastTestFailure = Err.Description
    Resume CleanExit
End Function

Public Function TestProductionForm_BatchDisplaysCompletedCount() As Long
    Dim wb As Workbook
    Dim report As String
    Dim loOutput As ListObject
    Dim loLog As ListObject
    Dim lr As ListRow

    Set wb = Application.Workbooks.Add

    On Error GoTo CleanFail
    If Not modRoleWorkbookSurfaces.EnsureProductionWorkbookSurface(wb, report) Then GoTo CleanExit
    Set loOutput = FindTable(wb, "ProductionOutput")
    Set loLog = FindTable(wb, "ProductionLog")
    If loOutput Is Nothing Or loLog Is Nothing Then GoTo CleanExit

    If Not loOutput.DataBodyRange Is Nothing Then loOutput.DataBodyRange.Delete
    If Not loLog.DataBodyRange Is Nothing Then loLog.DataBodyRange.Delete

    Set lr = loOutput.ListRows.Add
    SetTableValueByColumn loOutput, lr.Index, "PROCESS", "BREW"
    SetTableValueByColumn loOutput, lr.Index, "OUTPUT", "Brew Black Tea"
    SetTableValueByColumn loOutput, lr.Index, "System_Key", "SYS-BREWED-1301"

    wb.Activate
    If frmProduction.TestProductionOutputDisplayedBatch(wb, 0) <> "0" Then GoTo CleanExit

    Set lr = loLog.ListRows.Add
    SetTableValueByColumn loLog, lr.Index, "PROCESS", "BREW"
    SetTableValueByColumn loLog, lr.Index, "OUTPUT", "Brew Black Tea"
    SetTableValueByColumn loLog, lr.Index, "ITEM", "Brew Black Tea"
    SetTableValueByColumn loLog, lr.Index, "REAL OUTPUT", 400
    SetTableValueByColumn loLog, lr.Index, "BATCH", 1
    SetTableValueByColumn loLog, lr.Index, "System_Key", "SYS-BREWED-1301"
    SetTableValueByColumn loLog, lr.Index, "TIMESTAMP", Now

    If frmProduction.TestProductionOutputDisplayedBatch(wb, 0) = "1" Then
        TestProductionForm_BatchDisplaysCompletedCount = 1
    End If

CleanExit:
    On Error Resume Next
    Unload frmProduction
    On Error GoTo 0
    CloseNoSavePhase6 wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestProductionForm_AssignmentIncludesOutputRecipeRows() As Long
    Dim ingredients(1 To 2, 1 To 7) As Variant
    Dim formOutputRows As Long

    On Error GoTo CleanFail

    ingredients(1, 1) = "ING-OAR-1"
    ingredients(1, 2) = "Apple Juice"
    ingredients(1, 3) = "GAL"
    ingredients(1, 4) = "BLEND"
    ingredients(1, 5) = "USED"
    ingredients(1, 6) = 1
    ingredients(1, 7) = 100
    ingredients(2, 1) = "OUT-OAR-1"
    ingredients(2, 2) = "Finished Apple Juice"
    ingredients(2, 3) = "GAL"
    ingredients(2, 4) = "BLEND"
    ingredients(2, 5) = "OUTPUT"
    ingredients(2, 6) = 1
    ingredients(2, 7) = 100

    formOutputRows = frmProduction.TestFillAssignmentIoCount(ingredients, "OUTPUT")
    If formOutputRows = 1 Then
        TestProductionForm_AssignmentIncludesOutputRecipeRows = 1
    End If

CleanExit:
    On Error Resume Next
    Unload frmProduction
    On Error GoTo 0
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestProductionForm_AssignmentOutputRejectsAcceptableInventory() As Long
    Dim wb As Workbook
    Dim report As String
    Dim loRecipes As ListObject
    Dim lr As ListRow

    Set wb = Application.Workbooks.Add

    On Error GoTo CleanFail
    If Not modRoleWorkbookSurfaces.EnsureProductionWorkbookSurface(wb, report) Then GoTo CleanExit
    Set loRecipes = FindTable(wb, "Recipes")
    If loRecipes Is Nothing Then GoTo CleanExit
    If Not loRecipes.DataBodyRange Is Nothing Then loRecipes.DataBodyRange.Delete

    Set lr = loRecipes.ListRows.Add
    SetTableValueByColumn loRecipes, lr.Index, "RECIPE", "Output Guard Recipe"
    SetTableValueByColumn loRecipes, lr.Index, "RECIPE_ID", "R-OUT-GUARD"
    SetTableValueByColumn loRecipes, lr.Index, "PROCESS", "1"
    SetTableValueByColumn loRecipes, lr.Index, "INPUT/OUTPUT", "OUTPUT"
    SetTableValueByColumn loRecipes, lr.Index, "INGREDIENT", "Finished Tea"
    SetTableValueByColumn loRecipes, lr.Index, "UOM", "LB"
    SetTableValueByColumn loRecipes, lr.Index, "AMOUNT", 400
    SetTableValueByColumn loRecipes, lr.Index, "INGREDIENT_ID", "OUT-FINISHED-TEA"

    wb.Activate
    If frmProduction.TestAssignmentOutputSelectionClearsStaging(wb, "R-OUT-GUARD", "Finished Tea") = 1 Then
        TestProductionForm_AssignmentOutputRejectsAcceptableInventory = 1
    End If

CleanExit:
    On Error Resume Next
    Unload frmProduction
    On Error GoTo 0
    CloseNoSavePhase6 wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestAdminAddInventoryItemForm_ConfiguresWithoutTypeMismatch() As Long
    Dim formUnderTest As frmAddInventoryItem
    Dim customFields As Object
    Dim defaultStateOk As Boolean

    On Error GoTo CleanFail

    Set formUnderTest = New frmAddInventoryItem
    formUnderTest.Configure "WH-TEST", "S1", "admin", "ITM-TEST-001", 101, "A1"
    defaultStateOk = (formUnderTest.ImagePath = "" _
                      And formUnderTest.NonCountedItem = False _
                      And formUnderTest.StartingQty = 1)
    formUnderTest.TestSetQuantityMode "Utility"
    Set customFields = formUnderTest.CustomFields

    If formUnderTest.GeneratedSku = "ITM-TEST-001" _
       And defaultStateOk _
       And formUnderTest.NonCountedItem = True _
       And formUnderTest.StartingQty = 0 _
       And CStr(customFields("TRACK_QTY")) = "FALSE" _
       And CStr(customFields("ITEM_KIND")) = "UTILITY" Then
        TestAdminAddInventoryItemForm_ConfiguresWithoutTypeMismatch = 1
    End If

CleanExit:
    On Error Resume Next
    If Not formUnderTest Is Nothing Then Unload formUnderTest
    On Error GoTo 0
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestProductionForm_AssignmentInventorySearchFiltersCachedRows() As Long
    Dim rowsData(1 To 3, 1 To 7) As Variant

    On Error GoTo CleanFail

    rowsData(1, 1) = "101"
    rowsData(1, 2) = "Apple Juice"
    rowsData(1, 3) = "GAL"
    rowsData(1, 4) = 12
    rowsData(1, 5) = "A1"
    rowsData(1, 6) = "Cold pressed juice"
    rowsData(1, 7) = "SKU-AJ"
    rowsData(2, 1) = "102"
    rowsData(2, 2) = "Orange Juice"
    rowsData(2, 3) = "GAL"
    rowsData(2, 4) = 8
    rowsData(2, 5) = "B2"
    rowsData(2, 6) = "Citrus juice"
    rowsData(2, 7) = "SKU-OJ"
    rowsData(3, 1) = "103"
    rowsData(3, 2) = "Cane Sugar"
    rowsData(3, 3) = "LB"
    rowsData(3, 4) = 20
    rowsData(3, 5) = "C3"
    rowsData(3, 6) = "Dry ingredient"
    rowsData(3, 7) = "SKU-SUG"

    If frmProduction.TestFilterAssignmentInventoryCount(rowsData, "juice") <> 2 Then GoTo CleanExit
    If frmProduction.TestFilterAssignmentInventoryCount(rowsData, "B2") <> 1 Then GoTo CleanExit
    If frmProduction.TestFilterAssignmentInventoryCount(rowsData, "SKU-SUG") <> 1 Then GoTo CleanExit

    TestProductionForm_AssignmentInventorySearchFiltersCachedRows = 1

CleanExit:
    On Error Resume Next
    Unload frmProduction
    On Error GoTo 0
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestProductionForm_RunLocationRejectsMismatchedInventory() As Long
    On Error GoTo CleanFail

    If frmProduction.TestProductionRunLocationAllowed("CLEARVIEW", "CLEARVIEW", 400) <> 1 Then GoTo CleanExit
    If frmProduction.TestProductionRunLocationAllowed("CLEARVIEW", "A1", 400) <> 0 Then GoTo CleanExit
    If frmProduction.TestProductionRunLocationAllowed("CLEARVIEW", "A1", 0) <> 1 Then GoTo CleanExit
    If frmProduction.TestProductionRunLocationAllowed("", "CLEARVIEW", 400) <> 0 Then GoTo CleanExit

    TestProductionForm_RunLocationRejectsMismatchedInventory = 1

CleanExit:
    On Error Resume Next
    Unload frmProduction
    On Error GoTo 0
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestProductionForm_SaveRecipeAppliesSelectedBuilderLineEdit() As Long
    Dim wb As Workbook
    Dim report As String
    Dim loLines As ListObject
    Dim lr As ListRow

    Set wb = Application.Workbooks.Add

    On Error GoTo CleanFail
    If Not modRoleWorkbookSurfaces.EnsureProductionWorkbookSurface(wb, report) Then GoTo CleanExit
    Set loLines = FindTable(wb, "RecipeBuilder")
    If loLines Is Nothing Then GoTo CleanExit

    If Not loLines.DataBodyRange Is Nothing Then loLines.DataBodyRange.ClearContents
    ' Keep the surface's first physical staging row blank. Visible list indexes
    ' must still map to the correct nonblank table rows.
    Set lr = loLines.ListRows.Add(AlwaysInsert:=False)
    SetTableValueByColumn loLines, lr.Index, "PROCESS", "2"
    SetTableValueByColumn loLines, lr.Index, "INPUT/OUTPUT", "OUTPUT"
    SetTableValueByColumn loLines, lr.Index, "INGREDIENT", "First Visible Line"
    SetTableValueByColumn loLines, lr.Index, "UOM", "LBS"
    SetTableValueByColumn loLines, lr.Index, "AMOUNT", 400

    Set lr = loLines.ListRows.Add(AlwaysInsert:=False)
    SetTableValueByColumn loLines, lr.Index, "PROCESS", "3"
    SetTableValueByColumn loLines, lr.Index, "INPUT/OUTPUT", "OUTPUT"
    SetTableValueByColumn loLines, lr.Index, "INGREDIENT", "Selected Second Line"
    SetTableValueByColumn loLines, lr.Index, "UOM", "LBS"
    SetTableValueByColumn loLines, lr.Index, "AMOUNT", 398
    wb.Activate

    If frmProduction.TestRecipeBuilderSelectedLineProcessUpdate(wb, "9", 1) = 1 Then
        If CStr(GetTableValueByColumn(loLines, 2, "PROCESS")) <> "2" Then
            mLastTestFailure = "Update Line changed the first visible recipe line instead of the selected line."
            GoTo CleanExit
        End If
        If CStr(GetTableValueByColumn(loLines, 3, "PROCESS")) <> "9" Then
            mLastTestFailure = "Update Line did not change the selected second recipe line."
            GoTo CleanExit
        End If
        TestProductionForm_SaveRecipeAppliesSelectedBuilderLineEdit = 1
    End If

CleanExit:
    On Error Resume Next
    Unload frmProduction
    On Error GoTo 0
    CloseNoSavePhase6 wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestProductionForm_RecipeBuilderMovesLinesAndSupportsInstruction() As Long
    Dim wb As Workbook
    Dim report As String
    Dim loLines As ListObject
    Dim lr As ListRow
    Dim moveResult As String

    Set wb = Application.Workbooks.Add

    On Error GoTo CleanFail
    If Not modRoleWorkbookSurfaces.EnsureProductionWorkbookSurface(wb, report) Then GoTo CleanExit
    Set loLines = FindTable(wb, "RecipeBuilder")
    If loLines Is Nothing Then GoTo CleanExit
    If Not loLines.DataBodyRange Is Nothing Then
        loLines.DataBodyRange.ClearContents
    End If

    If loLines.ListRows.Count = 0 Then
        Set lr = loLines.ListRows.Add(AlwaysInsert:=False)
    Else
        Set lr = loLines.ListRows(1)
    End If
    SetTableValueByColumn loLines, lr.Index, "PROCESS", "1"
    SetTableValueByColumn loLines, lr.Index, "INPUT/OUTPUT", "INSTRUCTION"
    SetTableValueByColumn loLines, lr.Index, "INGREDIENT", "Heat water"
    SetTableValueByColumn loLines, lr.Index, "RECIPE_LIST_ROW", 1

    Set lr = loLines.ListRows.Add(AlwaysInsert:=False)
    SetTableValueByColumn loLines, lr.Index, "PROCESS", "1"
    SetTableValueByColumn loLines, lr.Index, "INPUT/OUTPUT", "USED"
    SetTableValueByColumn loLines, lr.Index, "INGREDIENT", "Black Tea"
    SetTableValueByColumn loLines, lr.Index, "RECIPE_LIST_ROW", 2
    wb.Activate

    If frmProduction.TestRecipeBuilderHasInstructionIo() <> 1 Then
        mLastTestFailure = "INSTRUCTION was not present in the Recipe Builder In/Out dropdown."
        GoTo CleanExit
    End If
    If frmProduction.TestRecipeBuilderUomCatalogContains("EA") <> 1 Then
        mLastTestFailure = "Recipe Builder UOM dropdown did not load the warehouse/default catalog."
        GoTo CleanExit
    End If
    If modUomSettings.NormalizeConfiguredUomName("  To   Line  ") <> "TO LINE" Then
        mLastTestFailure = "Spaced UOM names were not normalized while preserving meaningful spaces."
        GoTo CleanExit
    End If
    If frmProduction.TestRecipeBuilderSelectUom("TO LINE") <> "TO LINE" Then
        mLastTestFailure = "Recipe Builder could not safely select a newly added or legacy spaced UOM."
        GoTo CleanExit
    End If
    If frmProduction.TestRecipeBuilderLineActionsFitLayout() <> 1 Then
        mLastTestFailure = "Recipe Builder line action controls overlap or extend into the recipe command column."
        GoTo CleanExit
    End If
    If frmProduction.TestRecipeBuilderLifecycleAndHeadersReady() <> 1 Then
        mLastTestFailure = "Recipe Builder line headers or Release for Production control were not built."
        GoTo CleanExit
    End If
    moveResult = frmProduction.TestRecipeBuilderLineMove(wb, 1, -1)
    If moveResult = "Black Tea|Heat water|1|2" Then
        TestProductionForm_RecipeBuilderMovesLinesAndSupportsInstruction = 1
    Else
        mLastTestFailure = "Unexpected move result: " & moveResult
    End If

CleanExit:
    On Error Resume Next
    Unload frmProduction
    On Error GoTo 0
    CloseNoSavePhase6 wb
    Exit Function
CleanFail:
    mLastTestFailure = Err.Description
    Resume CleanExit
End Function

Public Function TestProductionRecipeIdentity_Base36LatestNameAndBoundWorkbook() As Long
    Dim targetWb As Workbook
    Dim distractorWb As Workbook
    Dim targetHeader As ListObject
    Dim distractorHeader As ListObject
    Dim designs(1 To 2, 1 To 6) As Variant
    Dim canonicalRecipes(1 To 1, 1 To 3) As Variant
    Dim pendingRecipes(1 To 2, 1 To 3) As Variant
    Dim legacyRecipes(1 To 2, 1 To 3) As Variant
    Dim unifiedRecipes As Variant
    Dim unifiedById As Object
    Dim unifiedRow As Long
    Dim report As String
    Dim result As String

    On Error GoTo CleanFail
    If mProduction.TestNextBase36RecipeId(Array("1", "002")) <> "003" Then
        mLastTestFailure = "Base-36 recipe ID allocation did not normalize legacy numeric 1 and advance to 003."
        GoTo CleanExit
    End If

    designs(1, 1) = "001"
    designs(1, 2) = "v1"
    designs(1, 3) = "RECIPE"
    designs(1, 4) = "Brewed Black Tea"
    designs(1, 5) = "old name"
    designs(1, 6) = "RELEASED"
    designs(2, 1) = "001"
    designs(2, 2) = "20260725180000"
    designs(2, 3) = "RECIPE"
    designs(2, 4) = "Malawi Brewed Black Slury"
    designs(2, 5) = "new name"
    designs(2, 6) = "DRAFT"
    If mProduction.TestLatestRecipeNameFromDesignRows(designs, "001") <> "Malawi Brewed Black Slury" Then
        mLastTestFailure = "Latest Designs replay row did not supply the current recipe name."
        GoTo CleanExit
    End If

    canonicalRecipes(1, 1) = "001"
    canonicalRecipes(1, 2) = "Canonical 001"
    pendingRecipes(1, 1) = "1"
    pendingRecipes(1, 2) = "Pending duplicate 001"
    pendingRecipes(2, 1) = "002"
    pendingRecipes(2, 2) = "Pending 002"
    legacyRecipes(1, 1) = "002"
    legacyRecipes(1, 2) = "Legacy duplicate 002"
    legacyRecipes(2, 1) = "003"
    legacyRecipes(2, 2) = "Legacy 003"
    unifiedRecipes = mProduction.TestUnifiedRecipeList( _
        canonicalRecipes, pendingRecipes, legacyRecipes)
    Set unifiedById = CreateObject("Scripting.Dictionary")
    unifiedById.CompareMode = vbTextCompare
    For unifiedRow = LBound(unifiedRecipes, 1) To UBound(unifiedRecipes, 1)
        unifiedById(CStr(unifiedRecipes(unifiedRow, 1))) = CStr(unifiedRecipes(unifiedRow, 2))
    Next unifiedRow
    If unifiedById.Count <> 3 _
       Or unifiedById("001") <> "Canonical 001" _
       Or unifiedById("002") <> "Pending 002" _
       Or unifiedById("003") <> "Legacy 003" Then
        mLastTestFailure = "Unified Saved Recipes list did not preserve canonical, pending, legacy authority order."
        GoTo CleanExit
    End If

    Set targetWb = Application.Workbooks.Add
    Set distractorWb = Application.Workbooks.Add
    If Not modRoleWorkbookSurfaces.EnsureProductionWorkbookSurface(targetWb, report) Then GoTo CleanExit
    If Not modRoleWorkbookSurfaces.EnsureProductionWorkbookSurface(distractorWb, report) Then GoTo CleanExit
    Set targetHeader = FindTable(targetWb, "RB_AddRecipeName")
    Set distractorHeader = FindTable(distractorWb, "RB_AddRecipeName")
    distractorWb.Activate

    result = frmProduction.TestWriteRecipeHeaderToBoundWorkbook( _
        targetWb, "Malawi Brewed Black Slury", "003")
    If result <> "003|Malawi Brewed Black Slury" Then
        mLastTestFailure = "Recipe header was not written to the form-bound operator workbook: " & result
        GoTo CleanExit
    End If
    If Not distractorHeader Is Nothing Then
        If Not distractorHeader.DataBodyRange Is Nothing Then
            If CStr(distractorHeader.DataBodyRange.Cells(1, distractorHeader.ListColumns("RECIPE_NAME").Index).Value) <> "" Then
                mLastTestFailure = "Recipe header leaked into the active distractor workbook."
                GoTo CleanExit
            End If
        End If
    End If

    TestProductionRecipeIdentity_Base36LatestNameAndBoundWorkbook = 1

CleanExit:
    On Error Resume Next
    Unload frmProduction
    CloseNoSavePhase6 targetWb
    CloseNoSavePhase6 distractorWb
    On Error GoTo 0
    Exit Function
CleanFail:
    mLastTestFailure = Err.Description
    Resume CleanExit
End Function

Public Function TestProductionForm_AssignmentReflectsSavedRecipeProcess() As Long
    Dim wb As Workbook
    Dim report As String
    Dim loRecipes As ListObject
    Dim lr As ListRow

    Set wb = Application.Workbooks.Add

    On Error GoTo CleanFail
    If Not modRoleWorkbookSurfaces.EnsureProductionWorkbookSurface(wb, report) Then GoTo CleanExit
    Set loRecipes = FindTable(wb, "Recipes")
    If loRecipes Is Nothing Then GoTo CleanExit

    Set lr = loRecipes.ListRows.Add
    SetTableValueByColumn loRecipes, lr.Index, "RECIPE", "Brewed Black Tea"
    SetTableValueByColumn loRecipes, lr.Index, "RECIPE_ID", "R-TEA"
    SetTableValueByColumn loRecipes, lr.Index, "DESCRIPTION", "strong tea"
    SetTableValueByColumn loRecipes, lr.Index, "PROCESS", "1"
    SetTableValueByColumn loRecipes, lr.Index, "INPUT/OUTPUT", "OUTPUT"
    SetTableValueByColumn loRecipes, lr.Index, "INGREDIENT", "Brew Black Tea"
    SetTableValueByColumn loRecipes, lr.Index, "UOM", "LBS"
    SetTableValueByColumn loRecipes, lr.Index, "AMOUNT", 400
    SetTableValueByColumn loRecipes, lr.Index, "INGREDIENT_ID", "OUT-TEA"
    wb.Activate

    If frmProduction.TestAssignmentIngredientProcessForRecipe(wb, "R-TEA", "Brew Black Tea") = "1" Then
        TestProductionForm_AssignmentReflectsSavedRecipeProcess = 1
    End If

CleanExit:
    On Error Resume Next
    Unload frmProduction
    On Error GoTo 0
    CloseNoSavePhase6 wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestProductionRecipes_LocalRowsWinOverStaleRuntime() As Long
    Dim runtimeRoot As String
    Dim result As String

    On Error GoTo CleanFail

    runtimeRoot = BuildRoleSurfaceTempRoot("prod_recipe_local_first")
    result = CStr(Application.Run("'" & ThisWorkbook.Name & "'!mProduction.TestProductionRecipesLocalRowsWinOverStaleRuntime", runtimeRoot))
    If Left$(result, 2) = "OK" Then
        TestProductionRecipes_LocalRowsWinOverStaleRuntime = 1
    Else
        mLastTestFailure = result
    End If

CleanExit:
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestProductionRecipeBuilder_SaveAfterLoadPersistsEditedLines() As Long
    Dim runtimeRoot As String
    Dim result As String

    On Error GoTo CleanFail

    runtimeRoot = BuildRoleSurfaceTempRoot("prod_recipe_save_after_load")
    result = CStr(Application.Run("'" & ThisWorkbook.Name & "'!mProduction.TestProductionRecipeBuilderSaveAfterLoadPersistsEditedLines", runtimeRoot))
    If Left$(result, 2) = "OK" Then
        TestProductionRecipeBuilder_SaveAfterLoadPersistsEditedLines = 1
    Else
        mLastTestFailure = result
    End If

CleanExit:
    Exit Function
CleanFail:
    mLastTestFailure = Err.Description
    Resume CleanExit
End Function

Public Function TestEnsureShippingWorkbookSurface_RecreatesDeletedArtifacts() As Long
    Dim wb As Workbook
    Dim report As String

    Set wb = Application.Workbooks.Add

    On Error GoTo CleanFail
    If Not modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface(wb, report) Then GoTo CleanExit

    DeleteTablePhase6 wb, "BoxBuilder"
    DeleteTablePhase6 wb, "BoxBOM"
    DeleteTablePhase6 wb, "AggregatePackages_Log"
    DeleteTablePhase6 wb, "ShippingBOMView"

    If Not modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface(wb, report) Then GoTo CleanExit
    If HasTable(wb, "BoxBuilder") _
       And HasTable(wb, "BoxBOM") _
       And HasTable(wb, "AggregatePackages_Log") _
       And HasTable(wb, "ShippingBOMView") _
       And TableHasColumns(wb, "BoxBuilder", Array("Box Name", "UOM", "LOCATION", "DESCRIPTION")) _
       And Not TableHasColumns(wb, "BoxBuilder", Array("ROW")) _
       And TableHasColumns(wb, "BoxBOM", Array("ITEM", "ROW", "QUANTITY", "UOM", "LOCATION", "DESCRIPTION")) _
       And Not WorksheetExists(wb, "ShippingBOM") Then
        TestEnsureShippingWorkbookSurface_RecreatesDeletedArtifacts = 1
    End If

CleanExit:
    CloseNoSavePhase6 wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestProductionIngredientPaletteRuntimeRoundTrip() As Long
    Dim runtimeRoot As String
    Dim result As String

    On Error GoTo CleanFail

    runtimeRoot = BuildRoleSurfaceTempRoot("prod_palette_runtime")
    result = CStr(Application.Run("'" & ThisWorkbook.Name & "'!mProduction.TestProductionIngredientPaletteRuntimeRoundTrip", runtimeRoot))
    If Left$(result, 2) = "OK" Then
        TestProductionIngredientPaletteRuntimeRoundTrip = 1
    Else
        mLastTestFailure = result
    End If

CleanExit:
    Exit Function
CleanFail:
    mLastTestFailure = Err.Description
    Resume CleanExit
End Function

Public Function TestProductionInventoryPickerPrefersCanonicalRuntime() As Long
    Dim runtimeRoot As String
    Dim result As String

    On Error GoTo CleanFail

    runtimeRoot = BuildRoleSurfaceTempRoot("prod_inventory_picker")
    result = CStr(Application.Run("'" & ThisWorkbook.Name & "'!mProduction.TestProductionInventoryPickerPrefersCanonicalRuntime", runtimeRoot))
    If Left$(result, 2) = "OK" Then
        TestProductionInventoryPickerPrefersCanonicalRuntime = 1
    Else
        mLastTestFailure = result
    End If

CleanExit:
    Exit Function
CleanFail:
    mLastTestFailure = Err.Description
    Resume CleanExit
End Function

Public Function TestEnsureAdminWorkbookSurface_CreatesExpectedTables() As Long
    Dim wb As Workbook
    Dim report As String

    Set wb = Application.Workbooks.Add

    On Error GoTo CleanFail
    If Not modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(wb, report) Then GoTo CleanExit
    If Not modAdminConsole.EnsureAdminSchema(wb, report) Then GoTo CleanExit

    If HasTable(wb, "UserCredentials") _
       And HasTable(wb, "Emails") _
       And HasTable(wb, "tblAdminAudit") _
       And HasTable(wb, "tblAdminPoisonQueue") _
       And WorksheetExists(wb, "AdminConsole") _
       And TableHasColumns(wb, "UserCredentials", Array("USER_ID", "USERNAME", "PIN", "ROLE", "STATUS", "LAST LOGIN")) _
       And TableHasColumns(wb, "Emails", Array("EMAIL_ID", "EMAIL_ADDRESS", "DISPLAY_NAME", "STATUS")) _
       And TableHasColumns(wb, "tblAdminAudit", Array("LoggedAtUTC", "Action", "UserId", "WarehouseId", "StationId", "TargetType", "TargetId", "Reason", "Detail", "Result")) _
       And TableHasColumns(wb, "tblAdminPoisonQueue", Array("SourceWorkbook", "SourceTable", "RowIndex", "EventID", "ParentEventId", "UndoOfEventId", "EventType", "CreatedAtUTC", "WarehouseId", "StationId", "UserId", "SKU", "Qty", "Location", "Note", "PayloadJson", "Status", "RetryCount", "ErrorCode", "ErrorMessage", "FailedAtUTC")) Then
        TestEnsureAdminWorkbookSurface_CreatesExpectedTables = 1
    End If

CleanExit:
    CloseNoSavePhase6 wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestResolveAdminTargetWorkbook_PrefersActiveVisibleWorkbook() As Long
    Dim wbVisible As Workbook
    Dim resolved As Workbook

    Set wbVisible = Application.Workbooks.Add

    On Error GoTo CleanFail
    wbVisible.Activate
    Set resolved = modAdminWorkbookTarget.ResolveAdminTargetWorkbook(Nothing, ThisWorkbook, False)

    If Not resolved Is Nothing Then
        If StrComp(resolved.Name, wbVisible.Name, vbTextCompare) = 0 Then
            TestResolveAdminTargetWorkbook_PrefersActiveVisibleWorkbook = 1
        End If
    End If

CleanExit:
    CloseNoSavePhase6 wbVisible
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestResolveAdminTargetWorkbook_ExplicitWorkbookWinsOverActiveWorkbook() As Long
    Dim wbActive As Workbook
    Dim wbExplicit As Workbook
    Dim resolved As Workbook

    Set wbActive = Application.Workbooks.Add
    Set wbExplicit = Application.Workbooks.Add

    On Error GoTo CleanFail
    wbActive.Activate
    Set resolved = modAdminWorkbookTarget.ResolveAdminTargetWorkbook(wbExplicit, ThisWorkbook, False)

    If Not resolved Is Nothing Then
        If StrComp(resolved.Name, wbExplicit.Name, vbTextCompare) = 0 Then
            TestResolveAdminTargetWorkbook_ExplicitWorkbookWinsOverActiveWorkbook = 1
        End If
    End If

CleanExit:
    CloseNoSavePhase6 wbExplicit
    CloseNoSavePhase6 wbActive
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestOpenUserManagement_WithoutWorkbookArgTargetsActiveWorkbook() As Long
    Dim wbVisible As Workbook
    Dim report As String

    Set wbVisible = Application.Workbooks.Add

    On Error GoTo CleanFail
    If Not modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(wbVisible, report) Then GoTo CleanExit
    wbVisible.Activate

    If Not modAdminConsole.OpenUserManagement(, report) Then GoTo CleanExit

    If StrComp(Application.ActiveWorkbook.Name, wbVisible.Name, vbTextCompare) = 0 _
       And StrComp(Application.ActiveSheet.Name, "UserCredentials", vbTextCompare) = 0 _
       And WorksheetExists(wbVisible, "UserCredentials") Then
        TestOpenUserManagement_WithoutWorkbookArgTargetsActiveWorkbook = 1
    End If

CleanExit:
    CloseNoSavePhase6 wbVisible
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestOpenAdminConsole_WithoutRuntime_DoesNotCreateDefaultWarehouse() As Long
    Dim wbVisible As Workbook
    Dim report As String
    Dim tempRoot As String
    Dim createdFolder As String
    Dim createdConfig As String
    Dim statusText As String

    Set wbVisible = Application.Workbooks.Add

    On Error GoTo CleanFail
    tempRoot = TestPhase2Helpers.BuildUniqueTestFolder("Phase6AdminConsoleNoRuntime")
    modRuntimeWorkbooks.SetCoreDataRootOverride tempRoot

    If Not modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(wbVisible, report) Then GoTo CleanExit
    wbVisible.Activate

    If Not modAdminConsole.OpenAdminConsole(wbVisible, report) Then GoTo CleanExit

    createdFolder = tempRoot & "\WH1"
    createdConfig = createdFolder & "\WH1.invSys.Config.xlsb"
    statusText = Trim$(CStr(wbVisible.Worksheets("AdminConsole").Range("B16").Value))

    If StrComp(CStr(wbVisible.Worksheets("AdminConsole").Range("B3").Value), "<none>", vbTextCompare) <> 0 Then GoTo CleanExit
    If StrComp(CStr(wbVisible.Worksheets("AdminConsole").Range("B4").Value), "<none>", vbTextCompare) <> 0 Then GoTo CleanExit
    If InStr(1, statusText, "did not create any warehouse files", vbTextCompare) = 0 Then GoTo CleanExit
    If Len(Dir$(createdFolder, vbDirectory)) > 0 Then GoTo CleanExit
    If Len(Dir$(createdConfig, vbNormal)) > 0 Then GoTo CleanExit

    TestOpenAdminConsole_WithoutRuntime_DoesNotCreateDefaultWarehouse = 1

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    CloseNoSavePhase6 wbVisible
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestEnsureProductionWorkbookSurface_RecreatesDeletedArtifacts() As Long
    Dim wb As Workbook
    Dim report As String

    Set wb = Application.Workbooks.Add

    On Error GoTo CleanFail
    If Not modRoleWorkbookSurfaces.EnsureProductionWorkbookSurface(wb, report) Then GoTo CleanExit

    DeleteTablePhase6 wb, "IP_ChooseIngredient"
    DeleteTablePhase6 wb, "ProductionLog"
    If WorksheetExists(wb, "IngredientPalette") Then
        DeleteWorksheetPhase6 wb, "IngredientPalette"
    Else
        DeleteWorksheetPhase6 wb, "IngredientsPalette"
    End If

    If Not modRoleWorkbookSurfaces.EnsureProductionWorkbookSurface(wb, report) Then GoTo CleanExit
    If HasTable(wb, "IP_ChooseIngredient") _
       And HasTable(wb, "ProductionLog") _
       And HasTable(wb, "IngredientPalette") _
       And WorksheetExistsAny(wb, Array("IngredientPalette", "IngredientsPalette")) Then
        TestEnsureProductionWorkbookSurface_RecreatesDeletedArtifacts = 1
    End If

CleanExit:
    CloseNoSavePhase6 wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Private Function HasTable(ByVal wb As Workbook, ByVal tableName As String) As Boolean
    HasTable = Not FindTable(wb, tableName) Is Nothing
End Function

Private Function TableHasColumns(ByVal wb As Workbook, ByVal tableName As String, ByVal expectedColumns As Variant) As Boolean
    Dim lo As ListObject
    Dim i As Long

    Set lo = FindTable(wb, tableName)
    If lo Is Nothing Then Exit Function

    For i = LBound(expectedColumns) To UBound(expectedColumns)
        If Not HasColumn(lo, CStr(expectedColumns(i))) Then Exit Function
    Next i

    TableHasColumns = True
End Function

Private Function TableColumnHidden(ByVal wb As Workbook, ByVal tableName As String, ByVal columnName As String) As Boolean
    Dim lo As ListObject
    Dim lc As ListColumn

    Set lo = FindTable(wb, tableName)
    If lo Is Nothing Then Exit Function

    For Each lc In lo.ListColumns
        If StrComp(lc.Name, columnName, vbTextCompare) = 0 Then
            TableColumnHidden = CBool(lc.Range.EntireColumn.Hidden)
            Exit Function
        End If
    Next lc
End Function

Private Function WorksheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet

    For Each ws In wb.Worksheets
        If StrComp(ws.Name, sheetName, vbTextCompare) = 0 Then
            WorksheetExists = True
            Exit Function
        End If
    Next ws
End Function

Private Function WorksheetExistsAny(ByVal wb As Workbook, ByVal sheetNames As Variant) As Boolean
    Dim i As Long
    For i = LBound(sheetNames) To UBound(sheetNames)
        If WorksheetExists(wb, CStr(sheetNames(i))) Then
            WorksheetExistsAny = True
            Exit Function
        End If
    Next i
End Function

Private Function FindTable(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet

    For Each ws In wb.Worksheets
        On Error Resume Next
        Set FindTable = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not FindTable Is Nothing Then Exit Function
    Next ws
End Function

Private Sub DeleteTablePhase6(ByVal wb As Workbook, ByVal tableName As String)
    Dim lo As ListObject

    Set lo = FindTable(wb, tableName)
    If lo Is Nothing Then Exit Sub
    On Error Resume Next
    lo.Delete
    On Error GoTo 0
End Sub

Private Sub DeleteWorksheetPhase6(ByVal wb As Workbook, ByVal sheetName As String)
    Dim ws As Worksheet

    For Each ws In wb.Worksheets
        If StrComp(ws.Name, sheetName, vbTextCompare) = 0 Then
            On Error Resume Next
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
            On Error GoTo 0
            Exit Sub
        End If
    Next ws
End Sub

Private Function HasColumn(ByVal lo As ListObject, ByVal columnName As String) As Boolean
    Dim lc As ListColumn

    If lo Is Nothing Then Exit Function
    For Each lc In lo.ListColumns
        If StrComp(lc.Name, columnName, vbTextCompare) = 0 Then
            HasColumn = True
            Exit Function
        End If
    Next lc
End Function

Private Sub SetTableValueByColumn(ByVal lo As ListObject, ByVal rowIndex As Long, _
                                  ByVal columnName As String, ByVal valueOut As Variant)
    Dim idx As Long

    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If rowIndex < 1 Or rowIndex > lo.DataBodyRange.Rows.Count Then Exit Sub
    idx = TableColumnIndex(lo, columnName)
    If idx = 0 Then Exit Sub
    lo.DataBodyRange.Cells(rowIndex, idx).Value = valueOut
End Sub

Private Function GetTableValueByColumn(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As Variant
    Dim idx As Long

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    If rowIndex < 1 Or rowIndex > lo.DataBodyRange.rows.Count Then Exit Function
    idx = TableColumnIndex(lo, columnName)
    If idx = 0 Then Exit Function
    GetTableValueByColumn = lo.DataBodyRange.Cells(rowIndex, idx).Value
End Function

Private Function TableColumnIndex(ByVal lo As ListObject, ByVal columnName As String) As Long
    Dim lc As ListColumn

    If lo Is Nothing Then Exit Function
    For Each lc In lo.ListColumns
        If StrComp(lc.Name, columnName, vbTextCompare) = 0 Then
            TableColumnIndex = lc.Index
            Exit Function
        End If
    Next lc
End Function

Private Sub CloseNoSavePhase6(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

Private Function BuildRoleSurfaceTempRoot(ByVal leafName As String) As String
    BuildRoleSurfaceTempRoot = Environ$("TEMP") & "\invSys_" & leafName & "_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(CLng(Timer * 1000))
    If Len(Dir$(BuildRoleSurfaceTempRoot, vbDirectory)) = 0 Then MkDir BuildRoleSurfaceTempRoot
End Function
