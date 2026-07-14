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
       And TableHasColumns(wb, "ShipmentsTally", Array("LINE_ID", "SERVER_RESERVE_EVENT_ID", "REF_NUMBER", "ITEMS", "QUANTITY", "ROW", "UOM", "LOCATION", "DESCRIPTION")) _
       And TableHasColumns(wb, "BoxBuilder", Array("Box Name", "UOM", "LOCATION", "DESCRIPTION")) _
       And Not TableHasColumns(wb, "BoxBuilder", Array("ROW")) _
       And TableHasColumns(wb, "BoxBOM", Array("ITEM", "ROW", "QUANTITY", "UOM", "LOCATION", "DESCRIPTION")) _
       And TableHasColumns(wb, "AggregatePackages", Array("ROW", "ITEM_CODE", "ITEM", "QUANTITY", "UOM", "LOCATION")) _
       And TableHasColumns(wb, "invSysData_Shipping", Array("ROW", "ITEM_CODE", "ITEM", "UOM", "LOCATION", "DESCRIPTION")) _
       And TableHasColumns(wb, "ShippingBOMView", Array("PackageRow", "PackageItem", "ComponentRow", "ComponentQty", "UpdatedAtUTC", "UpdatedBy")) _
       And TableHasColumns(wb, "AggregateBoxBOM_Log", Array("GUID", "USER", "ACTION", "ROW", "ITEM_CODE", "ITEM", "QTY_DELTA", "NEW_VALUE", "TIMESTAMP")) _
       And TableHasColumns(wb, "AggregatePackages_Log", Array("GUID", "USER", "ACTION", "ROW", "ITEM_CODE", "ITEM", "QTY_DELTA", "NEW_VALUE", "TIMESTAMP")) Then
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
       And TableHasColumns(wb, "IP_ChooseItem", Array("ITEMS", "UOM", "DESCRIPTION", "ROW", "RECIPE_ID", "INGREDIENT_ID")) _
       And TableHasColumns(wb, "InventoryPalette_generated", Array("INGREDIENT", "INGREDIENT_ID", "ITEM", "SPLIT %", "QUANTITY", "BASE QUANTITY", "PROCESS", "ROW")) _
       And TableHasColumns(wb, "IngredientPalette", Array("RECIPE_ID", "INGREDIENT_ID", "INPUT/OUTPUT", "ITEM", "PERCENT", "UOM", "AMOUNT", "ROW", "GUID")) _
       And TableHasColumns(wb, "TemplatesTable", Array("TEMPLATE_SCOPE", "RECIPE_ID", "INGREDIENT_ID", "PROCESS", "TARGET_TABLE", "TARGET_COLUMN", "FORMULA", "GUID", "NOTES", "ACTIVE", "CREATED_AT", "UPDATED_AT")) _
       And TableHasColumns(wb, "ProductionLog", Array("TIMESTAMP", "RECIPE", "RECIPE_ID", "DEPARTMENT", "DESCRIPTION", "PROCESS", "OUTPUT", "PREDICTED OUTPUT", "REAL OUTPUT", "BATCH", "BATCH_ID", "RECALL CODE", "ITEM_CODE", "VENDORS", "VENDOR_CODE", "ITEM", "UOM", "QUANTITY", "LOCATION", "ROW", "INPUT/OUTPUT", "INGREDIENT_ID", "GUID")) _
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

    If pageCount = 4 Then
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
