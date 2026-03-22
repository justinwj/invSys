Attribute VB_Name = "TestPhase6RoleSurfaces"
Option Explicit

Public Function TestEnsureReceivingWorkbookSurface_CreatesExpectedTables() As Long
    Dim wb As Workbook
    Dim report As String

    Set wb = Application.Workbooks.Add

    On Error GoTo CleanFail
    If Not modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface(wb, report) Then GoTo CleanExit
    If HasTable(wb, "ReceivedTally") _
       And HasTable(wb, "AggregateReceived") _
       And HasTable(wb, "ReceivedLog") _
       And HasTable(wb, "invSys") _
       And TableHasColumns(wb, "ReceivedTally", Array("REF_NUMBER", "ITEMS", "QUANTITY", "ROW")) _
       And TableHasColumns(wb, "AggregateReceived", Array("REF_NUMBER", "ITEM_CODE", "VENDORS", "VENDOR_CODE", "DESCRIPTION", "ITEM", "UOM", "QUANTITY", "LOCATION", "ROW")) _
       And TableHasColumns(wb, "ReceivedLog", Array("SNAPSHOT_ID", "ENTRY_DATE", "REF_NUMBER", "ITEMS", "QUANTITY", "UOM", "VENDOR", "LOCATION", "ITEM_CODE", "ROW")) _
       And TableHasColumns(wb, "invSys", Array("ROW", "ITEM_CODE", "ITEM", "UOM", "LOCATION", "DESCRIPTION")) Then
        TestEnsureReceivingWorkbookSurface_CreatesExpectedTables = 1
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
       And HasTable(wb, "AggregateBoxBOM_Log") _
       And HasTable(wb, "AggregatePackages_Log") _
       And HasTable(wb, "Check_invSys") _
       And HasTable(wb, "invSys") _
       And WorksheetExists(wb, "ShippingBOM") _
       And TableHasColumns(wb, "ShipmentsTally", Array("REF_NUMBER", "ITEMS", "QUANTITY", "ROW", "UOM", "LOCATION", "DESCRIPTION")) _
       And TableHasColumns(wb, "BoxBuilder", Array("Box Name", "UOM", "LOCATION", "DESCRIPTION", "ROW")) _
       And TableHasColumns(wb, "BoxBOM", Array("ITEM", "ROW", "QUANTITY", "UOM", "LOCATION", "DESCRIPTION")) _
       And TableHasColumns(wb, "AggregatePackages", Array("ROW", "ITEM_CODE", "ITEM", "QUANTITY", "UOM", "LOCATION")) _
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

Public Function TestEnsureProductionWorkbookSurface_CreatesExpectedTables() As Long
    Dim wb As Workbook
    Dim report As String

    Set wb = Application.Workbooks.Add

    On Error GoTo CleanFail
    If Not modRoleWorkbookSurfaces.EnsureProductionWorkbookSurface(wb, report) Then GoTo CleanExit
    If HasTable(wb, "RB_AddRecipeName") _
       And HasTable(wb, "RecipeBuilder") _
       And HasTable(wb, "RC_RecipeChoose") _
       And HasTable(wb, "ProductionOutput") _
       And HasTable(wb, "Prod_invSys_Check") _
       And HasTable(wb, "Recipes") _
       And HasTable(wb, "TemplatesTable") _
       And HasTable(wb, "ProductionLog") _
       And HasTable(wb, "BatchCodesLog") _
       And HasTable(wb, "invSys") _
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

Private Function WorksheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet

    For Each ws In wb.Worksheets
        If StrComp(ws.Name, sheetName, vbTextCompare) = 0 Then
            WorksheetExists = True
            Exit Function
        End If
    Next ws
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
