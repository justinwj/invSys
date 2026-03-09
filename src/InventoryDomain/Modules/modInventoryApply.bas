Attribute VB_Name = "modInventoryApply"
Option Explicit

Public Const APPLY_STATUS_APPLIED As String = "APPLIED"
Public Const APPLY_STATUS_SKIP_DUP As String = "SKIP_DUP"

Public Function ApplyReceiveEvent(ByVal evt As Object, _
                                  Optional ByVal inventoryWb As Workbook = Nothing, _
                                  Optional ByVal runId As String = "", _
                                  Optional ByRef statusOut As String = "", _
                                  Optional ByRef errorCode As String = "", _
                                  Optional ByRef errorMessage As String = "") As Boolean
    On Error GoTo FailApply

    Dim wb As Workbook
    Dim loLog As ListObject
    Dim loApplied As ListObject
    Dim eventId As String
    Dim sku As String
    Dim qty As Double
    Dim warehouseId As String
    Dim stationId As String
    Dim userId As String
    Dim locationVal As String
    Dim noteVal As String
    Dim sourceInbox As String
    Dim occurredAt As Date
    Dim appliedAt As Date
    Dim appliedSeq As Long
    Dim undoOfEventId As String
    Dim r As ListRow

    Set wb = ResolveInventoryWorkbook(GetEventString(evt, "WarehouseId"), inventoryWb)
    If wb Is Nothing Then
        errorCode = "INVENTORY_WORKBOOK_NOT_FOUND"
        errorMessage = "Inventory workbook not found."
        Exit Function
    End If

    If Not modInventorySchema.EnsureInventorySchema(wb) Then
        errorCode = "INVENTORY_SCHEMA_INVALID"
        errorMessage = "Unable to validate inventory schema."
        Exit Function
    End If

    Set loLog = FindListObjectByNameApply(wb, "tblInventoryLog")
    Set loApplied = FindListObjectByNameApply(wb, "tblAppliedEvents")
    If loLog Is Nothing Or loApplied Is Nothing Then
        errorCode = "INVENTORY_TABLE_MISSING"
        errorMessage = "Required inventory tables not found."
        Exit Function
    End If

    SetSheetProtectionApply loLog.Parent, False
    SetSheetProtectionApply loApplied.Parent, False

    eventId = GetEventString(evt, "EventID")
    warehouseId = GetEventString(evt, "WarehouseId")
    stationId = GetEventString(evt, "StationId")
    userId = GetEventString(evt, "UserId")
    sku = GetEventString(evt, "SKU")
    locationVal = GetEventString(evt, "Location")
    noteVal = GetEventString(evt, "Note")
    sourceInbox = GetEventString(evt, "SourceInbox")
    undoOfEventId = GetEventString(evt, "UndoOfEventId")

    If eventId = "" Then
        errorCode = "INVALID_EVENT"
        errorMessage = "EventID is required."
        Exit Function
    End If
    If Not TryGetEventDate(evt, "CreatedAtUTC", occurredAt) Then
        errorCode = "INVALID_EVENT"
        errorMessage = "CreatedAtUTC is required and must be a valid date."
        Exit Function
    End If
    If warehouseId = "" Or stationId = "" Or userId = "" Then
        errorCode = "INVALID_EVENT"
        errorMessage = "WarehouseId, StationId, and UserId are required."
        Exit Function
    End If
    If sku = "" Then
        errorCode = "INVALID_SKU"
        errorMessage = "SKU is required."
        Exit Function
    End If
    If Not TryGetEventDouble(evt, "Qty", qty) Then
        errorCode = "INVALID_QTY"
        errorMessage = "Qty is required and must be numeric."
        Exit Function
    End If
    If qty <= 0 Then
        errorCode = "INVALID_QTY"
        errorMessage = "Qty must be greater than zero."
        Exit Function
    End If

    If AppliedEventExists(loApplied, eventId) Then
        statusOut = APPLY_STATUS_SKIP_DUP
        ApplyReceiveEvent = True
        GoTo CleanExit
    End If

    If Not ValidateSkuExists(wb, sku) Then
        errorCode = "INVALID_SKU"
        errorMessage = "SKU not found in inventory catalog."
        GoTo CleanExit
    End If

    appliedAt = Now
    appliedSeq = GetNextAppliedSeq(wb)
    If runId = "" Then runId = "RUN-" & Format$(appliedAt, "yyyymmddhhnnss")

    Set r = loLog.ListRows.Add
    SetTableRowValue loLog, r.Index, "EventID", eventId
    SetTableRowValue loLog, r.Index, "UndoOfEventId", undoOfEventId
    SetTableRowValue loLog, r.Index, "AppliedSeq", appliedSeq
    SetTableRowValue loLog, r.Index, "EventType", "RECEIVE"
    SetTableRowValue loLog, r.Index, "OccurredAtUTC", occurredAt
    SetTableRowValue loLog, r.Index, "AppliedAtUTC", appliedAt
    SetTableRowValue loLog, r.Index, "WarehouseId", warehouseId
    SetTableRowValue loLog, r.Index, "StationId", stationId
    SetTableRowValue loLog, r.Index, "UserId", userId
    SetTableRowValue loLog, r.Index, "SKU", sku
    SetTableRowValue loLog, r.Index, "QtyDelta", qty
    SetTableRowValue loLog, r.Index, "Location", locationVal
    SetTableRowValue loLog, r.Index, "Note", noteVal

    Set r = loApplied.ListRows.Add
    SetTableRowValue loApplied, r.Index, "EventID", eventId
    SetTableRowValue loApplied, r.Index, "UndoOfEventId", undoOfEventId
    SetTableRowValue loApplied, r.Index, "AppliedSeq", appliedSeq
    SetTableRowValue loApplied, r.Index, "AppliedAtUTC", appliedAt
    SetTableRowValue loApplied, r.Index, "RunId", runId
    SetTableRowValue loApplied, r.Index, "SourceInbox", sourceInbox
    SetTableRowValue loApplied, r.Index, "Status", APPLY_STATUS_APPLIED

    statusOut = APPLY_STATUS_APPLIED
    ApplyReceiveEvent = True
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    If Not loLog Is Nothing Then SetSheetProtectionApply loLog.Parent, True
    If Not loApplied Is Nothing Then SetSheetProtectionApply loApplied.Parent, True
    On Error GoTo 0
    Exit Function

FailApply:
    On Error Resume Next
    If Not loLog Is Nothing Then SetSheetProtectionApply loLog.Parent, True
    If Not loApplied Is Nothing Then SetSheetProtectionApply loApplied.Parent, True
    On Error GoTo 0
    If errorCode = "" Then errorCode = "APPLY_EXCEPTION"
    If errorMessage = "" Then errorMessage = Err.Description
End Function

Public Function ResolveInventoryWorkbook(Optional ByVal warehouseId As String = "", _
                                         Optional ByVal inventoryWb As Workbook = Nothing) As Workbook
    Dim wb As Workbook

    If Not inventoryWb Is Nothing Then
        Set ResolveInventoryWorkbook = inventoryWb
        Exit Function
    End If

    For Each wb In Application.Workbooks
        If IsInventoryWorkbookName(wb.Name) Then
            If warehouseId = "" Or InStr(1, wb.Name, warehouseId, vbTextCompare) > 0 Then
                Set ResolveInventoryWorkbook = wb
                Exit Function
            End If
        End If
    Next wb

    For Each wb In Application.Workbooks
        If WorkbookHasListObjectApply(wb, "tblInventoryLog") And _
           WorkbookHasListObjectApply(wb, "tblAppliedEvents") And _
           WorkbookHasListObjectApply(wb, "tblLocks") Then
            Set ResolveInventoryWorkbook = wb
            Exit Function
        End If
    Next wb
End Function

Private Function AppliedEventExists(ByVal lo As ListObject, ByVal eventId As String) As Boolean
    Dim i As Long
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To lo.ListRows.Count
        If StrComp(SafeTrimApply(GetCellByColumnApply(lo, i, "EventID")), eventId, vbTextCompare) = 0 Then
            AppliedEventExists = True
            Exit Function
        End If
    Next i
End Function

Private Function ValidateSkuExists(ByVal wb As Workbook, ByVal sku As String) As Boolean
    Dim hasCatalog As Boolean

    ValidateSkuExists = SearchSkuInTable(FindListObjectByNameApply(wb, "tblSkuCatalog"), sku, hasCatalog)
    If ValidateSkuExists Then Exit Function
    If SearchSkuInTable(FindListObjectByNameApply(wb, "invSys"), sku, hasCatalog) Then
        ValidateSkuExists = True
        Exit Function
    End If
    If SearchSkuInTable(FindListObjectByNameApply(wb, "tblItemSearchIndex"), sku, hasCatalog) Then
        ValidateSkuExists = True
        Exit Function
    End If

    If Not hasCatalog Then ValidateSkuExists = True
End Function

Private Function SearchSkuInTable(ByVal lo As ListObject, ByVal sku As String, ByRef hasCatalog As Boolean) As Boolean
    Dim idx As Long
    Dim i As Long
    Dim valueInRow As String

    If lo Is Nothing Then Exit Function

    idx = GetColumnIndexApply(lo, "SKU")
    If idx = 0 Then idx = GetColumnIndexApply(lo, "ITEM_CODE")
    If idx = 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    hasCatalog = True
    For i = 1 To lo.ListRows.Count
        valueInRow = SafeTrimApply(lo.DataBodyRange.Cells(i, idx).Value)
        If valueInRow <> "" Then
            If StrComp(valueInRow, sku, vbTextCompare) = 0 Then
                SearchSkuInTable = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Function GetNextAppliedSeq(ByVal wb As Workbook) As Long
    Dim lo As ListObject
    Dim idx As Long
    Dim i As Long
    Dim currentVal As Variant

    Set lo = FindListObjectByNameApply(wb, "tblAppliedEvents")
    If lo Is Nothing Then
        GetNextAppliedSeq = 1
        Exit Function
    End If

    idx = GetColumnIndexApply(lo, "AppliedSeq")
    If idx = 0 Or lo.DataBodyRange Is Nothing Then
        GetNextAppliedSeq = 1
        Exit Function
    End If

    For i = 1 To lo.ListRows.Count
        currentVal = lo.DataBodyRange.Cells(i, idx).Value
        If IsNumeric(currentVal) Then
            If CLng(currentVal) > GetNextAppliedSeq Then GetNextAppliedSeq = CLng(currentVal)
        End If
    Next i

    GetNextAppliedSeq = GetNextAppliedSeq + 1
End Function

Private Function TryGetEventDate(ByVal evt As Object, ByVal key As String, ByRef valueOut As Date) As Boolean
    Dim rawValue As Variant
    If Not TryGetEventValue(evt, key, rawValue) Then Exit Function
    If Not IsDate(rawValue) Then Exit Function
    valueOut = CDate(rawValue)
    TryGetEventDate = True
End Function

Private Function TryGetEventDouble(ByVal evt As Object, ByVal key As String, ByRef valueOut As Double) As Boolean
    Dim rawValue As Variant
    If Not TryGetEventValue(evt, key, rawValue) Then Exit Function
    If Not IsNumeric(rawValue) Then Exit Function
    valueOut = CDbl(rawValue)
    TryGetEventDouble = True
End Function

Private Function GetEventString(ByVal evt As Object, ByVal key As String) As String
    Dim rawValue As Variant
    If TryGetEventValue(evt, key, rawValue) Then
        GetEventString = SafeTrimApply(rawValue)
    End If
End Function

Private Function TryGetEventValue(ByVal evt As Object, ByVal key As String, ByRef valueOut As Variant) As Boolean
    On Error Resume Next
    If evt Is Nothing Then Exit Function
    valueOut = evt(key)
    TryGetEventValue = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Sub SetTableRowValue(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String, ByVal valueOut As Variant)
    Dim idx As Long
    idx = GetColumnIndexApply(lo, columnName)
    If idx = 0 Then Exit Sub
    lo.DataBodyRange.Cells(rowIndex, idx).Value = valueOut
End Sub

Private Function GetCellByColumnApply(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As Variant
    Dim idx As Long
    idx = GetColumnIndexApply(lo, columnName)
    If idx = 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    GetCellByColumnApply = lo.DataBodyRange.Cells(rowIndex, idx).Value
End Function

Private Function GetColumnIndexApply(ByVal lo As ListObject, ByVal columnName As String) As Long
    Dim i As Long
    If lo Is Nothing Then Exit Function
    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).Name, columnName, vbTextCompare) = 0 Then
            GetColumnIndexApply = i
            Exit Function
        End If
    Next i
End Function

Private Function FindListObjectByNameApply(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    On Error Resume Next
    For Each ws In wb.Worksheets
        Set FindListObjectByNameApply = ws.ListObjects(tableName)
        If Not FindListObjectByNameApply Is Nothing Then Exit Function
    Next ws
    On Error GoTo 0
End Function

Private Function WorkbookHasListObjectApply(ByVal wb As Workbook, ByVal tableName As String) As Boolean
    WorkbookHasListObjectApply = Not (FindListObjectByNameApply(wb, tableName) Is Nothing)
End Function

Private Function IsInventoryWorkbookName(ByVal wbName As String) As Boolean
    Dim n As String
    n = LCase$(wbName)
    IsInventoryWorkbookName = (n Like "wh*.invsys.data.inventory.xlsb") Or _
                              (n Like "wh*.invsys.data.inventory.xlsx") Or _
                              (n Like "wh*.invsys.data.inventory.xlsm")
End Function

Private Function SafeTrimApply(ByVal v As Variant) As String
    On Error Resume Next
    SafeTrimApply = Trim$(CStr(v))
End Function

Private Sub SetSheetProtectionApply(ByVal ws As Worksheet, ByVal protectAfter As Boolean)
    If ws Is Nothing Then Exit Sub
    If protectAfter Then
        On Error Resume Next
        ws.Protect UserInterfaceOnly:=True
        On Error GoTo 0
    Else
        If Not ws.ProtectContents Then Exit Sub
        On Error Resume Next
        ws.Unprotect
        On Error GoTo 0
        If ws.ProtectContents Then
            Err.Raise vbObjectError + 2201, "modInventoryApply.SetSheetProtectionApply", _
                      "Worksheet '" & ws.Name & "' is protected and could not be unprotected. " & _
                      "Excel automation cannot add table rows while the sheet remains protected."
        End If
    End If
End Sub
