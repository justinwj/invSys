Attribute VB_Name = "modTS_Received"
Option Explicit

Private mInitializeCount As Long
Private mLastWorkbookName As String
Private mLastConfirmSucceeded As Boolean
Private mLastConfirmStatus As String
Private Const SHEET_RECEIVING As String = "ReceivedTally"
Private Const TABLE_RECEIVING As String = "ReceivedTally"
Private Const TABLE_AGG_RECEIVED As String = "AggregateReceived"

Public Sub InitializeReceivingUiForWorkbook(Optional ByVal targetWb As Workbook = Nothing)
    mInitializeCount = mInitializeCount + 1
    If Not targetWb Is Nothing Then mLastWorkbookName = targetWb.Name
End Sub

Public Sub ResetReceivingUiStub()
    mInitializeCount = 0
    mLastWorkbookName = vbNullString
    mLastConfirmSucceeded = False
    mLastConfirmStatus = vbNullString
End Sub

Public Function GetReceivingUiStubInitializeCount() As Long
    GetReceivingUiStubInitializeCount = mInitializeCount
End Function

Public Function GetReceivingUiStubLastWorkbookName() As String
    GetReceivingUiStubLastWorkbookName = mLastWorkbookName
End Function

Public Function RefreshReceivingUiForWorkbook(Optional ByVal targetWb As Workbook = Nothing, _
                                              Optional ByVal sourceType As String = "LOCAL", _
                                              Optional ByRef report As String = "") As Boolean
    InitializeReceivingUiForWorkbook targetWb
    EnforceReceivingSupportSheetsHidden targetWb
    RefreshReceivingUiForWorkbook = True
End Function

Public Sub EnforceReceivingSupportSheetsHidden(ByVal wb As Workbook)
    On Error GoTo CleanExit

    Dim supportNames As Variant
    Dim i As Long
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Sub
    supportNames = Array(SHEET_RECEIVING, "InventoryManagement", "ReceivedLog")
    For i = LBound(supportNames) To UBound(supportNames)
        Set ws = Nothing
        On Error Resume Next
        Set ws = wb.Worksheets(CStr(supportNames(i)))
        On Error GoTo CleanExit
        If Not ws Is Nothing Then
            If ws.Visible <> xlSheetVeryHidden Then
                If CanHideStubWorksheet(wb, ws) Then ws.Visible = xlSheetVeryHidden
            End If
        End If
    Next i

CleanExit:
End Sub

Public Function LoadReceivingFormInventory(Optional ByVal filterText As String = "") As Variant
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject

    Set wb = ResolveStubWorkbook(Application.ActiveWorkbook)
    If wb Is Nothing Then Exit Function
    Set ws = StubSheet(wb, "InventoryManagement")
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    Set lo = ws.ListObjects("invSys")
    On Error GoTo 0
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    LoadReceivingFormInventory = lo.DataBodyRange.Value2
End Function

Public Function LoadReceivingFormTable(ByVal tableName As String) As Variant
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject

    Set wb = ResolveStubWorkbook(Application.ActiveWorkbook)
    If wb Is Nothing Then Exit Function
    Set ws = StubSheet(wb, SHEET_RECEIVING)
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
    On Error GoTo 0
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    LoadReceivingFormTable = lo.DataBodyRange.Value2
End Function

Public Function StageReceivingFormLine(ByVal refNumber As String, ByVal rowValue As Long, _
                                       ByVal qty As Double, ByRef report As String) As Boolean
    StageReceivingFormLine = StageReceivingFormLineForWorkbook(Application.ActiveWorkbook, refNumber, rowValue, qty, report)
End Function

Public Function StageReceivingFormLineForWorkbook(ByVal targetWb As Workbook, _
                                                  ByVal refNumber As String, _
                                                  ByVal rowValue As Long, _
                                                  ByVal qty As Double, _
                                                  ByRef report As String) As Boolean
    StageReceivingFormLineForWorkbook = StageReceivingFormItemForWorkbook(targetWb, refNumber, rowValue, vbNullString, qty, report)
End Function

Public Function StageReceivingFormItemForWorkbook(ByVal targetWb As Workbook, _
                                                  ByVal refNumber As String, _
                                                  ByVal rowValue As Long, _
                                                  ByVal itemCodeValue As String, _
                                                  ByVal qty As Double, _
                                                  ByRef report As String) As Boolean
    Dim wb As Workbook
    Dim loInv As ListObject
    Dim loRt As ListObject
    Dim loAgg As ListObject
    Dim invIdx As Long
    Dim itemCode As String
    Dim itemName As String
    Dim uomVal As String
    Dim locVal As String
    Dim descVal As String
    Dim vendorVal As String
    Dim vendorCode As String
    Dim requestedItemCode As String

    refNumber = Trim$(refNumber)
    requestedItemCode = Trim$(itemCodeValue)
    If refNumber = "" Then report = "Ref number is required.": Exit Function
    If rowValue <= 0 And requestedItemCode = "" Then report = "Select an inventory item first.": Exit Function
    If qty <= 0 Then report = "Quantity must be greater than zero.": Exit Function

    Set wb = ResolveStubWorkbook(targetWb)
    If wb Is Nothing Then report = "Open a receiving operator workbook first.": Exit Function
    Set loInv = StubTable(wb, "invSys")
    Set loRt = StubTable(wb, TABLE_RECEIVING)
    Set loAgg = StubTable(wb, TABLE_AGG_RECEIVED)
    If loInv Is Nothing Or loRt Is Nothing Or loAgg Is Nothing Then report = "Receiving test tables not found.": Exit Function

    If requestedItemCode <> "" Then invIdx = FindStubRowByText(loInv, "ITEM_CODE", requestedItemCode)
    If invIdx = 0 And rowValue > 0 Then invIdx = FindStubRowByLong(loInv, "ROW", rowValue)
    If invIdx = 0 Then report = "Inventory row " & CStr(rowValue) & " was not found.": Exit Function
    itemCode = StubValue(loInv, invIdx, "ITEM_CODE")
    itemName = StubValue(loInv, invIdx, "ITEM")
    uomVal = StubValue(loInv, invIdx, "UOM")
    locVal = StubValue(loInv, invIdx, "LOCATION")
    descVal = StubValue(loInv, invIdx, "DESCRIPTION")
    vendorVal = StubValue(loInv, invIdx, "VENDOR(s)")
    vendorCode = StubValue(loInv, invIdx, "VENDOR_CODE")
    If rowValue <= 0 Then rowValue = CLng(Val(StubValue(loInv, invIdx, "ROW")))
    If itemName = "" And itemCode = "" Then report = "Inventory row " & CStr(rowValue) & " was not found.": Exit Function

    AppendStubReceivedTally loRt, refNumber, itemName, qty, rowValue
    AppendStubAggregate loAgg, refNumber, itemCode, vendorVal, vendorCode, descVal, itemName, uomVal, qty, locVal, rowValue
    report = "Staged " & CStr(qty) & " " & uomVal & " of " & itemName & "."
    StageReceivingFormItemForWorkbook = True
End Function

Public Sub ClearReceivingFormStaging()
    Dim wb As Workbook
    Set wb = ResolveStubWorkbook(Application.ActiveWorkbook)
    If wb Is Nothing Then Exit Sub
    ClearStubTable StubTable(wb, TABLE_RECEIVING)
    ClearStubTable StubTable(wb, TABLE_AGG_RECEIVED)
End Sub

Public Sub ConfirmWrites()
    mLastConfirmSucceeded = True
    mLastConfirmStatus = "Confirm Writes succeeded."
End Sub

Public Function LastConfirmWritesSucceeded() As Boolean
    LastConfirmWritesSucceeded = mLastConfirmSucceeded
End Function

Public Function LastConfirmWritesStatus() As String
    LastConfirmWritesStatus = mLastConfirmStatus
End Function

Public Sub MacroUndo()
End Sub

Public Sub MacroRedo()
End Sub

Private Sub AppendStubReceivedTally(ByVal lo As ListObject, ByVal refNumber As String, _
                                    ByVal itemName As String, ByVal qty As Double, ByVal rowValue As Long)
    Dim lr As ListRow
    Dim rowIndex As Long

    rowIndex = FindStubReceivedTallyMatch(lo, refNumber, itemName)
    If rowIndex > 0 Then
        SetStubValue lo, rowIndex, "QUANTITY", CDbl(Val(StubValue(lo, rowIndex, "QUANTITY"))) + qty
        SetStubValue lo, rowIndex, "ROW", rowValue
        Exit Sub
    End If

    Set lr = FirstBlankStubListRow(lo)
    If lr Is Nothing Then Set lr = lo.ListRows.Add
    SetStubValue lo, lr.Index, "REF_NUMBER", refNumber
    SetStubValue lo, lr.Index, "ITEMS", itemName
    SetStubValue lo, lr.Index, "QUANTITY", qty
    SetStubValue lo, lr.Index, "ROW", rowValue
End Sub

Private Sub AppendStubAggregate(ByVal lo As ListObject, ByVal refNumber As String, ByVal itemCode As String, _
                                ByVal vendors As String, ByVal vendorCode As String, ByVal descr As String, _
                                ByVal itemName As String, ByVal uomVal As String, ByVal qty As Double, _
                                ByVal locVal As String, ByVal rowValue As Long)
    Dim lr As ListRow
    Dim rowIndex As Long

    rowIndex = FindStubAggregateMatch(lo, itemCode, rowValue)
    If rowIndex = 0 Then
        Set lr = FirstBlankStubListRow(lo)
        If lr Is Nothing Then Set lr = lo.ListRows.Add
        rowIndex = lr.Index
    End If
    SetStubValue lo, rowIndex, "REF_NUMBER", AppendStubRef(StubValue(lo, rowIndex, "REF_NUMBER"), refNumber)
    SetStubValue lo, rowIndex, "ITEM_CODE", itemCode
    SetStubValue lo, rowIndex, "VENDORS", vendors
    SetStubValue lo, rowIndex, "VENDOR_CODE", vendorCode
    SetStubValue lo, rowIndex, "DESCRIPTION", descr
    SetStubValue lo, rowIndex, "ITEM", itemName
    SetStubValue lo, rowIndex, "UOM", uomVal
    SetStubValue lo, rowIndex, "QUANTITY", CDbl(Val(StubValue(lo, rowIndex, "QUANTITY"))) + qty
    SetStubValue lo, rowIndex, "LOCATION", locVal
    SetStubValue lo, rowIndex, "ROW", rowValue
End Sub

Private Function FindStubReceivedTallyMatch(ByVal lo As ListObject, ByVal refNumber As String, ByVal itemName As String) As Long
    Dim i As Long

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To lo.ListRows.Count
        If StrComp(StubValue(lo, i, "ITEMS"), itemName, vbTextCompare) = 0 _
           And StubRefListContains(StubValue(lo, i, "REF_NUMBER"), refNumber) Then
            FindStubReceivedTallyMatch = i
            Exit Function
        End If
    Next i
End Function

Private Function FindStubAggregateMatch(ByVal lo As ListObject, ByVal itemCode As String, ByVal rowValue As Long) As Long
    Dim i As Long

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To lo.ListRows.Count
        If rowValue > 0 And CLng(Val(StubValue(lo, i, "ROW"))) = rowValue Then
            FindStubAggregateMatch = i
            Exit Function
        End If
        If itemCode <> "" And StrComp(StubValue(lo, i, "ITEM_CODE"), itemCode, vbTextCompare) = 0 Then
            FindStubAggregateMatch = i
            Exit Function
        End If
    Next i
End Function

Private Function AppendStubRef(ByVal existingRef As String, ByVal newRef As String) As String
    If Trim$(existingRef) = "" Then
        AppendStubRef = Trim$(newRef)
    ElseIf StubRefListContains(existingRef, newRef) Then
        AppendStubRef = existingRef
    Else
        AppendStubRef = existingRef & "," & Trim$(newRef)
    End If
End Function

Private Function StubRefListContains(ByVal refList As String, ByVal refValue As String) As Boolean
    Dim token As Variant

    refValue = Trim$(refValue)
    If refValue = "" Then Exit Function
    For Each token In Split(refList, ",")
        If StrComp(Trim$(CStr(token)), refValue, vbTextCompare) = 0 Then
            StubRefListContains = True
            Exit Function
        End If
    Next token
End Function

Private Function FirstBlankStubListRow(ByVal lo As ListObject) As ListRow
    Dim lr As ListRow

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    For Each lr In lo.ListRows
        If Application.WorksheetFunction.CountA(lr.Range) = 0 Then
            Set FirstBlankStubListRow = lr
            Exit Function
        End If
    Next lr
End Function

Private Function ResolveStubWorkbook(ByVal targetWb As Workbook) As Workbook
    If Not targetWb Is Nothing Then
        If Not targetWb.IsAddin Then
            Set ResolveStubWorkbook = targetWb
            Exit Function
        End If
    End If
End Function

Private Function StubSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, sheetName, vbTextCompare) = 0 Then
            Set StubSheet = ws
            Exit Function
        End If
    Next ws
End Function

Private Function StubTable(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        On Error Resume Next
        Set StubTable = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not StubTable Is Nothing Then Exit Function
    Next ws
End Function

Private Function FindStubRowByLong(ByVal lo As ListObject, ByVal columnName As String, ByVal expectedValue As Long) As Long
    Dim idx As Long
    Dim r As Long
    idx = StubColumnIndex(lo, columnName)
    If idx = 0 Or lo.DataBodyRange Is Nothing Then Exit Function
    For r = 1 To lo.DataBodyRange.Rows.Count
        If CLng(Val(CStr(lo.DataBodyRange.Cells(r, idx).Value))) = expectedValue Then
            FindStubRowByLong = r
            Exit Function
        End If
    Next r
End Function

Private Function FindStubRowByText(ByVal lo As ListObject, ByVal columnName As String, ByVal expectedValue As String) As Long
    Dim idx As Long
    Dim r As Long

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    expectedValue = Trim$(expectedValue)
    If expectedValue = "" Then Exit Function
    idx = StubColumnIndex(lo, columnName)
    If idx = 0 Then Exit Function
    For r = 1 To lo.ListRows.Count
        If StrComp(Trim$(CStr(lo.DataBodyRange.Cells(r, idx).Value)), expectedValue, vbTextCompare) = 0 Then
            FindStubRowByText = r
            Exit Function
        End If
    Next r
End Function

Private Function StubValue(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As String
    Dim idx As Long
    idx = StubColumnIndex(lo, columnName)
    If idx = 0 Or lo.DataBodyRange Is Nothing Then Exit Function
    StubValue = Trim$(CStr(lo.DataBodyRange.Cells(rowIndex, idx).Value))
End Function

Private Sub SetStubValue(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String, ByVal valueOut As Variant)
    Dim idx As Long
    idx = StubColumnIndex(lo, columnName)
    If idx = 0 Or lo.DataBodyRange Is Nothing Then Exit Sub
    lo.DataBodyRange.Cells(rowIndex, idx).Value = valueOut
End Sub

Private Function StubColumnIndex(ByVal lo As ListObject, ByVal columnName As String) As Long
    Dim lc As ListColumn
    If lo Is Nothing Then Exit Function
    For Each lc In lo.ListColumns
        If StrComp(lc.Name, columnName, vbTextCompare) = 0 Then
            StubColumnIndex = lc.Index
            Exit Function
        End If
    Next lc
End Function

Private Sub ClearStubTable(ByVal lo As ListObject)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    lo.DataBodyRange.Delete
End Sub

Private Function CanHideStubWorksheet(ByVal wb As Workbook, ByVal wsToHide As Worksheet) As Boolean
    Dim ws As Worksheet
    Dim visibleCount As Long

    If wb Is Nothing Or wsToHide Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then visibleCount = visibleCount + 1
    Next ws
    CanHideStubWorksheet = (wsToHide.Visible <> xlSheetVisible Or visibleCount > 1)
End Function
