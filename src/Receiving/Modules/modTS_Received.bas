Attribute VB_Name = "modTS_Received"
Option Explicit

Private Const SHEET_RECEIVING As String = "ReceivedTally"
Private Const TABLE_STAGING As String = "ReceivedTally"
Private Const TABLE_AGGREGATE As String = "AggregateReceived"
Private Const TABLE_INVENTORY As String = "invSys"
Private Const SHEET_INVENTORY As String = "InventoryManagement"
Private Const SHEET_LOG As String = "ReceivedLog"

Private mLastConfirmSucceeded As Boolean
Private mLastConfirmStatus As String

Public Sub EnsureGeneratedButtons()
    Dim wb As Workbook
    Dim report As String

    Set wb = ResolveReceivingWorkbook(Application.ActiveWorkbook)
    If wb Is Nothing Then
        ShowReceivingMessage "Activate a Receiving operator workbook before refreshing.", vbExclamation
        Exit Sub
    End If
    If Not RefreshReceivingUiForWorkbook(wb, "LOCAL", report) Then
        ShowReceivingMessage report, vbExclamation
    End If
End Sub

Public Sub InitializeReceivingUiForWorkbook(Optional ByVal targetWb As Workbook = Nothing)
    Dim wb As Workbook
    Dim report As String
    Dim ws As Worksheet

    Set wb = ResolveReceivingWorkbook(targetWb)
    If wb Is Nothing Then Exit Sub
    If Not modOperationsPrimitiveBridge.EnsureReceivingWorkbookSurface(wb.Name, report) Then
        Err.Raise vbObjectError + 7680, "modTS_Received.InitializeReceivingUiForWorkbook", report
    End If
    Set ws = WorkbookSheet(wb, SHEET_RECEIVING)
    If ws Is Nothing Then Exit Sub
    EnsureReceivingButtons ws
    modOperationsPrimitiveBridge.ApplyShapeCapability _
        wb.Name, SHEET_RECEIVING, "btnConfirmWrites", "RECEIVE_POST"
    modOperationsPrimitiveBridge.InitializeReceivingAutoSnapshot wb.Name
    EnforceReceivingSupportSheetsHidden wb
End Sub

Public Function RefreshReceivingUiForWorkbook(Optional ByVal targetWb As Workbook = Nothing, _
                                              Optional ByVal sourceType As String = "LOCAL", _
                                              Optional ByRef report As String = "") As Boolean
    Dim wb As Workbook

    Set wb = ResolveReceivingWorkbook(targetWb)
    If wb Is Nothing Then
        report = "Activate a Receiving operator workbook before refreshing."
        Exit Function
    End If
    InitializeReceivingUiForWorkbook wb
    RefreshReceivingUiForWorkbook = _
        modOperationsPrimitiveBridge.RefreshInventoryReadModel( _
            wb.Name, "", sourceType, report)
    EnforceReceivingSupportSheetsHidden wb
End Function

Public Sub ShowReceivingForm()
    Dim wb As Workbook
    Dim frm As frmReceiving

    Set wb = ResolveReceivingWorkbook(Application.ActiveWorkbook)
    If wb Is Nothing Then
        ShowReceivingMessage "Open a Receiving operator workbook before using the Receiving form.", vbExclamation
        Exit Sub
    End If
    Set frm = New frmReceiving
    frm.SetOperatorWorkbook wb
    frm.InitializeFromReceiving
    EnforceReceivingSupportSheetsHidden wb
    frm.Show vbModeless
End Sub

Public Function RunReceivingConfirmWritesFormActionForTest(ByVal operatorWb As Workbook, _
                                                           Optional ByVal activatedWb As Workbook = Nothing) As String
    Dim frm As frmReceiving

    Set frm = New frmReceiving
    If Not activatedWb Is Nothing Then activatedWb.Activate
    RunReceivingConfirmWritesFormActionForTest = _
        frm.TestRunConfirmWritesActionForWorkbook(operatorWb, activatedWb)
    Unload frm
End Function

Public Function RunReceivingPurchasingTabContractForTest(ByVal operatorWb As Workbook) As String
    Dim frm As frmReceiving

    Set frm = New frmReceiving
    RunReceivingPurchasingTabContractForTest = frm.TestPurchasingTabContract(operatorWb)
    Unload frm
End Function

Public Function ReceivingFormInitializeSmokeForWorkbook(ByVal operatorWb As Workbook) As String
    Dim frm As frmReceiving

    On Error GoTo Failed
    Set frm = New frmReceiving
    ReceivingFormInitializeSmokeForWorkbook = frm.TestInitializeForWorkbook(operatorWb)
CleanExit:
    On Error Resume Next
    If Not frm Is Nothing Then Unload frm
    Set frm = Nothing
    On Error GoTo 0
    Exit Function
Failed:
    ReceivingFormInitializeSmokeForWorkbook = _
        "FAIL|" & CStr(Err.Number) & "|" & Err.Description
    Resume CleanExit
End Function

Public Sub EnforceReceivingSupportSheetsHidden(ByVal wb As Workbook)
    Dim names As Variant
    Dim nameValue As Variant
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Sub
    names = Array(SHEET_RECEIVING, SHEET_INVENTORY, SHEET_LOG)
    For Each nameValue In names
        Set ws = WorkbookSheet(wb, CStr(nameValue))
        If Not ws Is Nothing Then
            If CanHideWorksheet(wb, ws) Then ws.Visible = xlSheetVeryHidden
        End If
    Next nameValue
End Sub

Public Function LoadReceivingFormInventoryForWorkbook(ByVal operatorWb As Workbook, _
                                                      Optional ByVal filterText As String = "") As Variant
    Dim inventoryTable As ListObject
    Dim sourceValues As Variant
    Dim outputValues() As Variant
    Dim trimmedValues() As Variant
    Dim recordIndex As Long
    Dim fieldIndex As Long
    Dim outputIndex As Long
    Dim searchText As String
    Dim searchableText As String
    Dim systemKey As String
    Dim itemCode As String
    Dim itemName As String
    Dim uomValue As String
    Dim qtyValue As String
    Dim locationValue As String
    Dim descriptionValue As String
    Dim vendorValue As String

    Set inventoryTable = FindTable(operatorWb, TABLE_INVENTORY)
    If inventoryTable Is Nothing Or inventoryTable.DataBodyRange Is Nothing Then Exit Function
    If ColumnIndex(inventoryTable, "System_Key") = 0 _
       Or ColumnIndex(inventoryTable, "ITEM_CODE") = 0 Then Exit Function

    searchText = LCase$(Trim$(filterText))
    sourceValues = inventoryTable.DataBodyRange.Value2
    ReDim outputValues(1 To UBound(sourceValues, 1), 1 To 8)
    For recordIndex = 1 To UBound(sourceValues, 1)
        systemKey = CellText(inventoryTable, recordIndex, "System_Key")
        itemCode = CellText(inventoryTable, recordIndex, "ITEM_CODE")
        itemName = CellText(inventoryTable, recordIndex, "ITEM")
        If itemName = "" Then itemName = CellText(inventoryTable, recordIndex, "ItemName")
        uomValue = CellText(inventoryTable, recordIndex, "UOM")
        qtyValue = CellText(inventoryTable, recordIndex, "QtyAvailable")
        If qtyValue = "" Then qtyValue = CellText(inventoryTable, recordIndex, "TOTAL INV")
        locationValue = CellText(inventoryTable, recordIndex, "LOCATION")
        descriptionValue = CellText(inventoryTable, recordIndex, "DESCRIPTION")
        vendorValue = CellText(inventoryTable, recordIndex, "VENDOR(s)")
        If systemKey = "" Or itemCode = "" Then GoTo NextInventoryRecord

        searchableText = LCase$(systemKey & " " & itemCode & " " & itemName & " " & _
                                    descriptionValue & " " & vendorValue & " " & locationValue)
        If searchText <> "" Then
            If InStr(1, searchableText, searchText, vbTextCompare) = 0 Then
                GoTo NextInventoryRecord
            End If
        End If

        outputIndex = outputIndex + 1
        outputValues(outputIndex, 1) = systemKey
        outputValues(outputIndex, 2) = itemCode
        outputValues(outputIndex, 3) = itemName
        outputValues(outputIndex, 4) = uomValue
        outputValues(outputIndex, 5) = qtyValue
        outputValues(outputIndex, 6) = locationValue
        outputValues(outputIndex, 7) = descriptionValue
        outputValues(outputIndex, 8) = vendorValue
NextInventoryRecord:
    Next recordIndex

    If outputIndex = 0 Then Exit Function
    ReDim trimmedValues(1 To outputIndex, 1 To 8)
    For recordIndex = 1 To outputIndex
        For fieldIndex = 1 To 8
            trimmedValues(recordIndex, fieldIndex) = outputValues(recordIndex, fieldIndex)
        Next fieldIndex
    Next recordIndex
    LoadReceivingFormInventoryForWorkbook = trimmedValues
End Function

Public Function LoadReceivingFormInventory(Optional ByVal filterText As String = "") As Variant
    Dim wb As Workbook

    Set wb = ResolveReceivingWorkbook(Application.ActiveWorkbook)
    If wb Is Nothing Then Exit Function
    LoadReceivingFormInventory = _
        LoadReceivingFormInventoryForWorkbook(wb, filterText)
End Function

Public Function LoadReceivingFormTableForWorkbook(ByVal operatorWb As Workbook, _
                                                  ByVal tableName As String) As Variant
    Dim targetTable As ListObject

    Set targetTable = FindTable(operatorWb, tableName)
    If targetTable Is Nothing Or targetTable.DataBodyRange Is Nothing Then Exit Function
    LoadReceivingFormTableForWorkbook = targetTable.DataBodyRange.Value2
End Function

Public Function LoadReceivingFormTable(ByVal tableName As String) As Variant
    Dim wb As Workbook

    Set wb = ResolveReceivingWorkbook(Application.ActiveWorkbook)
    If wb Is Nothing Then Exit Function
    LoadReceivingFormTable = LoadReceivingFormTableForWorkbook(wb, tableName)
End Function

Public Function StageReceivingFormItemForWorkbook(ByVal targetWb As Workbook, _
                                                  ByVal refNumber As String, _
                                                  ByVal sourceSystemKey As String, _
                                                  ByVal itemCodeValue As String, _
                                                  ByVal qty As Double, _
                                                  ByRef report As String) As Boolean
    On Error GoTo Failed

    Dim inventoryTable As ListObject
    Dim stagingTable As ListObject
    Dim aggregateTable As ListObject
    Dim inventoryIndex As Long
    Dim stagingIndex As Long
    Dim aggregateIndex As Long
    Dim receivingSystemKey As String
    Dim eventId As String
    Dim itemCode As String
    Dim itemName As String
    Dim uomValue As String
    Dim locationValue As String

    refNumber = Trim$(refNumber)
    sourceSystemKey = Trim$(sourceSystemKey)
    itemCodeValue = Trim$(itemCodeValue)
    If targetWb Is Nothing Then report = "Receiving workbook was not provided.": Exit Function
    If refNumber = "" Then report = "Ref number is required.": Exit Function
    If sourceSystemKey = "" And itemCodeValue = "" Then report = "Select an inventory item first.": Exit Function
    If qty <= 0 Then report = "Quantity must be greater than zero.": Exit Function

    Set inventoryTable = FindTable(targetWb, TABLE_INVENTORY)
    Set stagingTable = FindTable(targetWb, TABLE_STAGING)
    Set aggregateTable = FindTable(targetWb, TABLE_AGGREGATE)
    If inventoryTable Is Nothing Or stagingTable Is Nothing Or aggregateTable Is Nothing Then
        report = "Receiving inventory or staging tables are missing."
        Exit Function
    End If

    If sourceSystemKey <> "" Then
        inventoryIndex = FindTableRecord(inventoryTable, "System_Key", sourceSystemKey)
    End If
    If inventoryIndex = 0 And itemCodeValue <> "" Then
        inventoryIndex = FindTableRecord(inventoryTable, "ITEM_CODE", itemCodeValue)
    End If
    If inventoryIndex = 0 Then
        report = "The selected inventory item was not found by System_Key or ITEM_CODE."
        Exit Function
    End If

    sourceSystemKey = CellText(inventoryTable, inventoryIndex, "System_Key")
    itemCode = CellText(inventoryTable, inventoryIndex, "ITEM_CODE")
    itemName = CellText(inventoryTable, inventoryIndex, "ITEM")
    If itemName = "" Then itemName = CellText(inventoryTable, inventoryIndex, "ItemName")
    uomValue = CellText(inventoryTable, inventoryIndex, "UOM")
    locationValue = CellText(inventoryTable, inventoryIndex, "LOCATION")

    stagingIndex = FindExistingStagingRecord(stagingTable, refNumber, sourceSystemKey)
    If stagingIndex > 0 Then
        receivingSystemKey = CellText(stagingTable, stagingIndex, "System_Key")
        eventId = CellText(stagingTable, stagingIndex, "EventId")
        SetCellValue stagingTable, stagingIndex, "QUANTITY", _
                     CellNumber(stagingTable, stagingIndex, "QUANTITY") + qty
    Else
        receivingSystemKey = modRoleEventWriter.CreateSystemKey()
        eventId = modRoleEventWriter.CreateSystemKey()
        stagingIndex = FirstBlankOrNewRecord(stagingTable).Index
        SetCellText stagingTable, stagingIndex, "REF_NUMBER", refNumber
        SetCellText stagingTable, stagingIndex, "ITEMS", itemName
        SetCellValue stagingTable, stagingIndex, "QUANTITY", qty
        SetCellText stagingTable, stagingIndex, "System_Key", receivingSystemKey
        SetCellText stagingTable, stagingIndex, "ITEM_CODE", itemCode
        SetCellText stagingTable, stagingIndex, "Source_System_Key", sourceSystemKey
        SetCellText stagingTable, stagingIndex, "EventId", eventId
        SetCellText stagingTable, stagingIndex, "WorkflowState", "STAGED"
    End If

    aggregateIndex = FindTableRecord(aggregateTable, "System_Key", receivingSystemKey)
    If aggregateIndex = 0 Then aggregateIndex = FirstBlankOrNewRecord(aggregateTable).Index
    PopulateAggregateRecord aggregateTable, aggregateIndex, inventoryTable, inventoryIndex, _
                            refNumber, receivingSystemKey, eventId, _
                            CellNumber(stagingTable, stagingIndex, "QUANTITY")

    report = "Staged " & CStr(qty) & " " & uomValue & " of " & itemName & _
             "; System_Key=" & receivingSystemKey
    StageReceivingFormItemForWorkbook = True
    Exit Function
Failed:
    report = "Receiving staging failed: " & Err.Description
End Function

Public Function RebuildAggregationForWorkbook(ByVal targetWb As Workbook, _
                                              Optional ByRef report As String = "") As Boolean
    On Error GoTo Failed

    Dim inventoryTable As ListObject
    Dim stagingTable As ListObject
    Dim aggregateTable As ListObject
    Dim stagingIndex As Long
    Dim inventoryIndex As Long
    Dim aggregateIndex As Long
    Dim sourceSystemKey As String
    Dim receivingSystemKey As String

    Set inventoryTable = FindTable(targetWb, TABLE_INVENTORY)
    Set stagingTable = FindTable(targetWb, TABLE_STAGING)
    Set aggregateTable = FindTable(targetWb, TABLE_AGGREGATE)
    If inventoryTable Is Nothing Or stagingTable Is Nothing Or aggregateTable Is Nothing Then
        report = "Receiving inventory or staging tables are missing."
        Exit Function
    End If
    If Not aggregateTable.DataBodyRange Is Nothing Then aggregateTable.DataBodyRange.Delete
    If stagingTable.DataBodyRange Is Nothing Then
        report = "OK|Rows=0"
        RebuildAggregationForWorkbook = True
        Exit Function
    End If

    For stagingIndex = 1 To stagingTable.ListRows.Count
        receivingSystemKey = CellText(stagingTable, stagingIndex, "System_Key")
        sourceSystemKey = CellText(stagingTable, stagingIndex, "Source_System_Key")
        If receivingSystemKey = "" Then
            report = "ReceivedTally contains a blank System_Key."
            Exit Function
        End If
        inventoryIndex = FindTableRecord(inventoryTable, "System_Key", sourceSystemKey)
        If inventoryIndex = 0 Then
            inventoryIndex = FindTableRecord( _
                inventoryTable, "ITEM_CODE", CellText(stagingTable, stagingIndex, "ITEM_CODE"))
        End If
        If inventoryIndex = 0 Then
            report = "A staged source inventory item is no longer available."
            Exit Function
        End If
        aggregateIndex = FirstBlankOrNewRecord(aggregateTable).Index
        PopulateAggregateRecord aggregateTable, aggregateIndex, inventoryTable, inventoryIndex, _
                                CellText(stagingTable, stagingIndex, "REF_NUMBER"), _
                                receivingSystemKey, _
                                CellText(stagingTable, stagingIndex, "EventId"), _
                                CellNumber(stagingTable, stagingIndex, "QUANTITY")
        SetCellText aggregateTable, aggregateIndex, "WorkflowState", _
                    CellText(stagingTable, stagingIndex, "WorkflowState")
    Next stagingIndex
    report = "OK|Rows=" & CStr(stagingTable.ListRows.Count)
    RebuildAggregationForWorkbook = True
    Exit Function
Failed:
    report = "Receiving aggregation rebuild failed: " & Err.Description
End Function

Public Sub RebuildAggregation()
    Dim wb As Workbook
    Dim report As String

    Set wb = ResolveReceivingWorkbook(Application.ActiveWorkbook)
    If wb Is Nothing Then Exit Sub
    If Not RebuildAggregationForWorkbook(wb, report) Then
        ShowReceivingMessage report, vbExclamation
    End If
End Sub

Public Sub ClearReceivingFormStagingForWorkbook(ByVal operatorWb As Workbook)
    Dim stagingTable As ListObject
    Dim aggregateTable As ListObject

    Set stagingTable = FindTable(operatorWb, TABLE_STAGING)
    Set aggregateTable = FindTable(operatorWb, TABLE_AGGREGATE)
    modReceivingPostingService.ClearReceivingStaging stagingTable, aggregateTable
End Sub

Public Sub ClearReceivingFormStaging()
    Dim wb As Workbook

    Set wb = ResolveReceivingWorkbook(Application.ActiveWorkbook)
    If wb Is Nothing Then Exit Sub
    ClearReceivingFormStagingForWorkbook wb
End Sub

Public Sub ConfirmWrites()
    Dim wb As Workbook
    Dim report As String

    mLastConfirmSucceeded = False
    mLastConfirmStatus = "Confirm Writes did not complete."
    Set wb = ResolveReceivingWorkbook(Application.ActiveWorkbook)
    If wb Is Nothing Then
        mLastConfirmStatus = "Activate a Receiving operator workbook before confirming writes."
        ShowReceivingMessage mLastConfirmStatus, vbExclamation
        Exit Sub
    End If
    mLastConfirmSucceeded = modReceivingPostingService.ExecuteConfirmWrites(wb, report)
    mLastConfirmStatus = report
    If Not mLastConfirmSucceeded Then ShowReceivingMessage report, vbExclamation
End Sub

Public Sub RecordConfirmWritesResult(ByVal succeeded As Boolean, ByVal statusText As String)
    mLastConfirmSucceeded = succeeded
    mLastConfirmStatus = statusText
End Sub

Public Function LastConfirmWritesSucceeded() As Boolean
    LastConfirmWritesSucceeded = mLastConfirmSucceeded
End Function

Public Function LastConfirmWritesStatus() As String
    LastConfirmWritesStatus = mLastConfirmStatus
End Function

Public Sub ShowReceivingDynamicItemSearch(ByVal targetCell As Range)
    ' The supported Receiving search is hosted by frmReceiving.
End Sub

Public Sub HandleReceivingSelectionChange(ByVal target As Range)
    ' Receiving support sheets are projections; operator selection has no write authority.
End Sub

Public Sub HandleReceivingSheetChange(ByVal target As Range)
    On Error GoTo CleanExit

    Dim stagingTable As ListObject
    Dim qtyColumn As ListColumn
    Dim changedCells As Range
    Dim cell As Range
    Dim recordIndex As Long

    If target Is Nothing Then Exit Sub
    Set stagingTable = target.ListObject
    If stagingTable Is Nothing Then Exit Sub
    If StrComp(stagingTable.Name, TABLE_STAGING, vbTextCompare) <> 0 Then Exit Sub
    Set qtyColumn = stagingTable.ListColumns("QUANTITY")
    If qtyColumn.DataBodyRange Is Nothing Then Exit Sub
    Set changedCells = Application.Intersect(target, qtyColumn.DataBodyRange)
    If changedCells Is Nothing Then Exit Sub

    For Each cell In changedCells.Cells
        recordIndex = cell.Row - stagingTable.DataBodyRange.Row + 1
        SyncQuantityFromStaging recordIndex, NzDbl(cell.Value)
    Next cell
CleanExit:
End Sub

Public Sub SyncQuantityFromStaging(ByVal stagingRecordIndex As Long, _
                                   ByVal newQty As Double)
    Dim stagingTable As ListObject
    Dim aggregateTable As ListObject
    Dim systemKey As String
    Dim aggregateIndex As Long

    If stagingRecordIndex <= 0 Then Exit Sub
    Set stagingTable = FindTable(Application.ActiveWorkbook, TABLE_STAGING)
    Set aggregateTable = FindTable(Application.ActiveWorkbook, TABLE_AGGREGATE)
    If stagingTable Is Nothing Or aggregateTable Is Nothing Then Exit Sub
    If stagingRecordIndex > stagingTable.ListRows.Count Then Exit Sub
    systemKey = CellText(stagingTable, stagingRecordIndex, "System_Key")
    If systemKey = "" Then Exit Sub
    aggregateIndex = FindTableRecord(aggregateTable, "System_Key", systemKey)
    If aggregateIndex = 0 Then Exit Sub
    SetCellValue aggregateTable, aggregateIndex, "QUANTITY", newQty
End Sub

Public Function NzDbl(ByVal valueIn As Variant) As Double
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Then Exit Function
    If IsNumeric(valueIn) Then NzDbl = CDbl(valueIn)
End Function

Private Sub PopulateAggregateRecord(ByVal aggregateTable As ListObject, _
                                    ByVal aggregateIndex As Long, _
                                    ByVal inventoryTable As ListObject, _
                                    ByVal inventoryIndex As Long, _
                                    ByVal refNumber As String, _
                                    ByVal receivingSystemKey As String, _
                                    ByVal eventId As String, _
                                    ByVal qty As Double)
    SetCellText aggregateTable, aggregateIndex, "REF_NUMBER", refNumber
    SetCellText aggregateTable, aggregateIndex, "ITEM_CODE", _
                CellText(inventoryTable, inventoryIndex, "ITEM_CODE")
    SetCellText aggregateTable, aggregateIndex, "VENDORS", _
                CellText(inventoryTable, inventoryIndex, "VENDOR(s)")
    SetCellText aggregateTable, aggregateIndex, "VENDOR_CODE", _
                CellText(inventoryTable, inventoryIndex, "VENDOR_CODE")
    SetCellText aggregateTable, aggregateIndex, "DESCRIPTION", _
                CellText(inventoryTable, inventoryIndex, "DESCRIPTION")
    SetCellText aggregateTable, aggregateIndex, "ITEM", _
                CellText(inventoryTable, inventoryIndex, "ITEM")
    If CellText(aggregateTable, aggregateIndex, "ITEM") = "" Then
        SetCellText aggregateTable, aggregateIndex, "ITEM", _
                    CellText(inventoryTable, inventoryIndex, "ItemName")
    End If
    SetCellText aggregateTable, aggregateIndex, "UOM", _
                CellText(inventoryTable, inventoryIndex, "UOM")
    SetCellValue aggregateTable, aggregateIndex, "QUANTITY", qty
    SetCellText aggregateTable, aggregateIndex, "LOCATION", _
                CellText(inventoryTable, inventoryIndex, "LOCATION")
    SetCellText aggregateTable, aggregateIndex, "System_Key", receivingSystemKey
    SetCellText aggregateTable, aggregateIndex, "EventId", eventId
    SetCellText aggregateTable, aggregateIndex, "WorkflowState", "STAGED"
End Sub

Private Function FindExistingStagingRecord(ByVal stagingTable As ListObject, _
                                           ByVal refNumber As String, _
                                           ByVal sourceSystemKey As String) As Long
    Dim recordIndex As Long

    If stagingTable Is Nothing Or stagingTable.DataBodyRange Is Nothing Then Exit Function
    For recordIndex = 1 To stagingTable.ListRows.Count
        If StrComp(CellText(stagingTable, recordIndex, "REF_NUMBER"), _
                   refNumber, vbTextCompare) = 0 _
           And StrComp(CellText(stagingTable, recordIndex, "Source_System_Key"), _
                       sourceSystemKey, vbBinaryCompare) = 0 Then
            FindExistingStagingRecord = recordIndex
            Exit Function
        End If
    Next recordIndex
End Function

Private Function ResolveReceivingWorkbook(Optional ByVal preferredWb As Workbook = Nothing) As Workbook
    If Not preferredWb Is Nothing Then
        If Not preferredWb.IsAddin Then
            Set ResolveReceivingWorkbook = preferredWb
            Exit Function
        End If
    End If
End Function

Private Function WorkbookSheet(ByVal wb As Workbook, _
                               ByVal sheetName As String) As Worksheet
    If wb Is Nothing Then Exit Function
    On Error Resume Next
    Set WorkbookSheet = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function FindTable(ByVal wb As Workbook, _
                           ByVal tableName As String) As ListObject
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        On Error Resume Next
        Set FindTable = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not FindTable Is Nothing Then Exit Function
    Next ws
End Function

Private Function FindTableRecord(ByVal targetTable As ListObject, _
                                 ByVal columnName As String, _
                                 ByVal matchValue As String) As Long
    Dim recordIndex As Long
    Dim columnNumber As Long

    If targetTable Is Nothing Or targetTable.DataBodyRange Is Nothing Then Exit Function
    columnNumber = ColumnIndex(targetTable, columnName)
    If columnNumber = 0 Then Exit Function
    For recordIndex = 1 To targetTable.ListRows.Count
        If StrComp(Trim$(CStr(targetTable.DataBodyRange.Cells(recordIndex, columnNumber).Value)), _
                   Trim$(matchValue), vbTextCompare) = 0 Then
            FindTableRecord = recordIndex
            Exit Function
        End If
    Next recordIndex
End Function

Private Function FirstBlankOrNewRecord(ByVal targetTable As ListObject) As ListRow
    Dim candidate As ListRow
    Dim column As ListColumn
    Dim hasValue As Boolean

    For Each candidate In targetTable.ListRows
        hasValue = False
        For Each column In targetTable.ListColumns
            If Trim$(CStr(candidate.Range.Cells(1, column.Index).Value)) <> "" Then
                hasValue = True
                Exit For
            End If
        Next column
        If Not hasValue Then
            Set FirstBlankOrNewRecord = candidate
            Exit Function
        End If
    Next candidate
    Set FirstBlankOrNewRecord = targetTable.ListRows.Add
End Function

Private Function ColumnIndex(ByVal targetTable As ListObject, _
                             ByVal columnName As String) As Long
    Dim column As ListColumn

    If targetTable Is Nothing Then Exit Function
    For Each column In targetTable.ListColumns
        If StrComp(Trim$(column.Name), Trim$(columnName), vbTextCompare) = 0 Then
            ColumnIndex = column.Index
            Exit Function
        End If
    Next column
End Function

Private Function CellText(ByVal targetTable As ListObject, _
                          ByVal recordIndex As Long, _
                          ByVal columnName As String) As String
    Dim columnNumber As Long
    Dim valueIn As Variant

    columnNumber = ColumnIndex(targetTable, columnName)
    If columnNumber = 0 Or targetTable.DataBodyRange Is Nothing Then Exit Function
    valueIn = targetTable.DataBodyRange.Cells(recordIndex, columnNumber).Value
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Then Exit Function
    CellText = Trim$(CStr(valueIn))
End Function

Private Function CellNumber(ByVal targetTable As ListObject, _
                            ByVal recordIndex As Long, _
                            ByVal columnName As String) As Double
    Dim valueIn As String

    valueIn = CellText(targetTable, recordIndex, columnName)
    If IsNumeric(valueIn) Then CellNumber = CDbl(valueIn)
End Function

Private Sub SetCellText(ByVal targetTable As ListObject, _
                        ByVal recordIndex As Long, _
                        ByVal columnName As String, _
                        ByVal valueText As String)
    SetCellValue targetTable, recordIndex, columnName, valueText
End Sub

Private Sub SetCellValue(ByVal targetTable As ListObject, _
                         ByVal recordIndex As Long, _
                         ByVal columnName As String, _
                         ByVal valueIn As Variant)
    Dim columnNumber As Long

    columnNumber = ColumnIndex(targetTable, columnName)
    If columnNumber = 0 Then
        Err.Raise vbObjectError + 7681, "modTS_Received.SetCellValue", _
                  "Receiving column is missing: " & columnName
    End If
    targetTable.DataBodyRange.Cells(recordIndex, columnNumber).Value = valueIn
End Sub

Private Function CanHideWorksheet(ByVal wb As Workbook, _
                                  ByVal sheetToHide As Worksheet) As Boolean
    Dim ws As Worksheet
    Dim visibleCount As Long

    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then visibleCount = visibleCount + 1
    Next ws
    CanHideWorksheet = (sheetToHide.Visible <> xlSheetVisible Or visibleCount > 1)
End Function

Private Sub EnsureReceivingButtons(ByVal ws As Worksheet)
    Dim button As Shape
    Dim anchor As Range

    If ws Is Nothing Then Exit Sub
    DeleteReceivingActionButton ws, "btnUndoMacro"
    DeleteReceivingActionButton ws, "btnRedoMacro"
    On Error Resume Next
    Set button = ws.Shapes("btnConfirmWrites")
    On Error GoTo 0
    Set anchor = ws.Range("C1")
    If button Is Nothing Then
        Set button = ws.Shapes.AddFormControl( _
            xlButtonControl, anchor.Left, 6, 118, 20)
        button.Name = "btnConfirmWrites"
    End If
    button.Left = anchor.Left
    button.Top = 6
    button.Width = 118
    button.Height = 20
    button.TextFrame.Characters.Text = "Confirm Writes"
    button.OnAction = "'" & ThisWorkbook.Name & "'!modTS_Received.ConfirmWrites"
End Sub

Private Sub DeleteReceivingActionButton(ByVal ws As Worksheet, _
                                        ByVal shapeName As String)
    On Error Resume Next
    ws.Shapes(shapeName).Delete
    On Error GoTo 0
End Sub

Private Sub ShowReceivingMessage(ByVal messageText As String, _
                                 ByVal style As VbMsgBoxStyle)
    If modUiQuiet.QuietUiIsActive() Then
        Debug.Print "invSys Receiving: " & messageText
    Else
        MsgBox messageText, style, "invSys Receiving"
    End If
End Sub
