Attribute VB_Name = "TestReceivingStabilization"
Option Explicit

Public Function TestReceivingWorkflowState_UsesOrderedTransitions() As Long
    Dim result As String

    result = modReceivingPostingService.ReceivingWorkflowContractProbe("SEQUENCE")
    If StrComp(result, "OK|State=READY", vbBinaryCompare) = 0 Then
        TestReceivingWorkflowState_UsesOrderedTransitions = 1
    Else
        Err.Raise vbObjectError + 7690, _
                  "TestReceivingWorkflowState_UsesOrderedTransitions", result
    End If
End Function

Public Function TestReceivingWorkflowState_PreservesEventIdentity() As Long
    Dim result As String

    result = modReceivingPostingService.ReceivingWorkflowContractProbe("IDENTITY")
    If Left$(result, 3) = "OK|" _
       And InStr(1, result, "EventIdStable=True", vbBinaryCompare) > 0 Then
        TestReceivingWorkflowState_PreservesEventIdentity = 1
    Else
        Err.Raise vbObjectError + 7691, _
                  "TestReceivingWorkflowState_PreservesEventIdentity", result
    End If
End Function

Public Function TestReceivingWorkflowState_RejectsMissingSystemKey() As Long
    Dim result As String

    result = modReceivingPostingService.ReceivingWorkflowContractProbe("MISSING_SYSTEM_KEY")
    If Left$(result, 3) = "OK|" _
       And InStr(1, result, "Rejected=True", vbBinaryCompare) > 0 Then
        TestReceivingWorkflowState_RejectsMissingSystemKey = 1
    Else
        Err.Raise vbObjectError + 7692, _
                  "TestReceivingWorkflowState_RejectsMissingSystemKey", result
    End If
End Function

Public Function TestReceivingStage_GeneratesStableDistinctSystemKeys() As Long
    Dim wb As Workbook
    Dim inventoryTable As ListObject
    Dim stagingTable As ListObject
    Dim sourceRecord As ListRow
    Dim report As String
    Dim firstKey As String
    Dim secondKey As String

    Set wb = Application.Workbooks.Add
    On Error GoTo Failed
    If Not modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface(wb, report) Then
        Err.Raise vbObjectError + 7693, _
                  "TestReceivingStage_GeneratesStableDistinctSystemKeys", report
    End If
    Set inventoryTable = FindTableReceivingTest(wb, "invSys")
    Set stagingTable = FindTableReceivingTest(wb, "ReceivedTally")
    Set sourceRecord = FirstBlankOrNewReceivingTest(inventoryTable)
    SetValueReceivingTest inventoryTable, sourceRecord.Index, "System_Key", "SYS-SOURCE-RECEIVING"
    SetValueReceivingTest inventoryTable, sourceRecord.Index, "ITEM_CODE", "SKU-RECEIVING"
    SetValueReceivingTest inventoryTable, sourceRecord.Index, "ITEM", "Receiving Test Item"
    SetValueReceivingTest inventoryTable, sourceRecord.Index, "UOM", "EA"
    SetValueReceivingTest inventoryTable, sourceRecord.Index, "LOCATION", "DOCK"

    If Not modTS_Received.StageReceivingFormItemForWorkbook( _
        wb, "REF-A", "SYS-SOURCE-RECEIVING", "SKU-RECEIVING", 2, report) Then
        Err.Raise vbObjectError + 7694, _
                  "TestReceivingStage_GeneratesStableDistinctSystemKeys", report
    End If
    firstKey = ValueReceivingTest(stagingTable, 1, "System_Key")
    If Not modTS_Received.StageReceivingFormItemForWorkbook( _
        wb, "REF-A", "SYS-SOURCE-RECEIVING", "SKU-RECEIVING", 3, report) Then
        Err.Raise vbObjectError + 7695, _
                  "TestReceivingStage_GeneratesStableDistinctSystemKeys", report
    End If
    If Not modTS_Received.StageReceivingFormItemForWorkbook( _
        wb, "REF-B", "SYS-SOURCE-RECEIVING", "SKU-RECEIVING", 1, report) Then
        Err.Raise vbObjectError + 7696, _
                  "TestReceivingStage_GeneratesStableDistinctSystemKeys", report
    End If
    secondKey = ValueReceivingTest(stagingTable, 2, "System_Key")

    If firstKey = "" Or secondKey = "" _
       Or StrComp(firstKey, "SYS-SOURCE-RECEIVING", vbBinaryCompare) = 0 _
       Or StrComp(firstKey, secondKey, vbBinaryCompare) = 0 _
       Or CDbl(ValueReceivingTest(stagingTable, 1, "QUANTITY")) <> 5 Then
        Err.Raise vbObjectError + 7697, _
                  "TestReceivingStage_GeneratesStableDistinctSystemKeys", _
                  "Receiving creation identity or same-reference merge was incorrect."
    End If
    TestReceivingStage_GeneratesStableDistinctSystemKeys = 1

CleanExit:
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
    Exit Function
Failed:
    Dim failureNumber As Long
    Dim failureDescription As String
    failureNumber = Err.Number
    failureDescription = Err.Description
    Resume CleanFailure
CleanFailure:
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise failureNumber, _
              "TestReceivingStage_GeneratesStableDistinctSystemKeys", _
              failureDescription
End Function

Public Function TestReceivingPurchasingTab_IsVisibleAndReadOnly() As Long
    Dim wb As Workbook
    Dim report As String
    Dim stagingTable As ListObject
    Dim aggregateTable As ListObject
    Dim beforeStaging As Long
    Dim beforeAggregate As Long
    Dim result As String

    Set wb = Application.Workbooks.Add
    On Error GoTo Failed
    If Not modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface(wb, report) Then
        Err.Raise vbObjectError + 7698, _
                  "TestReceivingPurchasingTab_IsVisibleAndReadOnly", report
    End If
    Set stagingTable = FindTableReceivingTest(wb, "ReceivedTally")
    Set aggregateTable = FindTableReceivingTest(wb, "AggregateReceived")
    beforeStaging = TableRecordCountReceivingTest(stagingTable)
    beforeAggregate = TableRecordCountReceivingTest(aggregateTable)

    result = modTS_Received.RunReceivingPurchasingTabContractForTest(wb)
    If Left$(result, 3) <> "OK|" _
       Or InStr(1, result, "Selected=Purchasing", vbBinaryCompare) = 0 _
       Or InStr(1, result, "EnabledPurchasingActions=0", vbBinaryCompare) = 0 _
       Or TableRecordCountReceivingTest(stagingTable) <> beforeStaging _
       Or TableRecordCountReceivingTest(aggregateTable) <> beforeAggregate Then
        Err.Raise vbObjectError + 7699, _
                  "TestReceivingPurchasingTab_IsVisibleAndReadOnly", result
    End If
    TestReceivingPurchasingTab_IsVisibleAndReadOnly = 1

CleanExit:
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
    Exit Function
Failed:
    Dim failureNumber As Long
    Dim failureDescription As String
    failureNumber = Err.Number
    failureDescription = Err.Description
    Resume CleanFailure
CleanFailure:
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise failureNumber, _
              "TestReceivingPurchasingTab_IsVisibleAndReadOnly", _
              failureDescription
End Function

Private Function FindTableReceivingTest(ByVal wb As Workbook, _
                                        ByVal tableName As String) As ListObject
    Dim ws As Worksheet

    For Each ws In wb.Worksheets
        On Error Resume Next
        Set FindTableReceivingTest = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not FindTableReceivingTest Is Nothing Then Exit Function
    Next ws
End Function

Private Function FirstBlankOrNewReceivingTest(ByVal targetTable As ListObject) As ListRow
    If targetTable.DataBodyRange Is Nothing Then
        Set FirstBlankOrNewReceivingTest = targetTable.ListRows.Add
    Else
        Set FirstBlankOrNewReceivingTest = targetTable.ListRows(1)
    End If
End Function

Private Sub SetValueReceivingTest(ByVal targetTable As ListObject, _
                                  ByVal recordIndex As Long, _
                                  ByVal columnName As String, _
                                  ByVal valueIn As Variant)
    targetTable.DataBodyRange.Cells( _
        recordIndex, targetTable.ListColumns(columnName).Index).Value = valueIn
End Sub

Private Function ValueReceivingTest(ByVal targetTable As ListObject, _
                                    ByVal recordIndex As Long, _
                                    ByVal columnName As String) As String
    ValueReceivingTest = CStr(targetTable.DataBodyRange.Cells( _
        recordIndex, targetTable.ListColumns(columnName).Index).Value)
End Function

Private Function TableRecordCountReceivingTest(ByVal targetTable As ListObject) As Long
    Dim recordIndex As Long
    Dim columnIndex As Long

    If targetTable Is Nothing Or targetTable.DataBodyRange Is Nothing Then Exit Function
    For recordIndex = 1 To targetTable.ListRows.Count
        For columnIndex = 1 To targetTable.ListColumns.Count
            If Trim$(CStr(targetTable.DataBodyRange.Cells(recordIndex, columnIndex).Value)) <> "" Then
                TableRecordCountReceivingTest = TableRecordCountReceivingTest + 1
                Exit For
            End If
        Next columnIndex
    Next recordIndex
End Function
