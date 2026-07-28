Attribute VB_Name = "modReceivingPostingService"
Option Explicit

Private Const SHEET_RECEIVING As String = "ReceivedTally"
Private Const SHEET_RECEIVED_LOG As String = "ReceivedLog"
Private Const TABLE_STAGING As String = "ReceivedTally"
Private Const TABLE_AGGREGATE As String = "AggregateReceived"
Private Const TABLE_LOG As String = "ReceivedLog"

Public Function ExecuteConfirmWrites(ByVal operatorWb As Workbook, _
                                     Optional ByRef report As String = "") As Boolean
    On Error GoTo Failed

    Dim aggregateTable As ListObject
    Dim stagingTable As ListObject
    Dim logTable As ListObject
    Dim states As Collection
    Dim state As cReceivingWorkflowState
    Dim rowIndex As Long
    Dim userId As String
    Dim queueError As String
    Dim runtimeReport As String
    Dim queuedWorkHandled As Boolean
    Dim queuedCount As Long

    If operatorWb Is Nothing Then
        report = "Receiving operator workbook was not provided."
        Exit Function
    End If
    If Not WorkbookIsOpenReceivingService(operatorWb) Then
        report = "The captured Receiving operator workbook is no longer open."
        Exit Function
    End If
    If Not modRoleUiAccess.CanCurrentUserPerformCapability( _
        "RECEIVE_POST", "", "", "", report) Then Exit Function

    Set stagingTable = FindTableReceivingService(operatorWb, TABLE_STAGING)
    Set aggregateTable = FindTableReceivingService(operatorWb, TABLE_AGGREGATE)
    Set logTable = FindTableReceivingService(operatorWb, TABLE_LOG)
    If stagingTable Is Nothing Or aggregateTable Is Nothing Or logTable Is Nothing Then
        report = "Receiving staging or log tables are missing."
        Exit Function
    End If
    If aggregateTable.DataBodyRange Is Nothing Then
        report = "AggregateReceived has no rows to confirm."
        Exit Function
    End If

    userId = Trim$(modRoleEventWriter.ResolveCurrentUserId())
    If userId = "" Then
        report = "Unable to resolve current user identity."
        Exit Function
    End If

    Set states = BuildValidatedStates(aggregateTable, report)
    If states Is Nothing Then Exit Function

    For rowIndex = 1 To states.Count
        Set state = states(rowIndex)
        If StrComp(state.CurrentState, state.StateValidated, vbBinaryCompare) = 0 Then
            queueError = ""
            If Not QueueReceivingState(state, userId, queueError) Then
                report = "Inbox queue failed for staged item " & CStr(rowIndex) & _
                         ": " & queueError
                Exit Function
            End If
            state.MarkSubmitted
            WriteWorkflowState aggregateTable, rowIndex, state
            WriteWorkflowStateBySystemKey stagingTable, state
            queuedCount = queuedCount + 1
        End If
    Next rowIndex

    If Not modOperationsPrimitiveBridge.RunBatchAndRefreshOperatorWorkbook( _
        operatorWb.Name, "", "LOCAL", runtimeReport, False, queuedWorkHandled) Then
        report = "Receiving rows remain submitted; processor application or snapshot refresh " & _
                 "did not complete. " & runtimeReport
        Exit Function
    End If

    For rowIndex = 1 To states.Count
        Set state = states(rowIndex)
        If StrComp(state.CurrentState, state.StateSubmitted, vbBinaryCompare) = 0 Then
            state.MarkProcessorApplied
            WriteWorkflowState aggregateTable, rowIndex, state
            WriteWorkflowStateBySystemKey stagingTable, state
        End If
        If StrComp(state.CurrentState, state.StateProcessorApplied, vbBinaryCompare) = 0 Then
            state.MarkSnapshotRefreshed
            WriteWorkflowState aggregateTable, rowIndex, state
            WriteWorkflowStateBySystemKey stagingTable, state
        End If
        AppendReceivedLog logTable, aggregateTable, rowIndex, state, userId
    Next rowIndex

    ClearReceivingStaging stagingTable, aggregateTable
    For Each state In states
        If StrComp(state.CurrentState, state.StateSnapshotRefreshed, vbBinaryCompare) = 0 Then
            state.MarkReady
        End If
    Next state

    report = "OK|State=READY|Lines=" & CStr(states.Count) & _
             "|Queued=" & CStr(queuedCount) & _
             "|ProcessorApplied=True|SnapshotRefreshed=True|StagingCleared=True"
    ExecuteConfirmWrites = True
    Exit Function

Failed:
    report = "Receiving Confirm Writes failed: " & Err.Description
End Function

Public Sub ClearReceivingStaging(ByVal stagingTable As ListObject, _
                                 ByVal aggregateTable As ListObject)
    If Not stagingTable Is Nothing Then
        If Not stagingTable.DataBodyRange Is Nothing Then stagingTable.DataBodyRange.Delete
    End If
    If Not aggregateTable Is Nothing Then
        If Not aggregateTable.DataBodyRange Is Nothing Then aggregateTable.DataBodyRange.Delete
    End If
End Sub

Public Function ReceivingWorkflowContractProbe(ByVal contractName As String) As String
    On Error GoTo Failed

    Dim state As cReceivingWorkflowState
    Dim firstEventId As String

    Set state = New cReceivingWorkflowState
    Select Case UCase$(Trim$(contractName))
        Case "SEQUENCE"
            state.Initialize "SYS-RECEIVING-PROBE", "", "SKU-PROBE", 2, "DOCK", "probe"
            state.MarkValidated
            state.MarkSubmitted
            state.MarkProcessorApplied
            state.MarkSnapshotRefreshed
            state.MarkReady
            ReceivingWorkflowContractProbe = "OK|State=" & state.CurrentState
        Case "IDENTITY"
            state.Initialize "SYS-RECEIVING-PROBE", "", "SKU-PROBE", 2, "DOCK", "probe"
            state.EnsureEventIdentities
            firstEventId = state.EventId
            state.EnsureEventIdentities
            If firstEventId = "" Or StrComp(firstEventId, state.EventId, vbBinaryCompare) <> 0 Then
                ReceivingWorkflowContractProbe = "FAIL|Event identity changed."
            Else
                ReceivingWorkflowContractProbe = _
                    "OK|System_Key=" & state.SystemKey & "|EventIdStable=True"
            End If
        Case "MISSING_SYSTEM_KEY"
            state.Initialize "", "", "SKU-PROBE", 2, "DOCK", "probe"
            state.EnsureEventIdentities
            ReceivingWorkflowContractProbe = "FAIL|Blank System_Key was accepted."
        Case Else
            ReceivingWorkflowContractProbe = "FAIL|UnknownContract=" & contractName
    End Select
    Exit Function

Failed:
    If UCase$(Trim$(contractName)) = "MISSING_SYSTEM_KEY" Then
        ReceivingWorkflowContractProbe = "OK|Rejected=True|Message=" & Err.Description
    Else
        ReceivingWorkflowContractProbe = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
    End If
End Function

Private Function BuildValidatedStates(ByVal aggregateTable As ListObject, _
                                      ByRef report As String) As Collection
    On Error GoTo Failed

    Dim states As Collection
    Dim state As cReceivingWorkflowState
    Dim rowIndex As Long
    Dim stateValue As String

    If Not RequiredColumnsPresent(aggregateTable, _
        Array("REF_NUMBER", "ITEM_CODE", "ITEM", "UOM", "QUANTITY", "LOCATION", _
              "System_Key", "EventId", "WorkflowState")) Then
        report = "AggregateReceived is missing required Receiving workflow columns."
        Exit Function
    End If

    Set states = New Collection
    For rowIndex = 1 To aggregateTable.ListRows.Count
        Set state = New cReceivingWorkflowState
        stateValue = CellText(aggregateTable, rowIndex, "WorkflowState")
        state.Initialize _
            CellText(aggregateTable, rowIndex, "System_Key"), _
            CellText(aggregateTable, rowIndex, "EventId"), _
            CellText(aggregateTable, rowIndex, "ITEM_CODE"), _
            CellNumber(aggregateTable, rowIndex, "QUANTITY"), _
            CellText(aggregateTable, rowIndex, "LOCATION"), _
            BuildEventNote(aggregateTable, rowIndex), _
            "GOOD", "", stateValue

        Select Case state.CurrentState
            Case state.StateStaged
                state.MarkValidated
            Case state.StateValidated, state.StateSubmitted, _
                 state.StateProcessorApplied, state.StateSnapshotRefreshed
                state.EnsureEventIdentities
            Case Else
                report = "AggregateReceived contains an invalid completed workflow row."
                Exit Function
        End Select
        WriteWorkflowState aggregateTable, rowIndex, state
        states.Add state
    Next rowIndex
    Set BuildValidatedStates = states
    Exit Function

Failed:
    report = "Receiving validation failed: " & Err.Description
End Function

Private Function QueueReceivingState(ByVal state As cReceivingWorkflowState, _
                                     ByVal userId As String, _
                                     ByRef errorMessage As String) As Boolean
    Dim eventId As String
    Dim systemKey As String

    eventId = state.EventId
    systemKey = state.SystemKey
    QueueReceivingState = modRoleEventWriter.QueueReceiveEventServer( _
        "", "", userId, state.Sku, state.Qty, state.Location, state.Note, _
        eventId, errorMessage, "", systemKey, state.ConditionValue, _
        state.AttributesJson)
    If QueueReceivingState Then
        If StrComp(eventId, state.EventId, vbBinaryCompare) <> 0 _
           Or StrComp(systemKey, state.SystemKey, vbBinaryCompare) <> 0 Then
            Err.Raise vbObjectError + 7670, "modReceivingPostingService.QueueReceivingState", _
                      "Receiving queue changed a preassigned event or System_Key identity."
        End If
    End If
End Function

Private Sub WriteWorkflowState(ByVal targetTable As ListObject, _
                               ByVal rowIndex As Long, _
                               ByVal state As cReceivingWorkflowState)
    SetCellText targetTable, rowIndex, "System_Key", state.SystemKey
    SetCellText targetTable, rowIndex, "EventId", state.EventId
    SetCellText targetTable, rowIndex, "WorkflowState", state.CurrentState
End Sub

Private Sub WriteWorkflowStateBySystemKey(ByVal targetTable As ListObject, _
                                          ByVal state As cReceivingWorkflowState)
    Dim rowIndex As Long

    rowIndex = FindTableRow(targetTable, "System_Key", state.SystemKey)
    If rowIndex = 0 Then Exit Sub
    WriteWorkflowState targetTable, rowIndex, state
End Sub

Private Sub AppendReceivedLog(ByVal logTable As ListObject, _
                              ByVal aggregateTable As ListObject, _
                              ByVal aggregateIndex As Long, _
                              ByVal state As cReceivingWorkflowState, _
                              ByVal userId As String)
    Dim targetRow As ListRow
    Dim logIndex As Long

    If FindTableRow(logTable, "EventId", state.EventId) > 0 Then Exit Sub
    Set targetRow = FirstBlankOrNewRow(logTable)
    logIndex = targetRow.Index
    SetCellText logTable, logIndex, "SNAPSHOT_ID", state.EventId
    SetCellValue logTable, logIndex, "ENTRY_DATE", Now
    SetCellText logTable, logIndex, "USER", userId
    SetCellText logTable, logIndex, "REF_NUMBER", _
                CellText(aggregateTable, aggregateIndex, "REF_NUMBER")
    SetCellText logTable, logIndex, "ITEMS", _
                CellText(aggregateTable, aggregateIndex, "ITEM")
    SetCellValue logTable, logIndex, "QUANTITY", state.Qty
    SetCellText logTable, logIndex, "UOM", _
                CellText(aggregateTable, aggregateIndex, "UOM")
    SetCellText logTable, logIndex, "VENDOR", _
                CellText(aggregateTable, aggregateIndex, "VENDORS")
    SetCellText logTable, logIndex, "LOCATION", state.Location
    SetCellText logTable, logIndex, "ITEM_CODE", state.Sku
    SetCellText logTable, logIndex, "System_Key", state.SystemKey
    SetCellText logTable, logIndex, "EventId", state.EventId
End Sub

Private Function BuildEventNote(ByVal aggregateTable As ListObject, _
                                ByVal rowIndex As Long) As String
    BuildEventNote = "REF_NUMBER=" & CellText(aggregateTable, rowIndex, "REF_NUMBER")
    If CellText(aggregateTable, rowIndex, "ITEM") <> "" Then
        BuildEventNote = BuildEventNote & _
                         "; ITEM=" & CellText(aggregateTable, rowIndex, "ITEM")
    End If
    If CellText(aggregateTable, rowIndex, "VENDORS") <> "" Then
        BuildEventNote = BuildEventNote & _
                         "; VENDORS=" & CellText(aggregateTable, rowIndex, "VENDORS")
    End If
End Function

Private Function FindTableReceivingService(ByVal wb As Workbook, _
                                           ByVal tableName As String) As ListObject
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function
    On Error Resume Next
    If StrComp(tableName, TABLE_LOG, vbTextCompare) = 0 Then
        Set ws = wb.Worksheets(SHEET_RECEIVED_LOG)
    Else
        Set ws = wb.Worksheets(SHEET_RECEIVING)
    End If
    If Not ws Is Nothing Then
        Set FindTableReceivingService = ws.ListObjects(tableName)
    End If
    On Error GoTo 0
End Function

Private Function RequiredColumnsPresent(ByVal targetTable As ListObject, _
                                        ByVal names As Variant) As Boolean
    Dim nameValue As Variant
    Dim requiredColumn As ListColumn

    For Each nameValue In names
        Set requiredColumn = FindColumnReceivingService(targetTable, CStr(nameValue))
        If requiredColumn Is Nothing Then Exit Function
    Next nameValue
    RequiredColumnsPresent = True
End Function

Private Function FindColumnReceivingService(ByVal targetTable As ListObject, _
                                            ByVal columnName As String) As ListColumn
    Dim column As ListColumn

    If targetTable Is Nothing Then Exit Function
    For Each column In targetTable.ListColumns
        If StrComp(Trim$(column.Name), Trim$(columnName), vbTextCompare) = 0 Then
            Set FindColumnReceivingService = column
            Exit Function
        End If
    Next column
End Function

Private Function CellText(ByVal targetTable As ListObject, _
                          ByVal rowIndex As Long, _
                          ByVal columnName As String) As String
    Dim valueIn As Variant
    Dim targetColumn As ListColumn

    Set targetColumn = FindColumnReceivingService(targetTable, columnName)
    If targetColumn Is Nothing Then Exit Function
    valueIn = targetTable.DataBodyRange.Cells(rowIndex, targetColumn.Index).Value
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Then Exit Function
    CellText = Trim$(CStr(valueIn))
End Function

Private Function CellNumber(ByVal targetTable As ListObject, _
                            ByVal rowIndex As Long, _
                            ByVal columnName As String) As Double
    Dim valueIn As Variant
    Dim targetColumn As ListColumn

    Set targetColumn = FindColumnReceivingService(targetTable, columnName)
    If targetColumn Is Nothing Then Exit Function
    valueIn = targetTable.DataBodyRange.Cells(rowIndex, targetColumn.Index).Value
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Then Exit Function
    If IsNumeric(valueIn) Then CellNumber = CDbl(valueIn)
End Function

Private Sub SetCellText(ByVal targetTable As ListObject, _
                        ByVal rowIndex As Long, _
                        ByVal columnName As String, _
                        ByVal valueText As String)
    SetCellValue targetTable, rowIndex, columnName, valueText
End Sub

Private Sub SetCellValue(ByVal targetTable As ListObject, _
                         ByVal rowIndex As Long, _
                         ByVal columnName As String, _
                         ByVal valueIn As Variant)
    Dim targetColumn As ListColumn

    Set targetColumn = FindColumnReceivingService(targetTable, columnName)
    If targetColumn Is Nothing Then
        Err.Raise vbObjectError + 7671, "modReceivingPostingService.SetCellValue", _
                  "Receiving column is missing: " & columnName
    End If
    targetTable.DataBodyRange.Cells(rowIndex, targetColumn.Index).Value = valueIn
End Sub

Private Function FindTableRow(ByVal targetTable As ListObject, _
                              ByVal columnName As String, _
                              ByVal matchValue As String) As Long
    Dim rowIndex As Long
    Dim targetColumn As ListColumn
    Dim valueIn As Variant

    If targetTable Is Nothing Or targetTable.DataBodyRange Is Nothing Then Exit Function
    Set targetColumn = FindColumnReceivingService(targetTable, columnName)
    If targetColumn Is Nothing Then Exit Function
    For rowIndex = 1 To targetTable.ListRows.Count
        valueIn = targetTable.DataBodyRange.Cells(rowIndex, targetColumn.Index).Value
        If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Then GoTo ContinueRow
        If StrComp(Trim$(CStr(valueIn)), _
                   Trim$(matchValue), vbBinaryCompare) = 0 Then
            FindTableRow = rowIndex
            Exit Function
        End If
ContinueRow:
    Next rowIndex
End Function

Private Function FirstBlankOrNewRow(ByVal targetTable As ListObject) As ListRow
    Dim candidate As ListRow

    If targetTable.DataBodyRange Is Nothing Then
        Set FirstBlankOrNewRow = targetTable.ListRows.Add
        Exit Function
    End If
    For Each candidate In targetTable.ListRows
        If Application.WorksheetFunction.CountA(candidate.Range) = 0 Then
            Set FirstBlankOrNewRow = candidate
            Exit Function
        End If
    Next candidate
    Set FirstBlankOrNewRow = targetTable.ListRows.Add
End Function

Private Function WorkbookIsOpenReceivingService(ByVal wb As Workbook) As Boolean
    Dim candidate As Workbook

    If wb Is Nothing Then Exit Function
    For Each candidate In Application.Workbooks
        If candidate Is wb Then
            WorkbookIsOpenReceivingService = True
            Exit Function
        End If
    Next candidate
End Function
