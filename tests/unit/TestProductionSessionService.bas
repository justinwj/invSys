Attribute VB_Name = "TestProductionSessionService"
Option Explicit

Private mLastTestFailure As String

Public Sub ClearLastTestFailure()
    mLastTestFailure = ""
End Sub

Public Function GetLastTestFailure() As String
    GetLastTestFailure = mLastTestFailure
End Function

Public Function TestProductionSession_AllocatesImmutableSystemKeys() As Long
    TestProductionSession_AllocatesImmutableSystemKeys = ExpectProbe( _
        "IDENTITY", _
        "OK|InputCount=2|InputKeysPreserved=TRUE|OutputKeyNonblank=TRUE|OutputKeyStable=TRUE")
End Function

Public Function TestProductionSession_RejectsBlankAndDuplicateInputKeys() As Long
    TestProductionSession_RejectsBlankAndDuplicateInputKeys = ExpectProbe( _
        "INVALID_INPUT_KEYS", _
        "OK|BlankRejected=TRUE|DuplicateRejected=TRUE")
End Function

Public Function TestProductionSession_AssignsExplicitConsumeAndCompleteEventIds() As Long
    TestProductionSession_AssignsExplicitConsumeAndCompleteEventIds = ExpectProbe( _
        "EVENT_IDENTITIES", _
        "OK|ConsumeNonblank=TRUE|CompleteNonblank=TRUE|Distinct=TRUE|Stable=TRUE")
End Function

Public Function TestProductionSession_BecomesReadyOnlyAfterProcessorAndRefresh() As Long
    TestProductionSession_BecomesReadyOnlyAfterProcessorAndRefresh = ExpectProbe( _
        "READY_SEQUENCE", _
        "OK|BeforeProcessor=FALSE|BeforeRefresh=FALSE|AfterRefresh=TRUE")
End Function

Public Function TestProductionSession_RecordsCompensationOnlyAfterConsumeApplied() As Long
    TestProductionSession_RecordsCompensationOnlyAfterConsumeApplied = ExpectProbe( _
        "COMPENSATION", _
        "OK|BeforeConsume=FALSE|AfterConsume=TRUE|AfterComplete=FALSE")
End Function

Public Function TestProductionSession_RoundTripsRestartState() As Long
    TestProductionSession_RoundTripsRestartState = ExpectProbe( _
        "RESTART", _
        "OK|StatePreserved=TRUE|OutputKeyPreserved=TRUE|EventIdsPreserved=TRUE")
End Function

Public Function TestProductionCompletionResult_IsStructuredAndSerializable() As Long
    TestProductionCompletionResult_IsStructuredAndSerializable = ExpectProbe( _
        "RESULT_ENVELOPE", _
        "OK|Status=READY|ProcessorVerified=TRUE|RefreshVerified=TRUE|CompensationRequired=FALSE")
End Function

Public Function TestProductionCheckIn_StagesBySystemKeyWithoutMutatingInventory() As Long
    Dim wb As Workbook
    Dim loInv As ListObject
    Dim loCheck As ListObject
    Dim lr As ListRow
    Dim report As String
    Dim result As String
    Dim resultParts As Variant
    Dim stagedText As String

    On Error GoTo TestFailed
    Set wb = Application.Workbooks.Add
    If Not modRoleWorkbookSurfaces.EnsureProductionWorkbookSurface(wb, report) Then _
        Err.Raise vbObjectError + 7655, "TestProductionCheckIn_StagesBySystemKeyWithoutMutatingInventory", report
    Set loInv = FindTableForSessionTest(wb, "invSys")
    Set loCheck = FindTableForSessionTest(wb, "Prod_invSys_Check")
    If loInv Is Nothing Or loCheck Is Nothing Then _
        Err.Raise vbObjectError + 7656, "TestProductionCheckIn_StagesBySystemKeyWithoutMutatingInventory", _
                  "Production inventory or Check In staging table was not created."
    If Not loInv.DataBodyRange Is Nothing Then loInv.DataBodyRange.Delete
    If Not loCheck.DataBodyRange Is Nothing Then loCheck.DataBodyRange.Delete
    Set lr = loInv.ListRows.Add
    SetTableValueForSessionTest loInv, lr.Index, "System_Key", "SYS-CHECKIN-INPUT"
    SetTableValueForSessionTest loInv, lr.Index, "ITEM_CODE", "SKU-CHECKIN"
    SetTableValueForSessionTest loInv, lr.Index, "ITEM", "Check In Input"
    SetTableValueForSessionTest loInv, lr.Index, "LOCATION", "A1"
    SetTableValueForSessionTest loInv, lr.Index, "USED", 0
    SetTableValueForSessionTest loInv, lr.Index, "TOTAL INV", 12

    result = mProduction.TestProductionSystemKeyPayloadStagesWithoutInventoryMutation( _
        loInv, loCheck, "SYS-CHECKIN-INPUT", "SKU-CHECKIN", 2)
    resultParts = Split(result, "|")
    If UBound(resultParts) >= 2 Then stagedText = Replace$(CStr(resultParts(2)), "Staged=", "")
    If UBound(resultParts) = 3 _
       And CStr(resultParts(0)) = "OK" _
       And CStr(resultParts(1)) = "SystemKey=SYS-CHECKIN-INPUT" _
       And IsNumeric(stagedText) _
       And Abs(CDbl(stagedText) - 2) < 0.0000001 _
       And CStr(resultParts(3)) = "InventoryUsedUnchanged=TRUE" Then
        TestProductionCheckIn_StagesBySystemKeyWithoutMutatingInventory = 1
    Else
        mLastTestFailure = result
    End If

CleanExit:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    On Error GoTo 0
    Exit Function

TestFailed:
    mLastTestFailure = Err.Description
    Resume CleanExit
End Function

Public Function TestProductionSession_PersistsThroughWorkbookCloseReopen() As Long
    Dim wb As Workbook
    Dim reopened As Workbook
    Dim session As cProductionRunSession
    Dim restored As cProductionRunSession
    Dim report As String
    Dim tempPath As String

    On Error GoTo TestFailed
    tempPath = Environ$("TEMP") & "\invSys_slice7_session_" & _
               Format$(Now, "yyyymmdd_hhnnss") & "_" & _
               Right$("000000" & CStr(CLng(Timer * 100)), 6) & ".xlsm"

    Set wb = Application.Workbooks.Add
    Set session = modProductionCompletionService.CreateProductionSession( _
        "SESSION-WORKBOOK-RESTART", "DESIGN-RESTART", "3", 7, "A1")
    session.AddInputAllocation "SYS-RESTART-INPUT", 7, "SKU-INPUT", "A1"
    session.EnsureOutputIdentity "SKU-OUTPUT", 7, "A1", "GOOD", "{}", "SYS-RESTART-OUTPUT"
    session.EnsureEventIdentities "EVT-RESTART-CONSUME", "EVT-RESTART-COMPLETE"
    session.MarkConsumeQueued
    session.MarkCompleteQueued
    session.RecordProcessorResult True, True, True

    If Not modProductionCompletionService.SaveProductionSessionToWorkbook(wb, session, report) Then _
        Err.Raise vbObjectError + 7650, "TestProductionSession_PersistsThroughWorkbookCloseReopen", report
    wb.SaveAs tempPath, 52
    wb.Close False
    Set wb = Nothing

    Set reopened = Application.Workbooks.Open(tempPath, UpdateLinks:=False, ReadOnly:=False)
    Set restored = modProductionCompletionService.LoadProductionSessionFromWorkbook(reopened, report)
    If restored Is Nothing Then _
        Err.Raise vbObjectError + 7651, "TestProductionSession_PersistsThroughWorkbookCloseReopen", report

    If restored.SessionId = "SESSION-WORKBOOK-RESTART" _
       And restored.OutputSystemKey = "SYS-RESTART-OUTPUT" _
       And restored.ConsumeEventId = "EVT-RESTART-CONSUME" _
       And restored.CompleteEventId = "EVT-RESTART-COMPLETE" _
       And restored.ProcessorVerified _
       And Not restored.RefreshVerified _
       And Not restored.ReadyForNextBatch Then
        TestProductionSession_PersistsThroughWorkbookCloseReopen = 1
    Else
        mLastTestFailure = "Persisted Production session did not survive workbook close/reopen."
    End If

CleanExit:
    On Error Resume Next
    If Not reopened Is Nothing Then reopened.Close False
    If Not wb Is Nothing Then wb.Close False
    If tempPath <> "" Then
        If Len(Dir$(tempPath)) > 0 Then Kill tempPath
    End If
    On Error GoTo 0
    Exit Function

TestFailed:
    mLastTestFailure = Err.Description
    Resume CleanExit
End Function

Private Function ExpectProbe(ByVal contractName As String, ByVal expectedResult As String) As Long
    Dim actualResult As String

    On Error GoTo ProbeMissing
    actualResult = CStr(Application.Run( _
        "modProductionCompletionService.ProductionSessionContractProbe", _
        contractName))
    On Error GoTo 0

    If StrComp(actualResult, expectedResult, vbBinaryCompare) = 0 Then
        ExpectProbe = 1
    Else
        mLastTestFailure = contractName & " expected '" & expectedResult & _
            "' but received '" & actualResult & "'."
    End If
    Exit Function

ProbeMissing:
    mLastTestFailure = contractName & _
        " completion-service contract is unavailable: " & Err.Description
End Function

Private Function FindTableForSessionTest(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In wb.Worksheets
        Set lo = Nothing
        On Error Resume Next
        Set lo = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not lo Is Nothing Then
            Set FindTableForSessionTest = lo
            Exit Function
        End If
    Next ws
End Function

Private Sub SetTableValueForSessionTest(ByVal lo As ListObject, ByVal rowIndex As Long, _
                                        ByVal headerName As String, ByVal valueOut As Variant)
    Dim lc As ListColumn

    For Each lc In lo.ListColumns
        If StrComp(Trim$(lc.Name), Trim$(headerName), vbTextCompare) = 0 Then
            lo.DataBodyRange.Cells(rowIndex, lc.Index).Value = valueOut
            Exit Sub
        End If
    Next lc
    Err.Raise vbObjectError + 7657, "SetTableValueForSessionTest", _
              "Column '" & headerName & "' was not found in table '" & lo.Name & "'."
End Sub
