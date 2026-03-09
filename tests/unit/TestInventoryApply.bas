Attribute VB_Name = "TestInventoryApply"
Option Explicit

Public Sub RunInventoryApplyTests()
    Dim passed As Long
    Dim failed As Long

    Tally TestApplyReceive_ValidEvent(), passed, failed
    Tally TestApplyReceive_InvalidSKU(), passed, failed
    Tally TestApplyReceive_Duplicate(), passed, failed
    Tally TestApplyReceive_ProtectedSheetReturnsClearError(), passed, failed

    Debug.Print "InventoryDomain.Apply tests - Passed: " & passed & " Failed: " & failed
End Sub

Public Function TestApplyReceive_ValidEvent() As Long
    Dim wbInv As Workbook
    Dim evt As Object
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim loLog As ListObject
    Dim loApplied As ListObject

    Set wbInv = TestPhase2Helpers.BuildPhase2InventoryWorkbook("WH1", Array("SKU-001"))
    Set evt = TestPhase2Helpers.CreateReceiveEvent("EVT-001", "WH1", "S1", "user1", "SKU-001", 5, "A1", "first receipt")

    On Error GoTo CleanFail
    If Not modInventoryApply.ApplyReceiveEvent(evt, wbInv, "RUN-001", statusOut, errorCode, errorMessage) Then GoTo CleanExit
    If UCase$(statusOut) <> "APPLIED" Then GoTo CleanExit

    Set loLog = wbInv.Worksheets("InventoryLog").ListObjects("tblInventoryLog")
    Set loApplied = wbInv.Worksheets("AppliedEvents").ListObjects("tblAppliedEvents")
    If loLog.ListRows.Count <> 2 Then GoTo CleanExit
    If loApplied.ListRows.Count <> 2 Then GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(loLog, 2, "EventID")) <> "EVT-001" Then GoTo CleanExit
    If CDbl(TestPhase2Helpers.GetRowValue(loLog, 2, "QtyDelta")) <> 5 Then GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(loApplied, 2, "Status")) <> "APPLIED" Then GoTo CleanExit

    TestApplyReceive_ValidEvent = 1

CleanExit:
    TestPhase2Helpers.CloseNoSave wbInv
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestApplyReceive_ProtectedSheetReturnsClearError() As Long
    Dim wbInv As Workbook
    Dim evt As Object
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String

    Set wbInv = TestPhase2Helpers.BuildPhase2InventoryWorkbook("WH1", Array("SKU-001"))
    Set evt = TestPhase2Helpers.CreateReceiveEvent("EVT-004", "WH1", "S1", "user1", "SKU-001", 3)

    wbInv.Worksheets("InventoryLog").Protect Password:="pw"

    On Error GoTo CleanFail
    If modInventoryApply.ApplyReceiveEvent(evt, wbInv, "RUN-001", statusOut, errorCode, errorMessage) Then GoTo CleanExit
    If UCase$(errorCode) <> "APPLY_EXCEPTION" Then GoTo CleanExit
    If InStr(1, errorMessage, "could not be unprotected", vbTextCompare) = 0 Then GoTo CleanExit

    TestApplyReceive_ProtectedSheetReturnsClearError = 1

CleanExit:
    TestPhase2Helpers.CloseNoSave wbInv
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestApplyReceive_InvalidSKU() As Long
    Dim wbInv As Workbook
    Dim evt As Object
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim loLog As ListObject

    Set wbInv = TestPhase2Helpers.BuildPhase2InventoryWorkbook("WH1", Array("SKU-001"))
    Set evt = TestPhase2Helpers.CreateReceiveEvent("EVT-002", "WH1", "S1", "user1", "BAD-SKU", 5)

    On Error GoTo CleanFail
    If modInventoryApply.ApplyReceiveEvent(evt, wbInv, "RUN-001", statusOut, errorCode, errorMessage) Then GoTo CleanExit
    If UCase$(errorCode) <> "INVALID_SKU" Then GoTo CleanExit

    Set loLog = wbInv.Worksheets("InventoryLog").ListObjects("tblInventoryLog")
    If loLog.ListRows.Count <> 1 Then GoTo CleanExit

    TestApplyReceive_InvalidSKU = 1

CleanExit:
    TestPhase2Helpers.CloseNoSave wbInv
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestApplyReceive_Duplicate() As Long
    Dim wbInv As Workbook
    Dim evt As Object
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim loLog As ListObject

    Set wbInv = TestPhase2Helpers.BuildPhase2InventoryWorkbook("WH1", Array("SKU-001"))
    Set evt = TestPhase2Helpers.CreateReceiveEvent("EVT-003", "WH1", "S1", "user1", "SKU-001", 1)

    On Error GoTo CleanFail
    If Not modInventoryApply.ApplyReceiveEvent(evt, wbInv, "RUN-001", statusOut, errorCode, errorMessage) Then GoTo CleanExit
    If Not modInventoryApply.ApplyReceiveEvent(evt, wbInv, "RUN-001", statusOut, errorCode, errorMessage) Then GoTo CleanExit
    If UCase$(statusOut) <> "SKIP_DUP" Then GoTo CleanExit

    Set loLog = wbInv.Worksheets("InventoryLog").ListObjects("tblInventoryLog")
    If loLog.ListRows.Count <> 2 Then GoTo CleanExit

    TestApplyReceive_Duplicate = 1

CleanExit:
    TestPhase2Helpers.CloseNoSave wbInv
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Private Sub Tally(ByVal testResult As Long, ByRef passed As Long, ByRef failed As Long)
    If testResult = 1 Then
        passed = passed + 1
    Else
        failed = failed + 1
    End If
End Sub
