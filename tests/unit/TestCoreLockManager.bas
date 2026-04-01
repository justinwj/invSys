Attribute VB_Name = "TestCoreLockManager"
Option Explicit

Public Sub RunLockManagerTests()
    Dim passed As Long
    Dim failed As Long

    Tally TestAcquireReleaseLock_Lifecycle(), passed, failed
    Tally TestHeartbeat_ExtendsExpiry(), passed, failed

    Debug.Print "Core.LockManager tests - Passed: " & passed & " Failed: " & failed
End Sub

Public Function TestAcquireReleaseLock_Lifecycle() As Long
    Dim wbCfg As Workbook
    Dim wbInv As Workbook
    Dim loLocks As ListObject
    Dim runId As String
    Dim msg As String

    Set wbCfg = TestPhase2Helpers.BuildPhase2ConfigWorkbook("WH1", "S1")
    Set wbInv = TestPhase2Helpers.BuildPhase2InventoryWorkbook("WH1")

    On Error GoTo CleanFail
    If Not modConfig.LoadConfig("WH1", "S1") Then GoTo CleanExit

    Set loLocks = wbInv.Worksheets("Locks").ListObjects("tblLocks")
    If Not modLockManager.AcquireLock("INVENTORY", "WH1", "svc_processor", "S1", wbInv, runId, msg) Then GoTo CleanExit
    If UCase$(CStr(TestPhase2Helpers.GetRowValue(loLocks, 1, "Status"))) <> "HELD" Then GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(loLocks, 1, "RunId")) = "" Then GoTo CleanExit

    If Not modLockManager.ReleaseLock("INVENTORY", runId, wbInv) Then GoTo CleanExit
    If UCase$(CStr(TestPhase2Helpers.GetRowValue(loLocks, 1, "Status"))) <> "EXPIRED" Then GoTo CleanExit

    TestAcquireReleaseLock_Lifecycle = 1

CleanExit:
    TestPhase2Helpers.CloseNoSave wbInv
    TestPhase2Helpers.CloseNoSave wbCfg
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestHeartbeat_ExtendsExpiry() As Long
    Dim wbCfg As Workbook
    Dim wbInv As Workbook
    Dim loLocks As ListObject
    Dim runId As String
    Dim msg As String
    Dim oldExpiry As Date
    Dim newExpiry As Date

    Set wbCfg = TestPhase2Helpers.BuildPhase2ConfigWorkbook("WH1", "S1")
    Set wbInv = TestPhase2Helpers.BuildPhase2InventoryWorkbook("WH1")

    On Error GoTo CleanFail
    If Not modConfig.LoadConfig("WH1", "S1") Then GoTo CleanExit

    Set loLocks = wbInv.Worksheets("Locks").ListObjects("tblLocks")
    If Not modLockManager.AcquireLock("INVENTORY", "WH1", "svc_processor", "S1", wbInv, runId, msg) Then GoTo CleanExit

    oldExpiry = CDate(TestPhase2Helpers.GetRowValue(loLocks, 1, "ExpiresAtUTC"))
    Application.Wait Now + TimeSerial(0, 0, 1)
    If Not modLockManager.UpdateHeartbeat("INVENTORY", runId, wbInv) Then GoTo CleanExit

    newExpiry = CDate(TestPhase2Helpers.GetRowValue(loLocks, 1, "ExpiresAtUTC"))
    If newExpiry <= oldExpiry Then GoTo CleanExit

    TestHeartbeat_ExtendsExpiry = 1

CleanExit:
    TestPhase2Helpers.CloseNoSave wbInv
    TestPhase2Helpers.CloseNoSave wbCfg
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
