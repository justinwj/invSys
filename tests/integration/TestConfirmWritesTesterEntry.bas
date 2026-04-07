Attribute VB_Name = "TestConfirmWritesTesterEntry"
Option Explicit

Private mLastError As String

Public Function RunConfirmWritesTesterIntegration() As Long
    On Error GoTo FailRun
    mLastError = vbNullString
    RunConfirmWritesTesterIntegration = test_ConfirmWrites_Tester.TestConfirmWrites_Tester_EndToEnd()
    Exit Function

FailRun:
    mLastError = Err.Description
End Function

Public Function GetConfirmWritesTesterIntegrationContext() As String
    GetConfirmWritesTesterIntegrationContext = test_ConfirmWrites_Tester.GetConfirmWritesTesterContextPacked()
    If mLastError <> "" Then GetConfirmWritesTesterIntegrationContext = GetConfirmWritesTesterIntegrationContext & "|Error=" & mLastError
End Function

Public Function GetConfirmWritesTesterIntegrationRows() As String
    GetConfirmWritesTesterIntegrationRows = test_ConfirmWrites_Tester.GetConfirmWritesTesterEvidenceRows()
End Function
