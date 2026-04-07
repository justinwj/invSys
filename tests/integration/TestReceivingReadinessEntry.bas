Attribute VB_Name = "TestReceivingReadinessEntry"
Option Explicit

Private mLastError As String

Public Function RunReceivingReadinessIntegration() As Long
    On Error GoTo FailRun
    mLastError = vbNullString
    RunReceivingReadinessIntegration = test_ReceivingReadiness.TestReceivingReadiness_StatusPanelRendersForKnownBadWorkbook()
    Exit Function

FailRun:
    mLastError = Err.Description
End Function

Public Function GetReceivingReadinessIntegrationContext() As String
    GetReceivingReadinessIntegrationContext = test_ReceivingReadiness.GetReceivingReadinessContextPacked()
    If mLastError <> "" Then GetReceivingReadinessIntegrationContext = GetReceivingReadinessIntegrationContext & "|Error=" & mLastError
End Function

Public Function GetReceivingReadinessIntegrationRows() As String
    GetReceivingReadinessIntegrationRows = test_ReceivingReadiness.GetReceivingReadinessEvidenceRows()
End Function
