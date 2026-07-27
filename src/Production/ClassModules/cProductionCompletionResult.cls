VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cProductionCompletionResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mSucceeded As Boolean
Private mStatus As String
Private mSessionId As String
Private mConsumeEventId As String
Private mCompleteEventId As String
Private mProcessorVerified As Boolean
Private mRefreshVerified As Boolean
Private mCompensationRequired As Boolean
Private mMessage As String

Public Sub InitializeFromSession(ByVal session As cProductionRunSession, Optional ByVal message As String = "")
    If session Is Nothing Then _
        Err.Raise vbObjectError + 7630, "cProductionCompletionResult.InitializeFromSession", _
                  "Production session is required."

    mSucceeded = session.ReadyForNextBatch
    mStatus = session.Status
    mSessionId = session.SessionId
    mConsumeEventId = session.ConsumeEventId
    mCompleteEventId = session.CompleteEventId
    mProcessorVerified = session.ProcessorVerified
    mRefreshVerified = session.RefreshVerified
    mCompensationRequired = session.CompensationRequired
    mMessage = message
    If mMessage = "" Then mMessage = session.FailureMessage
End Sub

Public Function ToEnvelope() As String
    ToEnvelope = IIf(mSucceeded, "OK", "FAIL") & _
        "|Status=" & mStatus & _
        "|SessionId=" & mSessionId & _
        "|ConsumeEventId=" & mConsumeEventId & _
        "|CompleteEventId=" & mCompleteEventId & _
        "|ProcessorVerified=" & IIf(mProcessorVerified, "TRUE", "FALSE") & _
        "|RefreshVerified=" & IIf(mRefreshVerified, "TRUE", "FALSE") & _
        "|CompensationRequired=" & IIf(mCompensationRequired, "TRUE", "FALSE") & _
        "|Message=" & Replace$(Replace$(mMessage, "|", "/"), vbCrLf, " ")
End Function

Public Property Get Succeeded() As Boolean
    Succeeded = mSucceeded
End Property

Public Property Get Status() As String
    Status = CStr(mStatus)
End Property

Public Property Get SessionId() As String
    SessionId = CStr(mSessionId)
End Property

Public Property Get ConsumeEventId() As String
    ConsumeEventId = CStr(mConsumeEventId)
End Property

Public Property Get CompleteEventId() As String
    CompleteEventId = CStr(mCompleteEventId)
End Property

Public Property Get ProcessorVerified() As Boolean
    ProcessorVerified = CBool(mProcessorVerified)
End Property

Public Property Get RefreshVerified() As Boolean
    RefreshVerified = CBool(mRefreshVerified)
End Property

Public Property Get CompensationRequired() As Boolean
    CompensationRequired = CBool(mCompensationRequired)
End Property

Public Property Get Message() As String
    Message = mMessage
End Property
