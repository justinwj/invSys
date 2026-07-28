VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cShippingWorkflowState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const STATE_STAGED As String = "STAGED"
Private Const STATE_RESERVED As String = "RESERVED"
Private Const STATE_SUBMITTED As String = "SUBMITTED"
Private Const STATE_APPLIED As String = "APPLIED"
Private Const STATE_COMPLETED As String = "COMPLETED"

Private mSystemKey As String
Private mShipmentLineId As String
Private mReservationEventId As String
Private mEventId As String
Private mState As String

Public Sub Initialize(ByVal systemKey As String, _
                      ByVal shipmentLineId As String, _
                      ByVal reservationEventId As String, _
                      ByVal eventId As String, _
                      Optional ByVal restoredState As String = STATE_STAGED)
    mSystemKey = Trim$(systemKey)
    mShipmentLineId = Trim$(shipmentLineId)
    mReservationEventId = Trim$(reservationEventId)
    mEventId = Trim$(eventId)
    mState = NormalizeState(restoredState)
End Sub

Public Sub EnsureEventIdentities()
    If mSystemKey = "" Then
        Err.Raise vbObjectError + 7681, "cShippingWorkflowState.EnsureEventIdentities", _
                  "Shipping requires the managed inventory System_Key."
    End If
    If mShipmentLineId = "" Then mShipmentLineId = modRoleEventWriter.CreateSystemKey()
    If mEventId = "" Then mEventId = modRoleEventWriter.CreateSystemKey()
End Sub

Public Sub MarkReserved(ByVal reservationEventId As String)
    RequireState STATE_STAGED, "MarkReserved"
    EnsureEventIdentities
    reservationEventId = Trim$(reservationEventId)
    If reservationEventId = "" Then
        Err.Raise vbObjectError + 7682, "cShippingWorkflowState.MarkReserved", _
                  "A reservation event identity is required."
    End If
    mReservationEventId = reservationEventId
    mState = STATE_RESERVED
End Sub

Public Sub MarkSubmitted()
    If StrComp(mState, STATE_RESERVED, vbBinaryCompare) <> 0 _
       And StrComp(mState, STATE_STAGED, vbBinaryCompare) <> 0 Then
        Err.Raise vbObjectError + 7683, "cShippingWorkflowState.MarkSubmitted", _
                  "Invalid Shipping workflow transition from " & mState & "."
    End If
    EnsureEventIdentities
    mState = STATE_SUBMITTED
End Sub

Public Sub MarkApplied()
    RequireState STATE_SUBMITTED, "MarkApplied"
    mState = STATE_APPLIED
End Sub

Public Sub MarkCompleted()
    RequireState STATE_APPLIED, "MarkCompleted"
    mState = STATE_COMPLETED
End Sub

Public Function ReleaseExact(ByVal shipmentLineId As String, _
                             ByVal reservationEventId As String) As Boolean
    ReleaseExact = _
        (mShipmentLineId <> "") _
        And (mReservationEventId <> "") _
        And (StrComp(mShipmentLineId, Trim$(shipmentLineId), vbBinaryCompare) = 0) _
        And (StrComp(mReservationEventId, Trim$(reservationEventId), vbBinaryCompare) = 0)
End Function

Public Property Get SystemKey() As String
    SystemKey = mSystemKey
End Property

Public Property Get ShipmentLineId() As String
    ShipmentLineId = mShipmentLineId
End Property

Public Property Get ReservationEventId() As String
    ReservationEventId = mReservationEventId
End Property

Public Property Get EventId() As String
    EventId = mEventId
End Property

Public Property Get CurrentState() As String
    CurrentState = mState
End Property

Public Property Get StateStaged() As String
    StateStaged = STATE_STAGED
End Property

Public Property Get StateReserved() As String
    StateReserved = STATE_RESERVED
End Property

Public Property Get StateSubmitted() As String
    StateSubmitted = STATE_SUBMITTED
End Property

Public Property Get StateApplied() As String
    StateApplied = STATE_APPLIED
End Property

Public Property Get StateCompleted() As String
    StateCompleted = STATE_COMPLETED
End Property

Private Sub RequireState(ByVal expectedState As String, ByVal memberName As String)
    If StrComp(mState, expectedState, vbBinaryCompare) <> 0 Then
        Err.Raise vbObjectError + 7684, "cShippingWorkflowState." & memberName, _
                  "Invalid Shipping workflow transition from " & mState & _
                  "; expected " & expectedState & "."
    End If
End Sub

Private Function NormalizeState(ByVal stateValue As String) As String
    stateValue = UCase$(Trim$(stateValue))
    Select Case stateValue
        Case STATE_STAGED, STATE_RESERVED, STATE_SUBMITTED, STATE_APPLIED, STATE_COMPLETED
            NormalizeState = stateValue
        Case Else
            NormalizeState = STATE_STAGED
    End Select
End Function
