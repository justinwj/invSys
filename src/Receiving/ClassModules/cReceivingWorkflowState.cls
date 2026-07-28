VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cReceivingWorkflowState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const STATE_STAGED As String = "STAGED"
Private Const STATE_VALIDATED As String = "VALIDATED"
Private Const STATE_SUBMITTED As String = "SUBMITTED"
Private Const STATE_PROCESSOR_APPLIED As String = "PROCESSOR_APPLIED"
Private Const STATE_SNAPSHOT_REFRESHED As String = "SNAPSHOT_REFRESHED"
Private Const STATE_READY As String = "READY"

Private mSystemKey As String
Private mEventId As String
Private mSku As String
Private mQty As Double
Private mLocation As String
Private mNote As String
Private mCondition As String
Private mAttributesJson As String
Private mState As String

Public Sub Initialize(ByVal systemKey As String, _
                      ByVal eventId As String, _
                      ByVal sku As String, _
                      ByVal qty As Double, _
                      ByVal location As String, _
                      ByVal note As String, _
                      Optional ByVal conditionValue As String = "GOOD", _
                      Optional ByVal attributesJson As String = "", _
                      Optional ByVal restoredState As String = STATE_STAGED)
    mSystemKey = Trim$(systemKey)
    mEventId = Trim$(eventId)
    mSku = Trim$(sku)
    mQty = qty
    mLocation = Trim$(location)
    mNote = Trim$(note)
    mCondition = UCase$(Trim$(conditionValue))
    If mCondition = "" Then mCondition = "GOOD"
    mAttributesJson = Trim$(attributesJson)
    mState = NormalizeState(restoredState)
End Sub

Public Sub EnsureEventIdentities()
    If mSystemKey = "" Then
        Err.Raise vbObjectError + 7660, "cReceivingWorkflowState.EnsureEventIdentities", _
                  "Receiving System_Key must be generated at the staging creation boundary."
    End If
    If mEventId = "" Then mEventId = modRoleEventWriter.CreateSystemKey()
End Sub

Public Sub MarkValidated()
    RequireState STATE_STAGED, "MarkValidated"
    ValidatePayload
    EnsureEventIdentities
    mState = STATE_VALIDATED
End Sub

Public Sub MarkSubmitted()
    RequireState STATE_VALIDATED, "MarkSubmitted"
    mState = STATE_SUBMITTED
End Sub

Public Sub MarkProcessorApplied()
    RequireState STATE_SUBMITTED, "MarkProcessorApplied"
    mState = STATE_PROCESSOR_APPLIED
End Sub

Public Sub MarkSnapshotRefreshed()
    RequireState STATE_PROCESSOR_APPLIED, "MarkSnapshotRefreshed"
    mState = STATE_SNAPSHOT_REFRESHED
End Sub

Public Sub MarkReady()
    RequireState STATE_SNAPSHOT_REFRESHED, "MarkReady"
    mState = STATE_READY
End Sub

Public Property Get SystemKey() As String
    SystemKey = mSystemKey
End Property

Public Property Get EventId() As String
    EventId = mEventId
End Property

Public Property Get Sku() As String
    Sku = mSku
End Property

Public Property Get Qty() As Double
    Qty = mQty
End Property

Public Property Get Location() As String
    Location = mLocation
End Property

Public Property Get Note() As String
    Note = mNote
End Property

Public Property Get ConditionValue() As String
    ConditionValue = mCondition
End Property

Public Property Get AttributesJson() As String
    AttributesJson = mAttributesJson
End Property

Public Property Get CurrentState() As String
    CurrentState = mState
End Property

Public Property Get StateStaged() As String
    StateStaged = STATE_STAGED
End Property

Public Property Get StateValidated() As String
    StateValidated = STATE_VALIDATED
End Property

Public Property Get StateSubmitted() As String
    StateSubmitted = STATE_SUBMITTED
End Property

Public Property Get StateProcessorApplied() As String
    StateProcessorApplied = STATE_PROCESSOR_APPLIED
End Property

Public Property Get StateSnapshotRefreshed() As String
    StateSnapshotRefreshed = STATE_SNAPSHOT_REFRESHED
End Property

Public Property Get StateReady() As String
    StateReady = STATE_READY
End Property

Private Sub ValidatePayload()
    If mSku = "" Then
        Err.Raise vbObjectError + 7661, "cReceivingWorkflowState.ValidatePayload", _
                  "Receiving SKU is required."
    End If
    If mQty <= 0 Then
        Err.Raise vbObjectError + 7662, "cReceivingWorkflowState.ValidatePayload", _
                  "Receiving quantity must be greater than zero."
    End If
End Sub

Private Sub RequireState(ByVal expectedState As String, ByVal memberName As String)
    If StrComp(mState, expectedState, vbBinaryCompare) <> 0 Then
        Err.Raise vbObjectError + 7663, "cReceivingWorkflowState." & memberName, _
                  "Invalid Receiving workflow transition from " & mState & _
                  "; expected " & expectedState & "."
    End If
End Sub

Private Function NormalizeState(ByVal stateValue As String) As String
    stateValue = UCase$(Trim$(stateValue))
    Select Case stateValue
        Case STATE_STAGED, STATE_VALIDATED, STATE_SUBMITTED, _
             STATE_PROCESSOR_APPLIED, STATE_SNAPSHOT_REFRESHED, STATE_READY
            NormalizeState = stateValue
        Case Else
            NormalizeState = STATE_STAGED
    End Select
End Function
