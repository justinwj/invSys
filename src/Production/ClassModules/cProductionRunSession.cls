VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cProductionRunSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mInitialized As Boolean
Private mSessionId As String
Private mDesignId As String
Private mDesignVersion As String
Private mQtyPlanned As Double
Private mLocation As String
Private mInputs As Collection
Private mInputIndex As Object
Private mOutputSystemKey As String
Private mOutputSku As String
Private mOutputQty As Double
Private mOutputCondition As String
Private mOutputAttributesJson As String
Private mConsumeEventId As String
Private mCompleteEventId As String
Private mConsumeQueued As Boolean
Private mCompleteQueued As Boolean
Private mConsumeApplied As Boolean
Private mCompleteApplied As Boolean
Private mProcessorVerified As Boolean
Private mRefreshVerified As Boolean
Private mStatus As String
Private mFailureCode As String
Private mFailureMessage As String
Private mCompensationRequired As Boolean

Private Sub Class_Initialize()
    Set mInputs = New Collection
    Set mInputIndex = CreateObject("Scripting.Dictionary")
    mInputIndex.CompareMode = vbTextCompare
    mStatus = "NEW"
    mOutputCondition = "GOOD"
End Sub

Public Sub Initialize(ByVal sessionId As String, _
                      Optional ByVal designId As String = "", _
                      Optional ByVal designVersion As String = "", _
                      Optional ByVal qtyPlanned As Double = 0, _
                      Optional ByVal location As String = "")
    sessionId = Trim$(sessionId)
    If mInitialized Then _
        Err.Raise vbObjectError + 7600, "cProductionRunSession.Initialize", "Production session is already initialized."
    If sessionId = "" Then _
        Err.Raise vbObjectError + 7601, "cProductionRunSession.Initialize", "Production session identity cannot be blank."
    If qtyPlanned < 0 Then _
        Err.Raise vbObjectError + 7602, "cProductionRunSession.Initialize", "Planned quantity cannot be negative."

    mSessionId = sessionId
    mDesignId = Trim$(designId)
    mDesignVersion = Trim$(designVersion)
    mQtyPlanned = qtyPlanned
    mLocation = Trim$(location)
    mInitialized = True
    mStatus = "STAGED"
End Sub

Public Sub AddInputAllocation(ByVal systemKey As String, _
                              ByVal qty As Double, _
                              Optional ByVal sku As String = "", _
                              Optional ByVal location As String = "")
    Dim allocation As Object

    RequireInitialized "AddInputAllocation"
    systemKey = Trim$(systemKey)
    If systemKey = "" Then _
        Err.Raise vbObjectError + 7603, "cProductionRunSession.AddInputAllocation", _
                  "Input allocation System_Key cannot be blank."
    If qty <= 0 Then _
        Err.Raise vbObjectError + 7604, "cProductionRunSession.AddInputAllocation", _
                  "Input allocation quantity must be greater than zero."
    If mInputIndex.Exists(systemKey) Then _
        Err.Raise vbObjectError + 7605, "cProductionRunSession.AddInputAllocation", _
                  "Input allocation System_Key '" & systemKey & "' is duplicated."

    Set allocation = CreateObject("Scripting.Dictionary")
    allocation.CompareMode = vbTextCompare
    allocation("System_Key") = systemKey
    allocation("SKU") = Trim$(sku)
    allocation("Qty") = qty
    allocation("Location") = Trim$(location)
    mInputs.Add allocation
    mInputIndex.Add systemKey, mInputs.Count
    If mStatus = "STAGED" Then mStatus = "VALIDATED"
End Sub

Public Function EnsureOutputIdentity(ByVal sku As String, _
                                     ByVal qty As Double, _
                                     Optional ByVal location As String = "", _
                                     Optional ByVal conditionValue As String = "GOOD", _
                                     Optional ByVal attributesJson As String = "", _
                                     Optional ByVal systemKey As String = "") As String
    RequireInitialized "EnsureOutputIdentity"
    sku = Trim$(sku)
    systemKey = Trim$(systemKey)
    conditionValue = UCase$(Trim$(conditionValue))
    If sku = "" Then _
        Err.Raise vbObjectError + 7606, "cProductionRunSession.EnsureOutputIdentity", _
                  "Output SKU cannot be blank."
    If qty <= 0 Then _
        Err.Raise vbObjectError + 7607, "cProductionRunSession.EnsureOutputIdentity", _
                  "Output quantity must be greater than zero."
    If conditionValue = "" Then conditionValue = "GOOD"

    If mOutputSystemKey <> "" Then
        If StrComp(mOutputSku, sku, vbTextCompare) <> 0 Or Abs(mOutputQty - qty) > 0.0000001 Then _
            Err.Raise vbObjectError + 7608, "cProductionRunSession.EnsureOutputIdentity", _
                      "Allocated output identity is immutable for this Production session."
        EnsureOutputIdentity = mOutputSystemKey
        Exit Function
    End If

    If systemKey = "" Then systemKey = modRoleEventWriter.CreateSystemKey()
    If systemKey = "" Then _
        Err.Raise vbObjectError + 7609, "cProductionRunSession.EnsureOutputIdentity", _
                  "The Production creation boundary did not allocate an output System_Key."
    If mInputIndex.Exists(systemKey) Then _
        Err.Raise vbObjectError + 7610, "cProductionRunSession.EnsureOutputIdentity", _
                  "Output System_Key must be distinct from every consumed inventory identity."

    mOutputSystemKey = systemKey
    mOutputSku = sku
    mOutputQty = qty
    mLocation = Trim$(location)
    mOutputCondition = conditionValue
    mOutputAttributesJson = attributesJson
    If mStatus = "STAGED" Or mStatus = "VALIDATED" Then mStatus = "PREPARED"
    EnsureOutputIdentity = mOutputSystemKey
End Function

Public Sub EnsureEventIdentities(Optional ByVal consumeEventId As String = "", _
                                 Optional ByVal completeEventId As String = "")
    RequireInitialized "EnsureEventIdentities"
    If mOutputSystemKey = "" Then _
        Err.Raise vbObjectError + 7611, "cProductionRunSession.EnsureEventIdentities", _
                  "Allocate the output System_Key before event identities."

    If mConsumeEventId = "" Then
        consumeEventId = Trim$(consumeEventId)
        If consumeEventId = "" Then consumeEventId = modRoleEventWriter.CreateSystemKey()
        mConsumeEventId = consumeEventId
    End If
    If mCompleteEventId = "" Then
        completeEventId = Trim$(completeEventId)
        If completeEventId = "" Then completeEventId = modRoleEventWriter.CreateSystemKey()
        mCompleteEventId = completeEventId
    End If
    If mConsumeEventId = "" Or mCompleteEventId = "" Then _
        Err.Raise vbObjectError + 7612, "cProductionRunSession.EnsureEventIdentities", _
                  "Both Production event identities must be nonblank."
    If StrComp(mConsumeEventId, mCompleteEventId, vbTextCompare) = 0 Then _
        Err.Raise vbObjectError + 7613, "cProductionRunSession.EnsureEventIdentities", _
                  "Consume and completion events require distinct identities."
End Sub

Public Sub MarkConsumeQueued()
    RequireInitialized "MarkConsumeQueued"
    If mConsumeEventId = "" Then _
        Err.Raise vbObjectError + 7614, "cProductionRunSession.MarkConsumeQueued", _
                  "Consume event identity is not allocated."
    mConsumeQueued = True
    mStatus = "SUBMITTED"
End Sub

Public Sub MarkCompleteQueued()
    RequireInitialized "MarkCompleteQueued"
    If Not mConsumeQueued Then _
        Err.Raise vbObjectError + 7615, "cProductionRunSession.MarkCompleteQueued", _
                  "The consume event must be queued before the output event."
    If mCompleteEventId = "" Then _
        Err.Raise vbObjectError + 7616, "cProductionRunSession.MarkCompleteQueued", _
                  "Completion event identity is not allocated."
    mCompleteQueued = True
End Sub

Public Sub RecordProcessorResult(ByVal consumeApplied As Boolean, _
                                 ByVal completeApplied As Boolean, _
                                 ByVal processorVerified As Boolean, _
                                 Optional ByVal failureCode As String = "", _
                                 Optional ByVal failureMessage As String = "")
    RequireInitialized "RecordProcessorResult"
    If completeApplied And Not consumeApplied Then _
        Err.Raise vbObjectError + 7617, "cProductionRunSession.RecordProcessorResult", _
                  "A Production output cannot be applied before its consumption event."

    mConsumeApplied = consumeApplied
    mCompleteApplied = completeApplied
    mProcessorVerified = processorVerified
    If processorVerified And consumeApplied And completeApplied Then
        mStatus = "PROCESSOR_APPLIED"
        mFailureCode = ""
        mFailureMessage = ""
        mCompensationRequired = False
    ElseIf Trim$(failureCode) <> "" Or Trim$(failureMessage) <> "" Then
        RecordFailure failureCode, failureMessage
    End If
    RefreshReadyState
End Sub

Public Sub RecordRefreshResult(ByVal refreshVerified As Boolean, _
                               Optional ByVal failureCode As String = "", _
                               Optional ByVal failureMessage As String = "")
    RequireInitialized "RecordRefreshResult"
    mRefreshVerified = refreshVerified
    If refreshVerified Then
        If mProcessorVerified And mConsumeApplied And mCompleteApplied Then mStatus = "REFRESHED"
    ElseIf Trim$(failureCode) <> "" Or Trim$(failureMessage) <> "" Then
        RecordFailure failureCode, failureMessage
    End If
    RefreshReadyState
End Sub

Public Sub RecordFailure(ByVal failureCode As String, ByVal failureMessage As String)
    RequireInitialized "RecordFailure"
    mFailureCode = Trim$(failureCode)
    mFailureMessage = Trim$(failureMessage)
    mCompensationRequired = (mConsumeApplied And Not mCompleteApplied)
    mStatus = IIf(mCompensationRequired, "COMPENSATION_REQUIRED", "FAILED")
End Sub

Public Sub RestoreOutputIdentity(ByVal systemKey As String, _
                                 ByVal sku As String, _
                                 ByVal qty As Double, _
                                 ByVal location As String, _
                                 ByVal conditionValue As String, _
                                 ByVal attributesJson As String)
    If mOutputSystemKey <> "" Then _
        Err.Raise vbObjectError + 7618, "cProductionRunSession.RestoreOutputIdentity", _
                  "Output identity was already restored."
    Call EnsureOutputIdentity(sku, qty, location, conditionValue, attributesJson, systemKey)
End Sub

Public Sub RestoreEventIdentities(ByVal consumeEventId As String, ByVal completeEventId As String)
    If mConsumeEventId <> "" Or mCompleteEventId <> "" Then _
        Err.Raise vbObjectError + 7619, "cProductionRunSession.RestoreEventIdentities", _
                  "Event identities were already restored."
    EnsureEventIdentities consumeEventId, completeEventId
End Sub

Public Sub RestoreProgress(ByVal consumeQueued As Boolean, _
                           ByVal completeQueued As Boolean, _
                           ByVal consumeApplied As Boolean, _
                           ByVal completeApplied As Boolean, _
                           ByVal processorVerified As Boolean, _
                           ByVal refreshVerified As Boolean, _
                           ByVal statusValue As String, _
                           ByVal failureCode As String, _
                           ByVal failureMessage As String, _
                           ByVal compensationRequired As Boolean)
    RequireInitialized "RestoreProgress"
    mConsumeQueued = consumeQueued
    mCompleteQueued = completeQueued
    mConsumeApplied = consumeApplied
    mCompleteApplied = completeApplied
    mProcessorVerified = processorVerified
    mRefreshVerified = refreshVerified
    mStatus = UCase$(Trim$(statusValue))
    If mStatus = "" Then mStatus = "STAGED"
    mFailureCode = Trim$(failureCode)
    mFailureMessage = Trim$(failureMessage)
    mCompensationRequired = compensationRequired
    RefreshReadyState
End Sub

Public Property Get SessionId() As String
    SessionId = mSessionId
End Property

Public Property Get DesignId() As String
    DesignId = mDesignId
End Property

Public Property Get DesignVersion() As String
    DesignVersion = mDesignVersion
End Property

Public Property Get QtyPlanned() As Double
    QtyPlanned = mQtyPlanned
End Property

Public Property Get Location() As String
    Location = mLocation
End Property

Public Property Get InputCount() As Long
    InputCount = mInputs.Count
End Property

Public Function InputAllocation(ByVal index As Long) As Object
    If index < 1 Or index > mInputs.Count Then _
        Err.Raise vbObjectError + 7620, "cProductionRunSession.InputAllocation", _
                  "Input allocation index is out of range."
    Set InputAllocation = mInputs(index)
End Function

Public Property Get OutputSystemKey() As String
    OutputSystemKey = mOutputSystemKey
End Property

Public Property Get OutputSku() As String
    OutputSku = mOutputSku
End Property

Public Property Get OutputQty() As Double
    OutputQty = mOutputQty
End Property

Public Property Get OutputCondition() As String
    OutputCondition = mOutputCondition
End Property

Public Property Get OutputAttributesJson() As String
    OutputAttributesJson = mOutputAttributesJson
End Property

Public Property Get ConsumeEventId() As String
    ConsumeEventId = mConsumeEventId
End Property

Public Property Get CompleteEventId() As String
    CompleteEventId = mCompleteEventId
End Property

Public Property Get ConsumeQueued() As Boolean
    ConsumeQueued = mConsumeQueued
End Property

Public Property Get CompleteQueued() As Boolean
    CompleteQueued = mCompleteQueued
End Property

Public Property Get ConsumeApplied() As Boolean
    ConsumeApplied = mConsumeApplied
End Property

Public Property Get CompleteApplied() As Boolean
    CompleteApplied = mCompleteApplied
End Property

Public Property Get ProcessorVerified() As Boolean
    ProcessorVerified = mProcessorVerified
End Property

Public Property Get RefreshVerified() As Boolean
    RefreshVerified = mRefreshVerified
End Property

Public Property Get ReadyForNextBatch() As Boolean
    ReadyForNextBatch = (mStatus = "READY")
End Property

Public Property Get Status() As String
    Status = mStatus
End Property

Public Property Get FailureCode() As String
    FailureCode = mFailureCode
End Property

Public Property Get FailureMessage() As String
    FailureMessage = mFailureMessage
End Property

Public Property Get CompensationRequired() As Boolean
    CompensationRequired = mCompensationRequired
End Property

Private Sub RequireInitialized(ByVal memberName As String)
    If Not mInitialized Then _
        Err.Raise vbObjectError + 7621, "cProductionRunSession." & memberName, _
                  "Production session has not been initialized."
End Sub

Private Sub RefreshReadyState()
    If mFailureCode <> "" Or mFailureMessage <> "" Or mCompensationRequired Then Exit Sub
    If mProcessorVerified And mRefreshVerified And mConsumeApplied And mCompleteApplied Then mStatus = "READY"
End Sub
