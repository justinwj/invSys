Attribute VB_Name = "modAdminDesignLifecycle"
Option Explicit

Private Const DESIGN_EVENT_RELEASE As String = "DESIGN_RELEASE"
Private Const DESIGN_EVENT_OBSOLETE As String = "DESIGN_OBSOLETE"

Public Sub Admin_ReleaseDesignVersion_Click()
    PromptAndQueueDesignLifecycleEvent DESIGN_EVENT_RELEASE
End Sub

Public Sub Admin_ObsoleteDesignVersion_Click()
    PromptAndQueueDesignLifecycleEvent DESIGN_EVENT_OBSOLETE
End Sub

Public Function ReleaseDesignVersion(ByVal designId As String, _
                                     ByVal designVersion As String, _
                                     Optional ByVal noteVal As String = "", _
                                     Optional ByRef eventIdOut As String = "", _
                                     Optional ByRef errorMessage As String = "") As Boolean
    ReleaseDesignVersion = QueueCurrentDesignLifecycleEvent( _
        DESIGN_EVENT_RELEASE, designId, designVersion, noteVal, eventIdOut, errorMessage)
End Function

Public Function ObsoleteDesignVersion(ByVal designId As String, _
                                      ByVal designVersion As String, _
                                      Optional ByVal noteVal As String = "", _
                                      Optional ByRef eventIdOut As String = "", _
                                      Optional ByRef errorMessage As String = "") As Boolean
    ObsoleteDesignVersion = QueueCurrentDesignLifecycleEvent( _
        DESIGN_EVENT_OBSOLETE, designId, designVersion, noteVal, eventIdOut, errorMessage)
End Function

Private Sub PromptAndQueueDesignLifecycleEvent(ByVal eventType As String)
    On Error GoTo FailPrompt

    Dim designId As String
    Dim designVersion As String
    Dim noteVal As String
    Dim eventId As String
    Dim queueError As String
    Dim processorReport As String
    Dim appliedCount As Long
    Dim actionLabel As String
    Dim expectedStatus As String
    Dim currentStatus As String

    If eventType = DESIGN_EVENT_RELEASE Then
        actionLabel = "Release Design"
        expectedStatus = "RELEASED"
    Else
        actionLabel = "Obsolete Design"
        expectedStatus = "OBSOLETE"
    End If

    designId = Trim$(InputBox("Enter the DesignId.", "invSys Admin - " & actionLabel))
    If designId = "" Then Exit Sub
    designVersion = Trim$(InputBox("Enter the immutable DesignVersion for " & designId & ".", _
                                   "invSys Admin - " & actionLabel))
    If designVersion = "" Then Exit Sub
    noteVal = Trim$(InputBox("Enter an audit note (optional).", "invSys Admin - " & actionLabel))

    If eventType = DESIGN_EVENT_OBSOLETE Then
        If MsgBox("Obsolete design " & designId & " version " & designVersion & "?" & vbCrLf & _
                  "Production will no longer offer this version for new runs.", _
                  vbQuestion Or vbYesNo Or vbDefaultButton2, _
                  "invSys Admin - Obsolete Design") <> vbYes Then Exit Sub
    End If

    If eventType = DESIGN_EVENT_RELEASE Then
        If Not ReleaseDesignVersion(designId, designVersion, noteVal, eventId, queueError) Then
            MsgBox actionLabel & " was not queued." & vbCrLf & vbCrLf & queueError, _
                   vbExclamation, "invSys Admin"
            Exit Sub
        End If
    Else
        If Not ObsoleteDesignVersion(designId, designVersion, noteVal, eventId, queueError) Then
            MsgBox actionLabel & " was not queued." & vbCrLf & vbCrLf & queueError, _
                   vbExclamation, "invSys Admin"
            Exit Sub
        End If
    End If

    appliedCount = modProcessor.RunBatch("", 0, processorReport)
    currentStatus = FindCurrentDesignStatus(designId, designVersion)
    If StrComp(currentStatus, expectedStatus, vbTextCompare) = 0 Then
        MsgBox actionLabel & " completed." & vbCrLf & _
               "Design: " & designId & " version " & designVersion & vbCrLf & _
               "Status: " & currentStatus & vbCrLf & _
               "EventID: " & eventId, vbInformation, "invSys Admin"
    Else
        MsgBox actionLabel & " was queued for processor handling." & vbCrLf & _
               "Design: " & designId & " version " & designVersion & vbCrLf & _
               "EventID: " & eventId & vbCrLf & _
               "Processor applied: " & CStr(appliedCount) & vbCrLf & vbCrLf & processorReport, _
               vbInformation, "invSys Admin"
    End If
    Exit Sub

FailPrompt:
    MsgBox actionLabel & " failed: " & Err.Description, vbExclamation, "invSys Admin"
End Sub

Private Function FindCurrentDesignStatus(ByVal designId As String, _
                                         ByVal designVersion As String) As String
    On Error GoTo CleanExit

    Dim designs As Variant
    Dim r As Long

    designs = modDesignsDomainBridge.ListDesignsBridge(Nothing, "")
    If IsEmpty(designs) Or Not IsArray(designs) Then Exit Function
    For r = LBound(designs, 1) To UBound(designs, 1)
        If StrComp(Trim$(CStr(designs(r, 1))), Trim$(designId), vbTextCompare) = 0 _
           And StrComp(Trim$(CStr(designs(r, 2))), Trim$(designVersion), vbTextCompare) = 0 Then
            FindCurrentDesignStatus = Trim$(CStr(designs(r, 6)))
            Exit Function
        End If
    Next r
CleanExit:
End Function

Public Function QueueAdminDesignLifecycleEvent(ByVal eventType As String, _
                                               ByVal warehouseId As String, _
                                               ByVal stationId As String, _
                                               ByVal userId As String, _
                                               ByVal designId As String, _
                                               ByVal designVersion As String, _
                                               Optional ByVal noteVal As String = "", _
                                               Optional ByVal targetInboxWb As Workbook = Nothing, _
                                               Optional ByRef eventIdOut As String = "", _
                                               Optional ByRef errorMessage As String = "") As Boolean
    eventType = NormalizeDesignLifecycleEventType(eventType, errorMessage)
    If eventType = "" Then Exit Function
    If Not ValidateDesignIdentity(designId, designVersion, errorMessage) Then Exit Function

    QueueAdminDesignLifecycleEvent = modRoleEventWriter.QueueDesignEvent( _
        eventType, warehouseId, stationId, userId, designId, designVersion, _
        "", "", noteVal, 0, targetInboxWb, eventIdOut, errorMessage)
End Function

Private Function QueueCurrentDesignLifecycleEvent(ByVal eventType As String, _
                                                  ByVal designId As String, _
                                                  ByVal designVersion As String, _
                                                  ByVal noteVal As String, _
                                                  ByRef eventIdOut As String, _
                                                  ByRef errorMessage As String) As Boolean
    eventType = NormalizeDesignLifecycleEventType(eventType, errorMessage)
    If eventType = "" Then Exit Function
    If Not ValidateDesignIdentity(designId, designVersion, errorMessage) Then Exit Function

    QueueCurrentDesignLifecycleEvent = modRoleEventWriter.QueueDesignEventCurrent( _
        eventType, designId, designVersion, "", noteVal, "", eventIdOut, errorMessage)
End Function

Private Function NormalizeDesignLifecycleEventType(ByVal eventType As String, _
                                                   ByRef errorMessage As String) As String
    eventType = UCase$(Trim$(eventType))
    Select Case eventType
        Case DESIGN_EVENT_RELEASE, DESIGN_EVENT_OBSOLETE
            NormalizeDesignLifecycleEventType = eventType
        Case Else
            errorMessage = "Admin Designs lifecycle supports only DESIGN_RELEASE or DESIGN_OBSOLETE."
    End Select
End Function

Private Function ValidateDesignIdentity(ByRef designId As String, _
                                        ByRef designVersion As String, _
                                        ByRef errorMessage As String) As Boolean
    designId = Trim$(designId)
    designVersion = Trim$(designVersion)
    If designId = "" Or designVersion = "" Then
        errorMessage = "DesignId and DesignVersion are required."
        Exit Function
    End If
    ValidateDesignIdentity = True
End Function
