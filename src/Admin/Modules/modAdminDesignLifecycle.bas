Attribute VB_Name = "modAdminDesignLifecycle"
Option Explicit

Private Const DESIGN_EVENT_RELEASE As String = "DESIGN_RELEASE"
Private Const DESIGN_EVENT_OBSOLETE As String = "DESIGN_OBSOLETE"

Public Sub Admin_ReleaseDesignVersion_Click()
    frmAdminDesignLifecycle.Show vbModal
End Sub

Public Sub Admin_ObsoleteDesignVersion_Click()
    frmAdminDesignLifecycle.Show vbModal
End Sub

Public Function DesignLifecycleFormLayoutSmokeForAutomation() As Long
    On Error GoTo CleanExit
    DesignLifecycleFormLayoutSmokeForAutomation = frmAdminDesignLifecycle.TestLayoutReady()
CleanExit:
    On Error Resume Next
    Unload frmAdminDesignLifecycle
    On Error GoTo 0
End Function

Public Function MigrateLegacyRecipesFromWorkbook(ByVal donorWb As Workbook, _
                                                 Optional ByRef report As String = "") As Boolean
    Dim migrationReport As String
    Dim processorReport As String
    Dim processedCount As Long

    If donorWb Is Nothing Then
        report = "No open legacy recipe workbook was found."
        Exit Function
    End If
    If donorWb.IsAddin Then
        report = "An operator/data workbook is required; an XLAM cannot be a migration donor."
        Exit Function
    End If
    If Not modAdminDesignMigration.QueueLegacyRecipeDesignMigration(donorWb, "Recipes", migrationReport) Then
        report = migrationReport
        Exit Function
    End If

    processedCount = modProcessor.RunBatch("", 0, processorReport)
    report = migrationReport & " Processor applied " & CStr(processedCount) & "."
    If Trim$(processorReport) <> "" Then report = report & " " & processorReport
    MigrateLegacyRecipesFromWorkbook = True
End Function

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

Public Function ExecuteDesignLifecycleCommand(ByVal eventType As String, _
                                              ByVal designId As String, _
                                              ByVal designVersion As String, _
                                              Optional ByVal noteVal As String = "", _
                                              Optional ByRef report As String = "") As Boolean
    Dim eventId As String
    Dim queueError As String
    Dim processorReport As String
    Dim appliedCount As Long
    Dim expectedStatus As String
    Dim currentStatus As String

    eventType = NormalizeDesignLifecycleEventType(eventType, report)
    If eventType = "" Then Exit Function
    If eventType = DESIGN_EVENT_RELEASE Then
        expectedStatus = "RELEASED"
    Else
        expectedStatus = "OBSOLETE"
    End If

    If Not QueueCurrentDesignLifecycleEvent(eventType, designId, designVersion, _
                                            noteVal, eventId, queueError) Then
        report = queueError
        Exit Function
    End If
    appliedCount = modProcessor.RunBatch("", 0, processorReport)
    currentStatus = FindCurrentDesignStatus(designId, designVersion)
    If StrComp(currentStatus, expectedStatus, vbTextCompare) = 0 Then
        report = expectedStatus & "; EventID=" & eventId
        ExecuteDesignLifecycleCommand = True
    Else
        report = "Queued EventID=" & eventId & "; processor applied " & _
                 CStr(appliedCount) & ". " & processorReport
    End If
End Function

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
