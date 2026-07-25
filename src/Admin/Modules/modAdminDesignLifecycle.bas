Attribute VB_Name = "modAdminDesignLifecycle"
Option Explicit

Private Const DESIGN_EVENT_RELEASE As String = "DESIGN_RELEASE"
Private Const DESIGN_EVENT_OBSOLETE As String = "DESIGN_OBSOLETE"

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
