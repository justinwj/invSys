Attribute VB_Name = "TestCoreRoleEventWriter"
Option Explicit

Private mLastTestFailure As String

Public Sub ClearLastTestFailure()
    mLastTestFailure = ""
End Sub

Public Function GetLastTestFailure() As String
    GetLastTestFailure = mLastTestFailure
End Function

Public Sub RunCoreRoleEventWriterTests()
    Dim passed As Long
    Dim failed As Long

    Tally TestQueueReceiveEvent_WritesInboxRow(), passed, failed
    Tally TestOpenInboxWorkbook_UsesStationPathInboxRoot(), passed, failed
    Tally TestQueueShipEvent_WritesInboxRow(), passed, failed
    Tally TestQueuePayloadEvent_DeniedWithoutCapability(), passed, failed
    Tally TestBuildPayloadJson_WithObjectItems(), passed, failed

    Debug.Print "Core.RoleEventWriter tests - Passed: " & passed & " Failed: " & failed
End Sub

Public Function TestQueueReceiveEvent_WritesInboxRow() As Long
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim wbInbox As Workbook
    Dim lo As ListObject
    Dim eventIdOut As String
    Dim errorMessage As String
    Dim rootPath As String

    rootPath = TestPhase2Helpers.BuildUniqueTestFolder("role_writer_receive")
    Set wbCfg = TestPhase2Helpers.BuildCanonicalConfigWorkbook("WHR1", "R1", rootPath, "RECEIVE")
    Set wbAuth = TestPhase2Helpers.BuildCanonicalAuthWorkbook("WHR1", rootPath)
    Set wbInbox = TestPhase2Helpers.BuildCanonicalReceiveInboxWorkbook("R1", rootPath)
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    TestPhase2Helpers.AddCapability wbAuth, "user1", "RECEIVE_POST", "WHR1", "R1", "ACTIVE"
    wbAuth.Save

    On Error GoTo CleanFail
    If Not modRoleEventWriter.QueueReceiveEvent("WHR1", "R1", "user1", "SKU-001", 4, "A1", "receive test", "", "", Now, wbInbox, eventIdOut, errorMessage) Then
        mLastTestFailure = errorMessage
        GoTo CleanExit
    End If

    Set lo = wbInbox.Worksheets("InboxReceive").ListObjects("tblInboxReceive")
    If lo.ListRows.Count <> 1 Then GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(lo, 1, "EventID")) <> eventIdOut Then GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(lo, 1, "EventType")) <> EVENT_TYPE_RECEIVE Then GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(lo, 1, "UserId")) <> "user1" Then GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(lo, 1, "System_Key")) = "" Then GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(lo, 1, "SKU")) <> "SKU-001" Then GoTo CleanExit
    If CDbl(TestPhase2Helpers.GetRowValue(lo, 1, "Qty")) <> 4 Then GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(lo, 1, "Status")) <> "NEW" Then GoTo CleanExit

    TestQueueReceiveEvent_WritesInboxRow = 1

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    TestPhase2Helpers.CloseNoSave wbInbox
    TestPhase2Helpers.CloseNoSave wbAuth
    TestPhase2Helpers.CloseNoSave wbCfg
    On Error Resume Next
    If rootPath <> "" Then CreateObject("Scripting.FileSystemObject").DeleteFolder rootPath, True
    On Error GoTo 0
    Exit Function
CleanFail:
    mLastTestFailure = Err.Description
    Resume CleanExit
End Function

Public Function TestOpenInboxWorkbook_UsesStationPathInboxRoot() As Long
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim wbInbox As Workbook
    Dim inboxRoot As String
    Dim expectedPath As String
    Dim errorMessage As String

    inboxRoot = Environ$("TEMP") & "\invsys_role_writer_" & Format$(Now, "yyyymmdd_hhnnss")
    If Len(Dir$(inboxRoot, vbDirectory)) = 0 Then MkDir inboxRoot

    Set wbCfg = TestPhase2Helpers.BuildPhase2ConfigWorkbook("WHR2", "R2", "RECEIVE")
    Set wbAuth = TestPhase2Helpers.BuildPhase2AuthWorkbook("WHR2")
    TestPhase2Helpers.SetWarehouseConfigValue wbCfg, "PathDataRoot", Environ$("TEMP") & "\invsys_wrong_data_root"
    TestPhase2Helpers.SetStationConfigValue wbCfg, "PathInboxRoot", inboxRoot
    TestPhase2Helpers.AddCapability wbAuth, "user1", "RECEIVE_POST", "WHR2", "R2", "ACTIVE"

    On Error GoTo CleanFail
    Set wbInbox = modRoleEventWriter.OpenInboxWorkbook(EVENT_TYPE_RECEIVE, "WHR2", "R2", errorMessage)
    If wbInbox Is Nothing Then GoTo CleanExit

    expectedPath = inboxRoot & "\invSys.Inbox.Receiving.R2.xlsb"
    If StrComp(wbInbox.FullName, expectedPath, vbTextCompare) <> 0 Then GoTo CleanExit
    If Len(Dir$(expectedPath, vbNormal)) = 0 Then GoTo CleanExit

    TestOpenInboxWorkbook_UsesStationPathInboxRoot = 1

CleanExit:
    TestPhase2Helpers.CloseNoSave wbInbox
    On Error Resume Next
    If expectedPath <> "" Then
        If Len(Dir$(expectedPath, vbNormal)) > 0 Then Kill expectedPath
    End If
    If Len(Dir$(inboxRoot, vbDirectory)) > 0 Then RmDir inboxRoot
    On Error GoTo 0
    TestPhase2Helpers.CloseNoSave wbAuth
    TestPhase2Helpers.CloseNoSave wbCfg
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestQueueShipEvent_WritesInboxRow() As Long
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim wbInbox As Workbook
    Dim lo As ListObject
    Dim payloadJson As String
    Dim eventIdOut As String
    Dim errorMessage As String

    Set wbCfg = TestPhase2Helpers.BuildPhase2ConfigWorkbook("WHS1", "H1", "SHIP")
    Set wbAuth = TestPhase2Helpers.BuildPhase2AuthWorkbook("WHS1")
    Set wbInbox = TestPhase2Helpers.BuildShipInboxWorkbook("H1")
    TestPhase2Helpers.AddCapability wbAuth, "user1", "SHIP_POST", "WHS1", "H1", "ACTIVE"

    Dim payloadItems As Collection
    Set payloadItems = New Collection
    payloadItems.Add modRoleEventWriter.CreatePayloadItem(101, "SKU-001", 2, "DOCK", "line 1")
    payloadItems.Add modRoleEventWriter.CreatePayloadItem(102, "SKU-002", 3, "DOCK", "line 2")
    payloadJson = modRoleEventWriter.BuildPayloadJsonFromCollection(payloadItems)

    On Error GoTo CleanFail
    If Not modRoleEventWriter.QueuePayloadEvent(EVENT_TYPE_SHIP, "WHS1", "H1", "user1", payloadJson, "ship test", "", "", Now, wbInbox, eventIdOut, errorMessage) Then GoTo CleanExit

    Set lo = wbInbox.Worksheets("InboxShip").ListObjects("tblInboxShip")
    If lo.ListRows.Count <> 2 Then GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(lo, 2, "EventID")) <> eventIdOut Then GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(lo, 2, "EventType")) <> EVENT_TYPE_SHIP Then GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(lo, 2, "PayloadJson")) <> payloadJson Then GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(lo, 2, "Status")) <> "NEW" Then GoTo CleanExit

    TestQueueShipEvent_WritesInboxRow = 1

CleanExit:
    TestPhase2Helpers.CloseNoSave wbInbox
    TestPhase2Helpers.CloseNoSave wbAuth
    TestPhase2Helpers.CloseNoSave wbCfg
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestQueuePayloadEvent_DeniedWithoutCapability() As Long
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim wbInbox As Workbook
    Dim lo As ListObject
    Dim payloadJson As String
    Dim eventIdOut As String
    Dim errorMessage As String

    Set wbCfg = TestPhase2Helpers.BuildPhase2ConfigWorkbook("WHP1", "P1", "PROD")
    Set wbAuth = TestPhase2Helpers.BuildPhase2AuthWorkbook("WHP1")
    Set wbInbox = TestPhase2Helpers.BuildProductionInboxWorkbook("P1")

    Dim payloadItems As Collection
    Set payloadItems = New Collection
    payloadItems.Add modRoleEventWriter.CreatePayloadItem(201, "SKU-001", 1, "LINE1", "made line", "MADE")
    payloadJson = modRoleEventWriter.BuildPayloadJsonFromCollection(payloadItems)

    On Error GoTo CleanFail
    If modRoleEventWriter.QueuePayloadEvent(EVENT_TYPE_PROD_COMPLETE, "WHP1", "P1", "user1", payloadJson, "prod test", "", "", Now, wbInbox, eventIdOut, errorMessage) Then GoTo CleanExit
    If InStr(1, errorMessage, "PROD_POST", vbTextCompare) = 0 Then GoTo CleanExit

    Set lo = wbInbox.Worksheets("InboxProd").ListObjects("tblInboxProd")
    If lo.ListRows.Count <> 0 Then GoTo CleanExit

    TestQueuePayloadEvent_DeniedWithoutCapability = 1

CleanExit:
    TestPhase2Helpers.CloseNoSave wbInbox
    TestPhase2Helpers.CloseNoSave wbAuth
    TestPhase2Helpers.CloseNoSave wbCfg
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestBuildPayloadJson_WithObjectItems() As Long
    Dim item1 As Object
    Dim item2 As Object
    Dim payloadJson As String

    On Error GoTo CleanFail
    Set item1 = modRoleEventWriter.CreatePayloadItem(101, "SKU-001", 2, "DOCK", "line 1")
    Set item2 = modRoleEventWriter.CreatePayloadItem(102, "SKU-002", 3, "DOCK", "line 2")
    payloadJson = modRoleEventWriter.BuildPayloadJson(item1, item2)

    If InStr(1, payloadJson, """SKU"":""SKU-001""", vbTextCompare) = 0 Then GoTo CleanExit
    If InStr(1, payloadJson, """SKU"":""SKU-002""", vbTextCompare) = 0 Then GoTo CleanExit

    TestBuildPayloadJson_WithObjectItems = 1

CleanExit:
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestQueueDesignMigrationEvent_PreservesDeterministicIdentity() As Long
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim wbInbox As Workbook
    Dim lo As ListObject
    Dim payloadJson As String
    Dim eventIdOut As String
    Dim errorMessage As String
    Dim migrationSourceId As String
    Dim rowIndex As Long
    Dim foundRow As Long
    Dim failureReason As String
    Dim rootPath As String

    rootPath = TestPhase2Helpers.BuildUniqueTestFolder("design_migration_inbox")
    Set wbCfg = TestPhase2Helpers.BuildCanonicalConfigWorkbook("WHD1", "D1", rootPath, "PROD")
    Set wbAuth = TestPhase2Helpers.BuildCanonicalAuthWorkbook("WHD1", rootPath)
    Set wbInbox = TestPhase2Helpers.BuildCanonicalProductionInboxWorkbook("D1", rootPath)
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    TestPhase2Helpers.AddCapability wbAuth, "user1", "PROD_POST", "WHD1", "D1", "ACTIVE"
    wbAuth.Save
    payloadJson = "[{""DesignType"":""RECIPE"",""DesignName"":""Migrated Tea""}]"
    eventIdOut = "MIG-DESIGN-TEA-1-ABC123"
    migrationSourceId = "LEGACY_RECIPE|DONOR.XLSB|RECIPES|TEA-1|1"

    On Error GoTo CleanFail
    If Not modRoleEventWriter.QueueDesignEvent( _
        "DESIGN_CREATE", "WHD1", "D1", "user1", "TEA-1", "1", _
        payloadJson, migrationSourceId, "migration test", Now, wbInbox, _
        eventIdOut, errorMessage) Then
        failureReason = "QueueDesignEvent failed: " & errorMessage
        GoTo CleanExit
    End If

    Set lo = wbInbox.Worksheets("InboxProd").ListObjects("tblInboxProd")
    For rowIndex = 1 To lo.ListRows.Count
        If CStr(TestPhase2Helpers.GetRowValue(lo, rowIndex, "EventID")) = "MIG-DESIGN-TEA-1-ABC123" Then
            foundRow = rowIndex
            Exit For
        End If
    Next rowIndex
    If foundRow = 0 Then
        failureReason = "Deterministic EventID was not found in tblInboxProd."
        GoTo CleanExit
    End If
    If CStr(TestPhase2Helpers.GetRowValue(lo, foundRow, "EventType")) <> "DESIGN_CREATE" Then failureReason = "EventType mismatch.": GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(lo, foundRow, "DesignId")) <> "TEA-1" Then failureReason = "DesignId mismatch.": GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(lo, foundRow, "DesignVersion")) <> "1" Then failureReason = "DesignVersion mismatch.": GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(lo, foundRow, "MigrationSourceId")) <> migrationSourceId Then failureReason = "MigrationSourceId mismatch.": GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(lo, foundRow, "PayloadJson")) <> payloadJson Then failureReason = "PayloadJson mismatch.": GoTo CleanExit
    TestQueueDesignMigrationEvent_PreservesDeterministicIdentity = 1

CleanExit:
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    TestPhase2Helpers.CloseNoSave wbInbox
    TestPhase2Helpers.CloseNoSave wbAuth
    TestPhase2Helpers.CloseNoSave wbCfg
    On Error Resume Next
    If rootPath <> "" Then CreateObject("Scripting.FileSystemObject").DeleteFolder rootPath, True
    On Error GoTo 0
    mLastTestFailure = failureReason
    Exit Function
CleanFail:
    If failureReason = "" Then failureReason = Err.Description
    Resume CleanExit
End Function

Public Function TestAdminDesignLifecycle_QueuesAuthorizedRelease() As Long
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim wbInbox As Workbook
    Dim lo As ListObject
    Dim eventIdOut As String
    Dim errorMessage As String
    Dim rowIndex As Long
    Dim foundRow As Long
    Dim failureReason As String
    Dim rootPath As String

    rootPath = TestPhase2Helpers.BuildUniqueTestFolder("design_lifecycle_inbox")
    Set wbCfg = TestPhase2Helpers.BuildCanonicalConfigWorkbook("WHD2", "D2", rootPath, "ADMIN")
    Set wbAuth = TestPhase2Helpers.BuildCanonicalAuthWorkbook("WHD2", rootPath)
    Set wbInbox = TestPhase2Helpers.BuildCanonicalProductionInboxWorkbook("D2", rootPath)
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    TestPhase2Helpers.AddCapability wbAuth, "admin1", "ADMIN_MAINT", "WHD2", "D2", "ACTIVE"
    wbAuth.Save

    On Error GoTo CleanFail
    If Not modAdminDesignLifecycle.QueueAdminDesignLifecycleEvent( _
        "DESIGN_RELEASE", "WHD2", "D2", "admin1", "TEA-ADMIN", "7", _
        "approved release", wbInbox, eventIdOut, errorMessage) Then
        failureReason = "QueueAdminDesignLifecycleEvent failed: " & errorMessage
        GoTo CleanExit
    End If

    Set lo = wbInbox.Worksheets("InboxProd").ListObjects("tblInboxProd")
    For rowIndex = 1 To lo.ListRows.Count
        If CStr(TestPhase2Helpers.GetRowValue(lo, rowIndex, "EventID")) = eventIdOut Then
            foundRow = rowIndex
            Exit For
        End If
    Next rowIndex
    If foundRow = 0 Then failureReason = "Queued release EventID was not found.": GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(lo, foundRow, "EventType")) <> "DESIGN_RELEASE" Then failureReason = "EventType mismatch.": GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(lo, foundRow, "DesignId")) <> "TEA-ADMIN" Then failureReason = "DesignId mismatch.": GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(lo, foundRow, "DesignVersion")) <> "7" Then failureReason = "DesignVersion mismatch.": GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(lo, foundRow, "PayloadJson")) <> "" Then failureReason = "Lifecycle payload must be empty.": GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(lo, foundRow, "Note")) <> "approved release" Then failureReason = "Lifecycle note mismatch.": GoTo CleanExit
    TestAdminDesignLifecycle_QueuesAuthorizedRelease = 1

CleanExit:
    mLastTestFailure = failureReason
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    TestPhase2Helpers.CloseNoSave wbInbox
    TestPhase2Helpers.CloseNoSave wbAuth
    TestPhase2Helpers.CloseNoSave wbCfg
    On Error Resume Next
    If rootPath <> "" Then CreateObject("Scripting.FileSystemObject").DeleteFolder rootPath, True
    On Error GoTo 0
    Exit Function

CleanFail:
    failureReason = "Unexpected error " & CStr(Err.Number) & ": " & Err.Description
    Resume CleanExit
End Function

Private Sub Tally(ByVal resultIn As Long, ByRef passed As Long, ByRef failed As Long)
    If resultIn = 1 Then
        passed = passed + 1
    Else
        failed = failed + 1
    End If
End Sub
