Attribute VB_Name = "test_ConfirmWrites_Tester"
Option Explicit

Private Type ConfirmWritesFixture
    RuntimeBase As String
    RuntimeRoot As String
    ShareRoot As String
    TemplateRoot As String
    WarehouseId As String
    StationId As String
    TesterUserId As String
    RuntimeUserId As String
    OperatorPath As String
    SharePointRoot As String
End Type

Private mCaseNames() As String
Private mCaseResults() As String
Private mCaseDetails() As String
Private mCaseCount As Long
Private mSummary As String
Private mRuntimeUserId As String
Private mTesterUserId As String
Private mLastReceiveEventId As String
Private mLastReceiveRefNumber As String

Public Function TestConfirmWrites_Tester_EndToEnd() As Long
    Dim fx As ConfirmWritesFixture
    Dim offlineFx As ConfirmWritesFixture
    Dim detailText As String

    On Error GoTo FailTest

    ResetConfirmWritesEvidence

    If SetupConfirmWritesFixture(fx, detailText, vbNullString, "primary") Then
        RecordConfirmWritesCase "ReadinessCheck_OK", RunReadinessCheckOkCase(fx, detailText), detailText
        RecordConfirmWritesCase "ReadinessCheck_MissingCapability", RunReadinessCheckMissingCapabilityCase(fx, detailText), detailText
        RecordConfirmWritesCase "RefreshInventory_Loads", RunRefreshInventoryLoadsCase(fx, detailText), detailText
        RecordConfirmWritesCase "ConfirmWrites_SingleRow", RunConfirmWritesSingleRowCase(fx, detailText), detailText
        RecordConfirmWritesCase "ProcessorApplies", RunProcessorAppliesCase(fx, detailText), detailText
        RecordConfirmWritesCase "SnapshotRefreshAfterPost", RunSnapshotRefreshAfterPostCase(fx, detailText), detailText
        RecordConfirmWritesCase "IdempotentSetup", RunIdempotentSetupCase(fx, detailText), detailText
    Else
        RecordFixtureFailureCases detailText
    End If

    CleanupConfirmWritesFixture fx

    If SetupConfirmWritesFixture(offlineFx, detailText, "C:\Invalid<SharePointRoot", "offline") Then
        RecordConfirmWritesCase "SharePointUnavailable", RunSharePointUnavailableCase(offlineFx, detailText), detailText
    Else
        RecordConfirmWritesCase "SharePointUnavailable", False, detailText
    End If
    CleanupConfirmWritesFixture offlineFx

    If AllConfirmWritesCasesPassed() Then
        mSummary = "Confirm Writes tester proving flow passed all eight deterministic cases."
        TestConfirmWrites_Tester_EndToEnd = 1
    Else
        mSummary = "One or more Confirm Writes tester proving cases failed."
    End If
    Exit Function

FailTest:
    RecordConfirmWritesCase "Harness.Exception", False, Err.Description
    CleanupConfirmWritesFixture fx
    CleanupConfirmWritesFixture offlineFx
    mSummary = "Confirm Writes tester integration raised an unexpected exception."
End Function

Public Function GetConfirmWritesTesterContextPacked() As String
    GetConfirmWritesTesterContextPacked = _
        "Summary=" & SafeConfirmWritesText(mSummary) & _
        "|RuntimeUser=" & SafeConfirmWritesText(mRuntimeUserId) & _
        "|TesterUser=" & SafeConfirmWritesText(mTesterUserId)
End Function

Public Function GetConfirmWritesTesterEvidenceRows() As String
    Dim i As Long

    For i = 1 To mCaseCount
        If Len(GetConfirmWritesTesterEvidenceRows) > 0 Then GetConfirmWritesTesterEvidenceRows = GetConfirmWritesTesterEvidenceRows & vbLf
        GetConfirmWritesTesterEvidenceRows = GetConfirmWritesTesterEvidenceRows & _
            mCaseNames(i) & vbTab & mCaseResults(i) & vbTab & mCaseDetails(i)
    Next i
End Function

Private Function SetupConfirmWritesFixture(ByRef fx As ConfirmWritesFixture, _
                                           ByRef detailText As String, _
                                           Optional ByVal sharePointRootOverride As String = "", _
                                           Optional ByVal suffix As String = "primary") As Boolean
    Dim spec As modTesterSetup.TesterSetupSpec
    Dim report As String
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim wbInv As Workbook
    Dim wbOutbox As Workbook

    fx.WarehouseId = "WH1"
    fx.StationId = "R1"
    fx.TesterUserId = "TESTER01"
    fx.RuntimeUserId = ResolveCurrentRuntimeUserIdConfirmWrites()
    fx.RuntimeBase = BuildConfirmWritesTempRoot(suffix)
    fx.RuntimeRoot = fx.RuntimeBase & "\runtime\" & fx.WarehouseId
    fx.ShareRoot = fx.RuntimeBase & "\sharepoint"
    fx.TemplateRoot = fx.RuntimeBase & "\templates"
    If Trim$(sharePointRootOverride) <> "" Then
        fx.SharePointRoot = Trim$(sharePointRootOverride)
    Else
        fx.SharePointRoot = fx.ShareRoot
        EnsureFolderRecursiveConfirmWrites fx.ShareRoot
    End If
    fx.OperatorPath = fx.RuntimeRoot & "\" & fx.WarehouseId & ".Receiving.Operator.xlsm"

    mRuntimeUserId = fx.RuntimeUserId
    mTesterUserId = fx.TesterUserId

    spec = BuildTesterSpecConfirmWrites(fx.TesterUserId, "123456", fx.WarehouseId, fx.StationId, fx.RuntimeRoot, fx.SharePointRoot)

    On Error GoTo CleanFail
    EnsureFolderRecursiveConfirmWrites fx.RuntimeRoot
    EnsureFolderRecursiveConfirmWrites fx.RuntimeRoot & "\inbox"
    EnsureFolderRecursiveConfirmWrites fx.RuntimeRoot & "\outbox"
    EnsureFolderRecursiveConfirmWrites fx.RuntimeRoot & "\snapshots"
    EnsureFolderRecursiveConfirmWrites fx.RuntimeRoot & "\config"
    modRuntimeWorkbooks.SetCoreDataRootOverride fx.RuntimeRoot
    Set wbCfg = modRuntimeWorkbooks.OpenOrCreateConfigWorkbookRuntime(fx.WarehouseId, fx.StationId, fx.RuntimeRoot, report)
    If wbCfg Is Nothing Then
        detailText = report
        GoTo CleanExit
    End If
    Set wbAuth = modRuntimeWorkbooks.OpenOrCreateAuthWorkbookRuntime(fx.WarehouseId, "svc_processor", fx.RuntimeRoot, report)
    If wbAuth Is Nothing Then
        detailText = report
        GoTo CleanExit
    End If
    Set wbInv = modInventoryApply.ResolveInventoryWorkbook(fx.WarehouseId)
    If wbInv Is Nothing Then
        detailText = "Canonical inventory workbook could not be resolved for fixture setup."
        GoTo CleanExit
    End If
    Set wbOutbox = OpenWorkbookConfirmWrites(fx.RuntimeRoot & "\" & fx.WarehouseId & ".Outbox.Events.xlsb", False)
    If wbOutbox Is Nothing Then
        Set wbOutbox = Application.Workbooks.Add(xlWBATWorksheet)
        wbOutbox.SaveAs Filename:=fx.RuntimeRoot & "\" & fx.WarehouseId & ".Outbox.Events.xlsb", FileFormat:=50
    End If
    If Not modWarehouseSync.EnsureOutboxSchema(wbOutbox, report) Then
        detailText = report
        GoTo CleanExit
    End If
    modWarehouseBootstrap.SetWarehouseBootstrapTemplateRootOverride fx.TemplateRoot
    If Not modTesterSetup.SetupTesterStation(spec) Then
        detailText = modTesterSetup.GetLastTesterSetupReport()
        GoTo CleanExit
    End If

    modRuntimeWorkbooks.SetCoreDataRootOverride fx.RuntimeRoot

    If Not FileExistsConfirmWrites(fx.OperatorPath) Then
        detailText = "SetupTesterStation did not create the receiving operator workbook."
        GoTo CleanExit
    End If
    If ReadSkuQtyConfirmWrites(fx.RuntimeRoot, fx.WarehouseId, "TEST-SKU-001") <> 100# Then
        detailText = "SetupTesterStation did not seed TEST-SKU-001 with QtyOnHand = 100."
        GoTo CleanExit
    End If
    If Not EnsureRuntimeUserCapabilitiesConfirmWrites(fx, detailText) Then GoTo CleanExit

    SetupConfirmWritesFixture = True
    detailText = "Fixture ready. Operator workbook exists, TEST-SKU-001 started at QtyOnHand = 100, and runtime auth is provisioned for user " & fx.RuntimeUserId & "."
    If StrComp(fx.RuntimeUserId, fx.TesterUserId, vbTextCompare) <> 0 Then
        detailText = detailText & " TESTER01 remains the requested setup user; runtime capabilities were mirrored because receiving resolves the executing machine user."
    End If

CleanExit:
    CloseWorkbookIfTransientConfirmWrites wbOutbox
    CloseWorkbookIfTransientConfirmWrites wbInv
    CloseWorkbookIfTransientConfirmWrites wbAuth
    CloseWorkbookIfTransientConfirmWrites wbCfg
    modWarehouseBootstrap.ClearWarehouseBootstrapTemplateRootOverride
    Exit Function
CleanFail:
    detailText = Err.Description
    Resume CleanExit
End Function

Private Sub CleanupConfirmWritesFixture(ByRef fx As ConfirmWritesFixture)
    CloseWorkbookByPathConfirmWrites fx.OperatorPath
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    If Trim$(fx.RuntimeBase) <> "" Then DeleteFolderRecursiveConfirmWrites fx.RuntimeBase
End Sub

Private Function RunReadinessCheckOkCase(ByRef fx As ConfirmWritesFixture, ByRef detailText As String) As Boolean
    Dim wbOps As Workbook
    Dim readiness As modReceivingInit.ReceivingReadinessResult

    On Error GoTo CleanFail
    Set wbOps = OpenReceivingWorkbookConfirmWrites(fx.OperatorPath, detailText)
    If wbOps Is Nothing Then GoTo CleanExit

    readiness = modReceivingInit.CheckReceivingReadinessForWorkbook(wbOps)
    If readiness.IsReady Then
        RunReadinessCheckOkCase = True
        detailText = "CheckReceivingReadiness returned ready for the configured WH1 receiving workbook."
    Else
        detailText = "IsReady=False; SnapshotStatus=" & readiness.SnapshotStatus & "; AuthStatus=" & readiness.AuthStatus & "; RuntimeStatus=" & readiness.RuntimeStatus & "; Messages=" & readiness.Messages
    End If

CleanExit:
    CloseWorkbookIfTransientConfirmWrites wbOps
    Exit Function
CleanFail:
    detailText = Err.Description
    Resume CleanExit
End Function

Private Function RunReadinessCheckMissingCapabilityCase(ByRef fx As ConfirmWritesFixture, ByRef detailText As String) As Boolean
    Dim wbOps As Workbook
    Dim readiness As modReceivingInit.ReceivingReadinessResult
    Dim restoreNeeded As Boolean

    On Error GoTo CleanFail
    restoreNeeded = True
    SetCapabilityStatusConfirmWrites fx.RuntimeRoot, fx.WarehouseId, fx.TesterUserId, "RECEIVE_POST", fx.StationId, "INACTIVE"
    If StrComp(fx.RuntimeUserId, fx.TesterUserId, vbTextCompare) <> 0 Then
        SetCapabilityStatusConfirmWrites fx.RuntimeRoot, fx.WarehouseId, fx.RuntimeUserId, "RECEIVE_POST", fx.StationId, "INACTIVE"
    End If

    Set wbOps = OpenReceivingWorkbookConfirmWrites(fx.OperatorPath, detailText)
    If wbOps Is Nothing Then GoTo CleanExit

    readiness = modReceivingInit.CheckReceivingReadinessForWorkbook(wbOps)
    If StrComp(readiness.AuthStatus, "MISSING_CAPABILITY", vbTextCompare) = 0 Then
        RunReadinessCheckMissingCapabilityCase = True
        detailText = "Readiness returned MISSING_CAPABILITY after RECEIVE_POST was removed from the effective runtime user."
    Else
        detailText = "Expected AuthStatus=MISSING_CAPABILITY but got " & readiness.AuthStatus & "."
    End If

CleanExit:
    If restoreNeeded Then
        SetCapabilityStatusConfirmWrites fx.RuntimeRoot, fx.WarehouseId, fx.TesterUserId, "RECEIVE_POST", fx.StationId, "ACTIVE"
        If StrComp(fx.RuntimeUserId, fx.TesterUserId, vbTextCompare) <> 0 Then
            SetCapabilityStatusConfirmWrites fx.RuntimeRoot, fx.WarehouseId, fx.RuntimeUserId, "RECEIVE_POST", fx.StationId, "ACTIVE"
        End If
    End If
    CloseWorkbookIfTransientConfirmWrites wbOps
    Exit Function
CleanFail:
    detailText = Err.Description
    Resume CleanExit
End Function

Private Function RunRefreshInventoryLoadsCase(ByRef fx As ConfirmWritesFixture, ByRef detailText As String) As Boolean
    Dim wbOps As Workbook
    Dim loInv As ListObject
    Dim rowIndex As Long
    Dim report As String

    On Error GoTo CleanFail
    Set wbOps = OpenReceivingWorkbookConfirmWrites(fx.OperatorPath, detailText)
    If wbOps Is Nothing Then GoTo CleanExit

    If Not modOperatorReadModel.RefreshInventoryReadModelForWorkbook(wbOps, fx.WarehouseId, "LOCAL", report) Then
        detailText = report
        GoTo CleanExit
    End If

    Set loInv = FindTableByNameConfirmWrites(wbOps, "invSys")
    rowIndex = FindRowByValueConfirmWrites(loInv, "ITEM_CODE", "TEST-SKU-001")
    If rowIndex = 0 Then
        detailText = "tblReadModel did not expose TEST-SKU-001 after refresh."
        GoTo CleanExit
    End If
    If CDbl(GetTableValueConfirmWrites(loInv, rowIndex, "QtyOnHand")) <> 100# Then
        detailText = "tblReadModel loaded TEST-SKU-001 but QtyOnHand was not 100."
        GoTo CleanExit
    End If

    RunRefreshInventoryLoadsCase = True
    detailText = "Refresh populated the read model with TEST-SKU-001 at QtyOnHand = 100."

CleanExit:
    CloseWorkbookIfTransientConfirmWrites wbOps
    Exit Function
CleanFail:
    detailText = Err.Description
    Resume CleanExit
End Function

Private Function RunConfirmWritesSingleRowCase(ByRef fx As ConfirmWritesFixture, ByRef detailText As String) As Boolean
    Const REF_NUMBER As String = "REF-CW-TEST-001"

    Dim wbOps As Workbook
    Dim wbInbox As Workbook
    Dim loInv As ListObject
    Dim loLog As ListObject
    Dim loRecv As ListObject
    Dim loAgg As ListObject
    Dim loInbox As ListObject
    Dim readModelRow As Long
    Dim inboxCountBefore As Long
    Dim logCountBefore As Long
    Dim inboxRow As Long
    Dim statusText As String
    Dim report As String

    On Error GoTo CleanFail
    Set wbOps = OpenReceivingWorkbookConfirmWrites(fx.OperatorPath, detailText)
    If wbOps Is Nothing Then GoTo CleanExit

    If Not modOperatorReadModel.RefreshInventoryReadModelForWorkbook(wbOps, fx.WarehouseId, "LOCAL", report) Then
        detailText = report
        GoTo CleanExit
    End If

    Set loInv = FindTableByNameConfirmWrites(wbOps, "invSys")
    Set loLog = FindTableByNameConfirmWrites(wbOps, "ReceivedLog")
    Set loRecv = FindTableByNameConfirmWrites(wbOps, "ReceivedTally")
    Set loAgg = FindTableByNameConfirmWrites(wbOps, "AggregateReceived")
    If loInv Is Nothing Or loLog Is Nothing Or loRecv Is Nothing Or loAgg Is Nothing Then
        detailText = "Receiving workbook did not expose the expected tables."
        GoTo CleanExit
    End If

    readModelRow = FindRowByValueConfirmWrites(loInv, "ITEM_CODE", "TEST-SKU-001")
    If readModelRow = 0 Then
        detailText = "TEST-SKU-001 was not available in the operator read model."
        GoTo CleanExit
    End If

    Set wbInbox = OpenReceiveInboxWorkbookConfirmWrites(fx, detailText)
    If wbInbox Is Nothing Then GoTo CleanExit
    Set loInbox = FindTableByNameConfirmWrites(wbInbox, "tblInboxReceive")
    If loInbox Is Nothing Then
        detailText = "tblInboxReceive was not available."
        GoTo CleanExit
    End If
    inboxCountBefore = loInbox.ListRows.Count
    logCountBefore = loLog.ListRows.Count

    wbOps.Activate
    modTS_Received.AddOrMergeFromSearch REF_NUMBER, _
                                        CStr(GetTableValueConfirmWrites(loInv, readModelRow, "ITEM")), _
                                        "TEST-SKU-001", _
                                        10, _
                                        vbNullString, _
                                        vbNullString, _
                                        CStr(GetTableValueConfirmWrites(loInv, readModelRow, "DESCRIPTION")), _
                                        ValueOrDefaultConfirmWrites(GetTableValueConfirmWrites(loInv, readModelRow, "UOM"), "EA"), _
                                        ValueOrDefaultConfirmWrites(GetTableValueConfirmWrites(loInv, readModelRow, "LOCATION"), "A1"), _
                                        CLng(GetTableValueConfirmWrites(loInv, readModelRow, "ROW"))
    modTS_Received.ConfirmWrites

    Set wbInbox = ReopenWorkbookConfirmWrites(wbInbox)
    If wbInbox Is Nothing Then
        detailText = "Receive inbox workbook could not be reopened after Confirm Writes."
        GoTo CleanExit
    End If
    Set loInbox = FindTableByNameConfirmWrites(wbInbox, "tblInboxReceive")
    inboxRow = FindLastInboxRowByNoteConfirmWrites(loInbox, "REF_NUMBER=" & REF_NUMBER)
    If inboxRow = 0 Then
        detailText = "Confirm Writes did not persist a receive event to tblInboxReceive."
        GoTo CleanExit
    End If

    mLastReceiveEventId = CStr(GetTableValueConfirmWrites(loInbox, inboxRow, "EventID"))
    mLastReceiveRefNumber = REF_NUMBER
    statusText = UCase$(CStr(GetTableValueConfirmWrites(loInbox, inboxRow, "Status")))

    If loRecv.ListRows.Count <> 0 Then
        detailText = "ReceivedTally was not cleared after Confirm Writes."
        GoTo CleanExit
    End If
    If loAgg.ListRows.Count <> 0 Then
        detailText = "AggregateReceived was not cleared after Confirm Writes."
        GoTo CleanExit
    End If
    If loLog.ListRows.Count <> logCountBefore + 1 Then
        detailText = "ReceivedLog did not append exactly one row."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetTableValueConfirmWrites(loLog, loLog.ListRows.Count, "REF_NUMBER")), REF_NUMBER, vbTextCompare) <> 0 Then
        detailText = "ReceivedLog did not append the expected REF_NUMBER."
        GoTo CleanExit
    End If
    If loInbox.ListRows.Count <> inboxCountBefore + 1 Then
        detailText = "tblInboxReceive row count did not increase by one."
        GoTo CleanExit
    End If

    RunConfirmWritesSingleRowCase = True
    detailText = "Confirm Writes cleared staging, appended ReceivedLog, and wrote inbox event " & mLastReceiveEventId & " with inbox status " & statusText & "."

CleanExit:
    CloseWorkbookIfTransientConfirmWrites wbInbox
    CloseWorkbookIfTransientConfirmWrites wbOps
    Exit Function
CleanFail:
    detailText = Err.Description
    Resume CleanExit
End Function

Private Function RunProcessorAppliesCase(ByRef fx As ConfirmWritesFixture, ByRef detailText As String) As Boolean
    Dim wbInv As Workbook
    Dim wbInbox As Workbook
    Dim loSku As ListObject
    Dim loInbox As ListObject
    Dim skuRow As Long
    Dim inboxRow As Long
    Dim processedCount As Long
    Dim report As String
    Dim statusText As String
    Dim qtyOnHand As Double

    On Error GoTo CleanFail
    processedCount = modProcessor.RunBatch(fx.WarehouseId, 500, report)

    Set wbInv = OpenWorkbookConfirmWrites(fx.RuntimeRoot & "\" & fx.WarehouseId & ".invSys.Data.Inventory.xlsb", False)
    Set wbInbox = OpenReceiveInboxWorkbookConfirmWrites(fx, detailText)
    If wbInv Is Nothing Or wbInbox Is Nothing Then GoTo CleanExit

    Set loSku = FindTableByNameConfirmWrites(wbInv, "tblSkuBalance")
    Set loInbox = FindTableByNameConfirmWrites(wbInbox, "tblInboxReceive")
    skuRow = FindRowByValueConfirmWrites(loSku, "SKU", "TEST-SKU-001")
    inboxRow = FindInboxRowByEventIdConfirmWrites(loInbox, mLastReceiveEventId)
    If skuRow = 0 Or inboxRow = 0 Then
        detailText = "Processor verification could not resolve the expected inventory or inbox row."
        GoTo CleanExit
    End If

    qtyOnHand = CDbl(GetTableValueConfirmWrites(loSku, skuRow, "QtyOnHand"))
    statusText = UCase$(CStr(GetTableValueConfirmWrites(loInbox, inboxRow, "Status")))
    If qtyOnHand <> 110# Then
        detailText = "QtyOnHand was " & CStr(qtyOnHand) & " instead of 110 after processor apply."
        GoTo CleanExit
    End If
    If statusText <> "PROCESSED" Then
        detailText = "Inbox status was " & statusText & " instead of PROCESSED. Report=" & report
        GoTo CleanExit
    End If

    RunProcessorAppliesCase = True
    detailText = "Inventory QtyOnHand reached 110 and the inbox row is PROCESSED. RunBatch processed " & CStr(processedCount) & " rows."

CleanExit:
    CloseWorkbookIfTransientConfirmWrites wbInbox
    CloseWorkbookIfTransientConfirmWrites wbInv
    Exit Function
CleanFail:
    detailText = Err.Description
    Resume CleanExit
End Function

Private Function RunSnapshotRefreshAfterPostCase(ByRef fx As ConfirmWritesFixture, ByRef detailText As String) As Boolean
    Dim wbOps As Workbook
    Dim loInv As ListObject
    Dim rowIndex As Long
    Dim report As String

    On Error GoTo CleanFail
    Set wbOps = OpenReceivingWorkbookConfirmWrites(fx.OperatorPath, detailText)
    If wbOps Is Nothing Then GoTo CleanExit

    If Not modOperatorReadModel.RefreshInventoryReadModelForWorkbook(wbOps, fx.WarehouseId, "LOCAL", report) Then
        detailText = report
        GoTo CleanExit
    End If

    Set loInv = FindTableByNameConfirmWrites(wbOps, "invSys")
    rowIndex = FindRowByValueConfirmWrites(loInv, "ITEM_CODE", "TEST-SKU-001")
    If rowIndex = 0 Then
        detailText = "tblReadModel did not contain TEST-SKU-001 after refresh."
        GoTo CleanExit
    End If
    If CDbl(GetTableValueConfirmWrites(loInv, rowIndex, "QtyOnHand")) <> 110# Then
        detailText = "tblReadModel QtyOnHand was not 110 after refresh."
        GoTo CleanExit
    End If
    If StrComp(CStr(GetTableValueConfirmWrites(loInv, rowIndex, "SourceType")), "LOCAL", vbTextCompare) <> 0 Then
        detailText = "tblReadModel SourceType was not LOCAL after refresh."
        GoTo CleanExit
    End If
    If ReadBoolLikeConfirmWrites(GetTableValueConfirmWrites(loInv, rowIndex, "IsStale")) Then
        detailText = "tblReadModel remained stale after refresh."
        GoTo CleanExit
    End If

    RunSnapshotRefreshAfterPostCase = True
    detailText = "Snapshot refresh showed TEST-SKU-001 at QtyOnHand = 110 with current LOCAL metadata."

CleanExit:
    CloseWorkbookIfTransientConfirmWrites wbOps
    Exit Function
CleanFail:
    detailText = Err.Description
    Resume CleanExit
End Function

Private Function RunIdempotentSetupCase(ByRef fx As ConfirmWritesFixture, ByRef detailText As String) As Boolean
    Dim spec As modTesterSetup.TesterSetupSpec
    Dim qtyAfter As Double
    Dim skuCount As Long
    Dim testerCapCount As Long
    Dim runtimeCapCount As Long
    Dim fileStampBefore As Date
    Dim fileSizeBefore As Long
    Dim fileStampAfter As Date
    Dim fileSizeAfter As Long

    On Error GoTo CleanFail
    fileStampBefore = FileDateTime(fx.OperatorPath)
    fileSizeBefore = FileLen(fx.OperatorPath)

    spec = BuildTesterSpecConfirmWrites(fx.TesterUserId, "123456", fx.WarehouseId, fx.StationId, fx.RuntimeRoot, fx.SharePointRoot)
    modWarehouseBootstrap.SetWarehouseBootstrapTemplateRootOverride fx.TemplateRoot
    If Not modTesterSetup.SetupTesterStation(spec) Then
        detailText = "Second SetupTesterStation failed: " & modTesterSetup.GetLastTesterSetupReport()
        GoTo CleanExit
    End If
    modWarehouseBootstrap.ClearWarehouseBootstrapTemplateRootOverride

    fileStampAfter = FileDateTime(fx.OperatorPath)
    fileSizeAfter = FileLen(fx.OperatorPath)
    qtyAfter = ReadSkuQtyConfirmWrites(fx.RuntimeRoot, fx.WarehouseId, "TEST-SKU-001")
    skuCount = CountSkuRowsConfirmWrites(fx.RuntimeRoot, fx.WarehouseId, "TEST-SKU-001")
    testerCapCount = CountCapabilityRowsConfirmWrites(fx.RuntimeRoot, fx.WarehouseId, fx.TesterUserId, "RECEIVE_POST", fx.StationId)
    runtimeCapCount = CountCapabilityRowsConfirmWrites(fx.RuntimeRoot, fx.WarehouseId, fx.RuntimeUserId, "RECEIVE_POST", fx.StationId)

    If qtyAfter <> 110# Then
        detailText = "Idempotent rerun changed TEST-SKU-001 QtyOnHand from 110."
        GoTo CleanExit
    End If
    If skuCount <> 1 Then
        detailText = "Idempotent rerun duplicated TEST-SKU-001 inventory rows."
        GoTo CleanExit
    End If
    If testerCapCount <> 1 Then
        detailText = "Idempotent rerun duplicated TESTER01 RECEIVE_POST capability rows."
        GoTo CleanExit
    End If
    If runtimeCapCount <> 1 Then
        detailText = "Idempotent rerun duplicated runtime RECEIVE_POST capability rows."
        GoTo CleanExit
    End If
    If fileStampAfter <> fileStampBefore Or fileSizeAfter <> fileSizeBefore Then
        detailText = "Receiving operator workbook was overwritten during rerun."
        GoTo CleanExit
    End If
    If InStr(1, modTesterSetup.GetLastTesterSetupReport(), "Runtime=EXISTING", vbTextCompare) = 0 Then
        detailText = "Rerun did not report reuse of the existing runtime."
        GoTo CleanExit
    End If

    RunIdempotentSetupCase = True
    detailText = "Rerun reused the runtime, preserved the workbook file, and did not duplicate seed or capability rows."

CleanExit:
    modWarehouseBootstrap.ClearWarehouseBootstrapTemplateRootOverride
    Exit Function
CleanFail:
    detailText = Err.Description
    Resume CleanExit
End Function

Private Function RunSharePointUnavailableCase(ByRef fx As ConfirmWritesFixture, ByRef detailText As String) As Boolean
    Dim wbOps As Workbook
    Dim report As String
    Dim bannerText As String

    On Error GoTo CleanFail
    Set wbOps = OpenReceivingWorkbookConfirmWrites(fx.OperatorPath, detailText)
    If wbOps Is Nothing Then GoTo CleanExit

    If Not modOperatorReadModel.RefreshInventoryReadModelFromSharePointForWorkbook(wbOps, fx.WarehouseId, report) Then
        detailText = report
        GoTo CleanExit
    End If

    bannerText = GetReadModelStatusBannerTextConfirmWrites(wbOps)
    If InStr(1, bannerText, "INVENTORY SNAPSHOT STALE", vbTextCompare) = 0 Then
        detailText = "Read-model banner did not show STALE metadata. Banner=" & bannerText
        GoTo CleanExit
    End If
    If InStr(1, bannerText, "ERROR", vbTextCompare) > 0 Then
        detailText = "Read-model banner showed ERROR instead of stale metadata. Banner=" & bannerText
        GoTo CleanExit
    End If

    RunSharePointUnavailableCase = True
    detailText = "Offline SharePoint refresh completed without hard failure and rendered a stale read-model banner: " & bannerText

CleanExit:
    CloseWorkbookIfTransientConfirmWrites wbOps
    Exit Function
CleanFail:
    detailText = Err.Description
    Resume CleanExit
End Function

Private Sub RecordFixtureFailureCases(ByVal detailText As String)
    RecordConfirmWritesCase "ReadinessCheck_OK", False, detailText
    RecordConfirmWritesCase "ReadinessCheck_MissingCapability", False, detailText
    RecordConfirmWritesCase "RefreshInventory_Loads", False, detailText
    RecordConfirmWritesCase "ConfirmWrites_SingleRow", False, detailText
    RecordConfirmWritesCase "ProcessorApplies", False, detailText
    RecordConfirmWritesCase "SnapshotRefreshAfterPost", False, detailText
    RecordConfirmWritesCase "IdempotentSetup", False, detailText
End Sub

Private Function BuildTesterSpecConfirmWrites(ByVal userId As String, _
                                              ByVal pinText As String, _
                                              ByVal warehouseId As String, _
                                              ByVal stationId As String, _
                                              ByVal pathLocal As String, _
                                              ByVal pathSharePointRoot As String) As modTesterSetup.TesterSetupSpec
    Dim spec As modTesterSetup.TesterSetupSpec

    spec.UserId = userId
    spec.PinHash = modAuth.HashUserCredential(pinText)
    spec.WarehouseId = warehouseId
    spec.StationId = stationId
    spec.PathLocal = pathLocal
    spec.PathSharePointRoot = pathSharePointRoot
    BuildTesterSpecConfirmWrites = spec
End Function

Private Function EnsureRuntimeUserCapabilitiesConfirmWrites(ByRef fx As ConfirmWritesFixture, ByRef detailText As String) As Boolean
    If Trim$(fx.RuntimeUserId) = "" Then
        detailText = "Current runtime user could not be resolved."
        Exit Function
    End If

    EnsureUserCapabilityActiveConfirmWrites fx.RuntimeRoot, fx.WarehouseId, fx.RuntimeUserId, "RECEIVE_POST", fx.StationId
    EnsureUserCapabilityActiveConfirmWrites fx.RuntimeRoot, fx.WarehouseId, fx.RuntimeUserId, "RECEIVE_VIEW", fx.StationId
    EnsureUserCapabilityActiveConfirmWrites fx.RuntimeRoot, fx.WarehouseId, fx.RuntimeUserId, "READMODEL_REFRESH", fx.StationId
    EnsureUserCapabilityActiveConfirmWrites fx.RuntimeRoot, fx.WarehouseId, "svc_processor", "INBOX_PROCESS", "*"
    EnsureRuntimeUserCapabilitiesConfirmWrites = True
End Function

Private Sub EnsureUserCapabilityActiveConfirmWrites(ByVal runtimeRoot As String, _
                                                    ByVal warehouseId As String, _
                                                    ByVal userId As String, _
                                                    ByVal capability As String, _
                                                    ByVal stationId As String)
    SetCapabilityStatusConfirmWrites runtimeRoot, warehouseId, userId, capability, stationId, "ACTIVE"
End Sub

Private Sub SetCapabilityStatusConfirmWrites(ByVal runtimeRoot As String, _
                                             ByVal warehouseId As String, _
                                             ByVal userId As String, _
                                             ByVal capability As String, _
                                             ByVal stationId As String, _
                                             ByVal statusText As String)
    Dim wbAuth As Workbook
    Dim loUsers As ListObject
    Dim loCaps As ListObject
    Dim userRow As Long
    Dim capRow As Long

    Set wbAuth = OpenWorkbookConfirmWrites(runtimeRoot & "\" & warehouseId & ".invSys.Auth.xlsb", False)
    If wbAuth Is Nothing Then Exit Sub

    Set loUsers = FindTableByNameConfirmWrites(wbAuth, "tblUsers")
    Set loCaps = FindTableByNameConfirmWrites(wbAuth, "tblCapabilities")
    If loUsers Is Nothing Or loCaps Is Nothing Then GoTo CleanExit

    userRow = EnsureUserRowConfirmWrites(loUsers, userId)
    SetTableValueConfirmWrites loUsers, userRow, "UserId", userId
    SetTableValueConfirmWrites loUsers, userRow, "DisplayName", userId
    SetTableValueConfirmWrites loUsers, userRow, "Status", "Active"
    If Trim$(CStr(GetTableValueConfirmWrites(loUsers, userRow, "PinHash"))) = "" Then
        SetTableValueConfirmWrites loUsers, userRow, "PinHash", modAuth.HashUserCredential("123456")
    End If

    capRow = EnsureCapabilityRowConfirmWrites(loCaps, userId, capability, warehouseId, stationId)
    SetTableValueConfirmWrites loCaps, capRow, "UserId", userId
    SetTableValueConfirmWrites loCaps, capRow, "Capability", capability
    SetTableValueConfirmWrites loCaps, capRow, "WarehouseId", warehouseId
    SetTableValueConfirmWrites loCaps, capRow, "StationId", stationId
    SetTableValueConfirmWrites loCaps, capRow, "Status", statusText
    wbAuth.Save

CleanExit:
    CloseWorkbookIfTransientConfirmWrites wbAuth
End Sub

Private Function EnsureUserRowConfirmWrites(ByVal lo As ListObject, ByVal userId As String) As Long
    EnsureUserRowConfirmWrites = FindRowByValueConfirmWrites(lo, "UserId", userId)
    If EnsureUserRowConfirmWrites > 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then
        lo.ListRows.Add
        EnsureUserRowConfirmWrites = 1
    Else
        EnsureUserRowConfirmWrites = lo.ListRows.Add.Index
    End If
End Function

Private Function EnsureCapabilityRowConfirmWrites(ByVal lo As ListObject, _
                                                  ByVal userId As String, _
                                                  ByVal capability As String, _
                                                  ByVal warehouseId As String, _
                                                  ByVal stationId As String) As Long
    EnsureCapabilityRowConfirmWrites = FindCapabilityRowConfirmWrites(lo, userId, capability, warehouseId, stationId)
    If EnsureCapabilityRowConfirmWrites > 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then
        lo.ListRows.Add
        EnsureCapabilityRowConfirmWrites = 1
    Else
        EnsureCapabilityRowConfirmWrites = lo.ListRows.Add.Index
    End If
End Function

Private Function FindCapabilityRowConfirmWrites(ByVal lo As ListObject, _
                                                ByVal userId As String, _
                                                ByVal capability As String, _
                                                ByVal warehouseId As String, _
                                                ByVal stationId As String) As Long
    Dim i As Long

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To lo.ListRows.Count
        If StrComp(CStr(GetTableValueConfirmWrites(lo, i, "UserId")), userId, vbTextCompare) = 0 _
           And StrComp(CStr(GetTableValueConfirmWrites(lo, i, "Capability")), capability, vbTextCompare) = 0 _
           And StrComp(CStr(GetTableValueConfirmWrites(lo, i, "WarehouseId")), warehouseId, vbTextCompare) = 0 _
           And StrComp(CStr(GetTableValueConfirmWrites(lo, i, "StationId")), stationId, vbTextCompare) = 0 Then
            FindCapabilityRowConfirmWrites = i
            Exit Function
        End If
    Next i
End Function

Private Function CountCapabilityRowsConfirmWrites(ByVal runtimeRoot As String, _
                                                  ByVal warehouseId As String, _
                                                  ByVal userId As String, _
                                                  ByVal capability As String, _
                                                  ByVal stationId As String) As Long
    Dim wbAuth As Workbook
    Dim loCaps As ListObject
    Dim i As Long

    Set wbAuth = OpenWorkbookConfirmWrites(runtimeRoot & "\" & warehouseId & ".invSys.Auth.xlsb", True)
    If wbAuth Is Nothing Then Exit Function

    Set loCaps = FindTableByNameConfirmWrites(wbAuth, "tblCapabilities")
    If loCaps Is Nothing Or loCaps.DataBodyRange Is Nothing Then GoTo CleanExit

    For i = 1 To loCaps.ListRows.Count
        If StrComp(CStr(GetTableValueConfirmWrites(loCaps, i, "UserId")), userId, vbTextCompare) = 0 _
           And StrComp(CStr(GetTableValueConfirmWrites(loCaps, i, "Capability")), capability, vbTextCompare) = 0 _
           And StrComp(CStr(GetTableValueConfirmWrites(loCaps, i, "StationId")), stationId, vbTextCompare) = 0 Then
            CountCapabilityRowsConfirmWrites = CountCapabilityRowsConfirmWrites + 1
        End If
    Next i

CleanExit:
    CloseWorkbookIfTransientConfirmWrites wbAuth
End Function

Private Function OpenReceivingWorkbookConfirmWrites(ByVal workbookPath As String, ByRef detailText As String) As Workbook
    Set OpenReceivingWorkbookConfirmWrites = OpenWorkbookConfirmWrites(workbookPath, False)
    If OpenReceivingWorkbookConfirmWrites Is Nothing Then
        detailText = "Receiving workbook could not be opened: " & workbookPath
        Exit Function
    End If
    OpenReceivingWorkbookConfirmWrites.Activate
    modReceivingInit.EnsureReceivingSurfaceForWorkbook OpenReceivingWorkbookConfirmWrites
End Function

Private Function OpenReceiveInboxWorkbookConfirmWrites(ByRef fx As ConfirmWritesFixture, ByRef detailText As String) As Workbook
    Dim inboxPath As String
    Dim report As String

    inboxPath = modRoleEventWriter.ResolveInboxWorkbookPath("RECEIVE", fx.WarehouseId, fx.StationId, report)
    If inboxPath = "" Then
        detailText = report
        Exit Function
    End If

    Set OpenReceiveInboxWorkbookConfirmWrites = OpenWorkbookConfirmWrites(inboxPath, False)
    If OpenReceiveInboxWorkbookConfirmWrites Is Nothing Then detailText = "Receive inbox workbook could not be opened: " & inboxPath
End Function

Private Function OpenWorkbookConfirmWrites(ByVal workbookPath As String, ByVal readOnlyMode As Boolean) As Workbook
    Dim wb As Workbook

    Set wb = FindOpenWorkbookByPathConfirmWrites(workbookPath)
    If wb Is Nothing Then
        If Len(Dir$(workbookPath, vbNormal)) = 0 Then Exit Function
        Set wb = Application.Workbooks.Open(Filename:=workbookPath, UpdateLinks:=0, ReadOnly:=readOnlyMode, IgnoreReadOnlyRecommended:=True, Notify:=False, AddToMru:=False)
    End If
    Set OpenWorkbookConfirmWrites = wb
End Function

Private Function ReopenWorkbookConfirmWrites(ByVal wb As Workbook) As Workbook
    Dim workbookPath As String

    If wb Is Nothing Then Exit Function
    workbookPath = wb.FullName
    CloseWorkbookIfTransientConfirmWrites wb
    Set ReopenWorkbookConfirmWrites = OpenWorkbookConfirmWrites(workbookPath, False)
End Function

Private Sub CloseWorkbookByPathConfirmWrites(ByVal workbookPath As String)
    Dim wb As Workbook

    Set wb = FindOpenWorkbookByPathConfirmWrites(workbookPath)
    CloseWorkbookIfTransientConfirmWrites wb
End Sub

Private Sub CloseWorkbookIfTransientConfirmWrites(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    If Not wb.ReadOnly Then
        If wb.Saved = False Then wb.Save
    End If
    wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

Private Function FindOpenWorkbookByPathConfirmWrites(ByVal workbookPath As String) As Workbook
    Dim wb As Workbook

    workbookPath = Trim$(workbookPath)
    If workbookPath = "" Then Exit Function
    For Each wb In Application.Workbooks
        If StrComp(Trim$(wb.FullName), workbookPath, vbTextCompare) = 0 Then
            Set FindOpenWorkbookByPathConfirmWrites = wb
            Exit Function
        End If
    Next wb
End Function

Private Function FindTableByNameConfirmWrites(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function
    On Error Resume Next
    For Each ws In wb.Worksheets
        Set FindTableByNameConfirmWrites = ws.ListObjects(tableName)
        If Not FindTableByNameConfirmWrites Is Nothing Then Exit Function
    Next ws
    On Error GoTo 0
End Function

Private Function FindRowByValueConfirmWrites(ByVal lo As ListObject, ByVal columnName As String, ByVal expectedValue As String) As Long
    Dim i As Long

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To lo.ListRows.Count
        If StrComp(CStr(GetTableValueConfirmWrites(lo, i, columnName)), expectedValue, vbTextCompare) = 0 Then
            FindRowByValueConfirmWrites = i
            Exit Function
        End If
    Next i
End Function

Private Function FindInboxRowByEventIdConfirmWrites(ByVal lo As ListObject, ByVal eventId As String) As Long
    FindInboxRowByEventIdConfirmWrites = FindRowByValueConfirmWrites(lo, "EventID", eventId)
End Function

Private Function FindLastInboxRowByNoteConfirmWrites(ByVal lo As ListObject, ByVal noteText As String) As Long
    Dim i As Long

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    For i = lo.ListRows.Count To 1 Step -1
        If InStr(1, CStr(GetTableValueConfirmWrites(lo, i, "Note")), noteText, vbTextCompare) > 0 Then
            FindLastInboxRowByNoteConfirmWrites = i
            Exit Function
        End If
    Next i
End Function

Private Function GetColumnIndexConfirmWrites(ByVal lo As ListObject, ByVal columnName As String) As Long
    Dim i As Long

    If lo Is Nothing Then Exit Function
    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).Name, columnName, vbTextCompare) = 0 Then
            GetColumnIndexConfirmWrites = i
            Exit Function
        End If
    Next i
End Function

Private Function GetTableValueConfirmWrites(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As Variant
    Dim colIndex As Long

    colIndex = GetColumnIndexConfirmWrites(lo, columnName)
    If colIndex = 0 Then Exit Function
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    If rowIndex < 1 Or rowIndex > lo.ListRows.Count Then Exit Function
    GetTableValueConfirmWrites = lo.DataBodyRange.Cells(rowIndex, colIndex).Value
End Function

Private Sub SetTableValueConfirmWrites(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String, ByVal valueOut As Variant)
    Dim colIndex As Long

    colIndex = GetColumnIndexConfirmWrites(lo, columnName)
    If colIndex = 0 Then Exit Sub
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    If rowIndex < 1 Or rowIndex > lo.ListRows.Count Then Exit Sub
    lo.DataBodyRange.Cells(rowIndex, colIndex).Value = valueOut
End Sub

Private Function ReadSkuQtyConfirmWrites(ByVal runtimeRoot As String, ByVal warehouseId As String, ByVal skuValue As String) As Double
    Dim wbInv As Workbook
    Dim loSku As ListObject
    Dim rowIndex As Long

    Set wbInv = OpenWorkbookConfirmWrites(runtimeRoot & "\" & warehouseId & ".invSys.Data.Inventory.xlsb", True)
    If wbInv Is Nothing Then Exit Function

    Set loSku = FindTableByNameConfirmWrites(wbInv, "tblSkuBalance")
    rowIndex = FindRowByValueConfirmWrites(loSku, "SKU", skuValue)
    If rowIndex > 0 And IsNumeric(GetTableValueConfirmWrites(loSku, rowIndex, "QtyOnHand")) Then
        ReadSkuQtyConfirmWrites = CDbl(GetTableValueConfirmWrites(loSku, rowIndex, "QtyOnHand"))
    End If

    CloseWorkbookIfTransientConfirmWrites wbInv
End Function

Private Function CountSkuRowsConfirmWrites(ByVal runtimeRoot As String, ByVal warehouseId As String, ByVal skuValue As String) As Long
    Dim wbInv As Workbook
    Dim loSku As ListObject
    Dim i As Long

    Set wbInv = OpenWorkbookConfirmWrites(runtimeRoot & "\" & warehouseId & ".invSys.Data.Inventory.xlsb", True)
    If wbInv Is Nothing Then Exit Function

    Set loSku = FindTableByNameConfirmWrites(wbInv, "tblSkuBalance")
    If loSku Is Nothing Or loSku.DataBodyRange Is Nothing Then GoTo CleanExit
    For i = 1 To loSku.ListRows.Count
        If StrComp(CStr(GetTableValueConfirmWrites(loSku, i, "SKU")), skuValue, vbTextCompare) = 0 Then
            CountSkuRowsConfirmWrites = CountSkuRowsConfirmWrites + 1
        End If
    Next i

CleanExit:
    CloseWorkbookIfTransientConfirmWrites wbInv
End Function

Private Function GetReadModelStatusBannerTextConfirmWrites(ByVal wb As Workbook) As String
    Dim ws As Worksheet
    Dim shp As Shape

    If wb Is Nothing Then Exit Function
    On Error Resume Next
    Set ws = wb.Worksheets("InventoryManagement")
    If Not ws Is Nothing Then Set shp = ws.Shapes("invSysReadModelStatus")
    On Error GoTo 0
    If shp Is Nothing Then Exit Function
    GetReadModelStatusBannerTextConfirmWrites = Trim$(shp.TextFrame.Characters.Text)
End Function

Private Function ResolveCurrentRuntimeUserIdConfirmWrites() As String
    ResolveCurrentRuntimeUserIdConfirmWrites = Trim$(modRoleEventWriter.ResolveCurrentUserId())
    If ResolveCurrentRuntimeUserIdConfirmWrites = "" Then ResolveCurrentRuntimeUserIdConfirmWrites = Trim$(Application.UserName)
End Function

Private Function BuildConfirmWritesTempRoot(ByVal suffix As String) As String
    Randomize
    BuildConfirmWritesTempRoot = Environ$("TEMP") & "\invSys_confirm_writes_tester_" & suffix & "_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & Format$(CLng(Rnd() * 10000), "0000")
End Function

Private Sub EnsureFolderRecursiveConfirmWrites(ByVal folderPath As String)
    Dim parentPath As String
    Dim slashPos As Long

    folderPath = Trim$(Replace$(folderPath, "/", "\"))
    If folderPath = "" Then Exit Sub
    If Len(Dir$(folderPath, vbDirectory)) > 0 Then Exit Sub

    slashPos = InStrRev(folderPath, "\")
    If slashPos > 3 Then
        parentPath = Left$(folderPath, slashPos - 1)
        If Len(Dir$(parentPath, vbDirectory)) = 0 Then EnsureFolderRecursiveConfirmWrites parentPath
    End If
    If Len(Dir$(folderPath, vbDirectory)) = 0 Then MkDir folderPath
End Sub

Private Sub DeleteFolderRecursiveConfirmWrites(ByVal folderPath As String)
    On Error Resume Next
    If Trim$(folderPath) <> "" Then CreateObject("Scripting.FileSystemObject").DeleteFolder folderPath, True
    On Error GoTo 0
End Sub

Private Function FileExistsConfirmWrites(ByVal filePath As String) As Boolean
    filePath = Trim$(Replace$(filePath, "/", "\"))
    If filePath = "" Then Exit Function
    FileExistsConfirmWrites = (Len(Dir$(filePath, vbNormal)) > 0)
End Function

Private Function ReadBoolLikeConfirmWrites(ByVal valueIn As Variant) As Boolean
    Dim valueText As String

    valueText = UCase$(Trim$(CStr(valueIn)))
    ReadBoolLikeConfirmWrites = (valueText = "TRUE" Or valueText = "YES" Or valueText = "1")
End Function

Private Function ValueOrDefaultConfirmWrites(ByVal valueIn As Variant, ByVal defaultText As String) As String
    ValueOrDefaultConfirmWrites = Trim$(CStr(valueIn))
    If ValueOrDefaultConfirmWrites = "" Then ValueOrDefaultConfirmWrites = defaultText
End Function

Private Sub ResetConfirmWritesEvidence()
    mCaseCount = 0
    Erase mCaseNames
    Erase mCaseResults
    Erase mCaseDetails
    mSummary = vbNullString
    mRuntimeUserId = vbNullString
    mTesterUserId = vbNullString
    mLastReceiveEventId = vbNullString
    mLastReceiveRefNumber = vbNullString
End Sub

Private Sub RecordConfirmWritesCase(ByVal caseName As String, ByVal passed As Boolean, ByVal detailText As String)
    mCaseCount = mCaseCount + 1
    ReDim Preserve mCaseNames(1 To mCaseCount)
    ReDim Preserve mCaseResults(1 To mCaseCount)
    ReDim Preserve mCaseDetails(1 To mCaseCount)
    mCaseNames(mCaseCount) = caseName
    mCaseResults(mCaseCount) = IIf(passed, "PASS", "FAIL")
    mCaseDetails(mCaseCount) = SafeConfirmWritesText(detailText)
End Sub

Private Function AllConfirmWritesCasesPassed() As Boolean
    Dim i As Long

    If mCaseCount = 0 Then Exit Function
    For i = 1 To mCaseCount
        If mCaseResults(i) <> "PASS" Then Exit Function
    Next i
    AllConfirmWritesCasesPassed = True
End Function

Private Function SafeConfirmWritesText(ByVal textIn As String) As String
    SafeConfirmWritesText = Replace$(Replace$(Trim$(textIn), vbCr, " "), vbLf, " ")
End Function
