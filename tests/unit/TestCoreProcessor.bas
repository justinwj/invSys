Attribute VB_Name = "TestCoreProcessor"
Option Explicit

Public Sub RunProcessorTests()
    Dim passed As Long
    Dim failed As Long

    Tally TestRunBatch_ProcessesInboxRow(), passed, failed
    Tally TestRunBatch_DuplicateMarkedSkipDup(), passed, failed
    Tally TestRunBatch_ProcessesShipRow(), passed, failed
    Tally TestRunBatch_ProcessesProdConsumeRow(), passed, failed
    Tally TestRunBatch_ProcessesProdCompleteRow(), passed, failed

    Debug.Print "Core.Processor tests - Passed: " & passed & " Failed: " & failed
End Sub

Public Function TestRunBatch_ProcessesInboxRow() As Long
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim wbInv As Workbook
    Dim wbInbox As Workbook
    Dim report As String
    Dim loInbox As ListObject
    Dim loLog As ListObject
    Dim loApplied As ListObject

    On Error GoTo CleanFail
    Set wbCfg = TestPhase2Helpers.BuildPhase2ConfigWorkbook("WH1", "S1")
    Set wbAuth = TestPhase2Helpers.BuildPhase2AuthWorkbook("WH1")
    Set wbInv = TestPhase2Helpers.BuildPhase2InventoryWorkbook("WH1", Array("SKU-001"))
    Set wbInbox = TestPhase2Helpers.BuildPhase2InboxWorkbook("S1")

    TestPhase2Helpers.AddCapability wbAuth, "user1", "RECEIVE_POST", "WH1", "S1", "ACTIVE"
    TestPhase2Helpers.AddCapability wbAuth, "svc_processor", "INBOX_PROCESS", "WH1", "*", "ACTIVE"
    TestPhase2Helpers.AddInboxReceiveRow wbInbox, "EVT-PROC-001", Now, "WH1", "S1", "user1", "SKU-001", 7, "A1", "processor test"
    If modProcessor.RunBatch("WH1", 500, report) <> 1 Then GoTo CleanExit

    Set loInbox = wbInbox.Worksheets("InboxReceive").ListObjects("tblInboxReceive")
    Set loLog = wbInv.Worksheets("InventoryLog").ListObjects("tblInventoryLog")
    Set loApplied = wbInv.Worksheets("AppliedEvents").ListObjects("tblAppliedEvents")

    If CStr(TestPhase2Helpers.GetRowValue(loInbox, 1, "Status")) <> "PROCESSED" Then GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(loLog, 2, "EventID")) <> "EVT-PROC-001" Then GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(loApplied, 2, "EventID")) <> "EVT-PROC-001" Then GoTo CleanExit

    TestRunBatch_ProcessesInboxRow = 1

CleanExit:
    TestPhase2Helpers.CloseNoSave wbInbox
    TestPhase2Helpers.CloseNoSave wbInv
    TestPhase2Helpers.CloseNoSave wbAuth
    TestPhase2Helpers.CloseNoSave wbCfg
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestRunBatch_DuplicateMarkedSkipDup() As Long
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim wbInv As Workbook
    Dim wbInbox As Workbook
    Dim report As String
    Dim loInbox As ListObject
    Dim loLog As ListObject

    On Error GoTo CleanFail
    Set wbCfg = TestPhase2Helpers.BuildPhase2ConfigWorkbook("WH1", "S1")
    Set wbAuth = TestPhase2Helpers.BuildPhase2AuthWorkbook("WH1")
    Set wbInv = TestPhase2Helpers.BuildPhase2InventoryWorkbook("WH1", Array("SKU-001"))
    Set wbInbox = TestPhase2Helpers.BuildPhase2InboxWorkbook("S1")

    TestPhase2Helpers.AddCapability wbAuth, "user1", "RECEIVE_POST", "WH1", "S1", "ACTIVE"
    TestPhase2Helpers.AddCapability wbAuth, "svc_processor", "INBOX_PROCESS", "WH1", "*", "ACTIVE"
    TestPhase2Helpers.AddInboxReceiveRow wbInbox, "EVT-PROC-002", Now, "WH1", "S1", "user1", "SKU-001", 2
    TestPhase2Helpers.AddInboxReceiveRow wbInbox, "EVT-PROC-002", DateAdd("s", 1, Now), "WH1", "S1", "user1", "SKU-001", 2
    Call modProcessor.RunBatch("WH1", 500, report)

    Set loInbox = wbInbox.Worksheets("InboxReceive").ListObjects("tblInboxReceive")
    Set loLog = wbInv.Worksheets("InventoryLog").ListObjects("tblInventoryLog")

    If CStr(TestPhase2Helpers.GetRowValue(loInbox, 1, "Status")) <> "PROCESSED" Then GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(loInbox, 2, "Status")) <> "SKIP_DUP" Then GoTo CleanExit
    If loLog.ListRows.Count <> 2 Then GoTo CleanExit

    TestRunBatch_DuplicateMarkedSkipDup = 1

CleanExit:
    TestPhase2Helpers.CloseNoSave wbInbox
    TestPhase2Helpers.CloseNoSave wbInv
    TestPhase2Helpers.CloseNoSave wbAuth
    TestPhase2Helpers.CloseNoSave wbCfg
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestRunBatch_ProcessesShipRow() As Long
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim wbInv As Workbook
    Dim wbInbox As Workbook
    Dim report As String
    Dim payloadJson As String
    Dim loInbox As ListObject
    Dim loLog As ListObject

    On Error GoTo CleanFail
    Set wbCfg = TestPhase2Helpers.BuildPhase2ConfigWorkbook("WH1", "S1")
    Set wbAuth = TestPhase2Helpers.BuildPhase2AuthWorkbook("WH1")
    Set wbInv = TestPhase2Helpers.BuildPhase2InventoryWorkbook("WH1", Array("SKU-001", "SKU-002"))
    Set wbInbox = TestPhase2Helpers.BuildShipInboxWorkbook("S1")

    TestPhase2Helpers.AddCapability wbAuth, "user1", "SHIP_POST", "WH1", "S1", "ACTIVE"
    TestPhase2Helpers.AddCapability wbAuth, "svc_processor", "INBOX_PROCESS", "WH1", "*", "ACTIVE"
    payloadJson = TestPhase2Helpers.BuildPayloadJson( _
        TestPhase2Helpers.CreatePayloadItem(101, "SKU-001", 4, "DOCK", "shipment A"), _
        TestPhase2Helpers.CreatePayloadItem(102, "SKU-002", 1, "DOCK", "shipment B"))
    TestPhase2Helpers.AddInboxShipRow wbInbox, "EVT-SHIP-PROC-001", Now, "WH1", "S1", "user1", payloadJson, "ship batch"
    If modProcessor.RunBatch("WH1", 500, report) <> 1 Then GoTo CleanExit

    Set loInbox = wbInbox.Worksheets("InboxShip").ListObjects("tblInboxShip")
    Set loLog = wbInv.Worksheets("InventoryLog").ListObjects("tblInventoryLog")
    If CStr(TestPhase2Helpers.GetRowValue(loInbox, 1, "Status")) <> "PROCESSED" Then GoTo CleanExit
    If loLog.ListRows.Count <> 3 Then GoTo CleanExit
    If CDbl(TestPhase2Helpers.GetRowValue(loLog, 2, "QtyDelta")) <> -4 Then GoTo CleanExit
    If CDbl(TestPhase2Helpers.GetRowValue(loLog, 3, "QtyDelta")) <> -1 Then GoTo CleanExit

    TestRunBatch_ProcessesShipRow = 1

CleanExit:
    TestPhase2Helpers.CloseNoSave wbInbox
    TestPhase2Helpers.CloseNoSave wbInv
    TestPhase2Helpers.CloseNoSave wbAuth
    TestPhase2Helpers.CloseNoSave wbCfg
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestRunBatch_ProcessesProdConsumeRow() As Long
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim wbInv As Workbook
    Dim wbInbox As Workbook
    Dim report As String
    Dim payloadJson As String
    Dim loInbox As ListObject
    Dim loLog As ListObject

    On Error GoTo CleanFail
    Set wbCfg = TestPhase2Helpers.BuildPhase2ConfigWorkbook("WH1", "S1")
    Set wbAuth = TestPhase2Helpers.BuildPhase2AuthWorkbook("WH1")
    Set wbInv = TestPhase2Helpers.BuildPhase2InventoryWorkbook("WH1", Array("SKU-COMP", "SKU-FG"))
    Set wbInbox = TestPhase2Helpers.BuildProductionInboxWorkbook("S1")

    TestPhase2Helpers.AddCapability wbAuth, "user1", "PROD_POST", "WH1", "S1", "ACTIVE"
    TestPhase2Helpers.AddCapability wbAuth, "svc_processor", "INBOX_PROCESS", "WH1", "*", "ACTIVE"
    payloadJson = TestPhase2Helpers.BuildPayloadJson( _
        TestPhase2Helpers.CreatePayloadItem(201, "SKU-COMP", 3, "LINE1", "used", "USED"), _
        TestPhase2Helpers.CreatePayloadItem(202, "SKU-FG", 1, "LINE1", "made", "MADE"))
    TestPhase2Helpers.AddInboxProductionRow wbInbox, "EVT-PROD-PROC-001", EVENT_TYPE_PROD_CONSUME, Now, "WH1", "S1", "user1", payloadJson, "prod consume"
    If modProcessor.RunBatch("WH1", 500, report) <> 1 Then GoTo CleanExit

    Set loInbox = wbInbox.Worksheets("InboxProd").ListObjects("tblInboxProd")
    Set loLog = wbInv.Worksheets("InventoryLog").ListObjects("tblInventoryLog")
    If CStr(TestPhase2Helpers.GetRowValue(loInbox, 1, "Status")) <> "PROCESSED" Then GoTo CleanExit
    If loLog.ListRows.Count <> 3 Then GoTo CleanExit
    If CDbl(TestPhase2Helpers.GetRowValue(loLog, 2, "QtyDelta")) <> -3 Then GoTo CleanExit
    If CDbl(TestPhase2Helpers.GetRowValue(loLog, 3, "QtyDelta")) <> 1 Then GoTo CleanExit

    TestRunBatch_ProcessesProdConsumeRow = 1

CleanExit:
    TestPhase2Helpers.CloseNoSave wbInbox
    TestPhase2Helpers.CloseNoSave wbInv
    TestPhase2Helpers.CloseNoSave wbAuth
    TestPhase2Helpers.CloseNoSave wbCfg
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestRunBatch_ProcessesProdCompleteRow() As Long
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim wbInv As Workbook
    Dim wbInbox As Workbook
    Dim report As String
    Dim payloadJson As String
    Dim loInbox As ListObject
    Dim loLog As ListObject

    On Error GoTo CleanFail
    Set wbCfg = TestPhase2Helpers.BuildPhase2ConfigWorkbook("WH1", "S1")
    Set wbAuth = TestPhase2Helpers.BuildPhase2AuthWorkbook("WH1")
    Set wbInv = TestPhase2Helpers.BuildPhase2InventoryWorkbook("WH1", Array("SKU-FG"))
    Set wbInbox = TestPhase2Helpers.BuildProductionInboxWorkbook("S1")

    TestPhase2Helpers.AddCapability wbAuth, "user1", "PROD_POST", "WH1", "S1", "ACTIVE"
    TestPhase2Helpers.AddCapability wbAuth, "svc_processor", "INBOX_PROCESS", "WH1", "*", "ACTIVE"
    payloadJson = TestPhase2Helpers.BuildPayloadJson( _
        TestPhase2Helpers.CreatePayloadItem(301, "SKU-FG", 8, "FG", "complete", "COMPLETE"))
    TestPhase2Helpers.AddInboxProductionRow wbInbox, "EVT-PROD-PROC-002", EVENT_TYPE_PROD_COMPLETE, Now, "WH1", "S1", "user1", payloadJson, "prod complete"
    If modProcessor.RunBatch("WH1", 500, report) <> 1 Then GoTo CleanExit

    Set loInbox = wbInbox.Worksheets("InboxProd").ListObjects("tblInboxProd")
    Set loLog = wbInv.Worksheets("InventoryLog").ListObjects("tblInventoryLog")
    If CStr(TestPhase2Helpers.GetRowValue(loInbox, 1, "Status")) <> "PROCESSED" Then GoTo CleanExit
    If loLog.ListRows.Count <> 2 Then GoTo CleanExit
    If CDbl(TestPhase2Helpers.GetRowValue(loLog, 2, "QtyDelta")) <> 8 Then GoTo CleanExit

    TestRunBatch_ProcessesProdCompleteRow = 1

CleanExit:
    TestPhase2Helpers.CloseNoSave wbInbox
    TestPhase2Helpers.CloseNoSave wbInv
    TestPhase2Helpers.CloseNoSave wbAuth
    TestPhase2Helpers.CloseNoSave wbCfg
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Private Sub Tally(ByVal testResult As Long, ByRef passed As Long, ByRef failed As Long)
    If testResult = 1 Then
        passed = passed + 1
    Else
        failed = failed + 1
    End If
End Sub
