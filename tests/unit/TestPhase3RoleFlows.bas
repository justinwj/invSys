Attribute VB_Name = "TestPhase3RoleFlows"
Option Explicit

Public Sub RunPhase3RoleFlowTests()
    Dim passed As Long
    Dim failed As Long

    Tally TestReceivingRoleFlow_QueuesAndProcessesEvent(), passed, failed
    Tally TestShippingRoleFlow_QueuesAndProcessesEvent(), passed, failed
    Tally TestProductionRoleFlow_QueuesAndProcessesEvent(), passed, failed

    Debug.Print "Phase 3 role flow tests - Passed: " & passed & " Failed: " & failed
End Sub

Public Function TestReceivingRoleFlow_QueuesAndProcessesEvent() As Long
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim wbInv As Workbook
    Dim wbInbox As Workbook
    Dim tempRoot As String
    Dim report As String
    Dim errorMessage As String
    Dim currentUserId As String
    Dim loInbox As ListObject
    Dim loLog As ListObject
    Dim wbRole As Workbook
    Dim inboxRow As Long

    On Error GoTo CleanFail
    currentUserId = modRoleEventWriter.ResolveCurrentUserId()
    tempRoot = TestPhase2Helpers.BuildUniqueTestFolder("Phase3Receive")
    Set wbCfg = TestPhase2Helpers.BuildPhase2ConfigWorkbook("WHR3", "S1", "RECEIVE")
    TestPhase2Helpers.SetWarehouseConfigValue wbCfg, "PathDataRoot", tempRoot
    Set wbAuth = TestPhase2Helpers.BuildPhase2AuthWorkbook("WHR3")
    TestPhase2Helpers.AddCapability wbAuth, currentUserId, "RECEIVE_POST", "WHR3", "S1", "ACTIVE"
    TestPhase2Helpers.AddCapability wbAuth, "svc_processor", "INBOX_PROCESS", "WHR3", "*", "ACTIVE"
    Set wbInv = TestPhase2Helpers.BuildCanonicalInventoryWorkbook("WHR3", tempRoot, Array("SKU-001"))
    Set wbInbox = TestPhase2Helpers.BuildCanonicalReceiveInboxWorkbook("S1", tempRoot)

    Set wbRole = Application.Workbooks.Add
    SetupReceivingRoleScaffold wbRole
    If Not modReceivingEventCreator.QueueReceiveEventsFromWorkbook(wbRole, errorMessage) Then GoTo CleanExit
    If modProcessor.RunBatch("WHR3", 500, report) <> 1 Then GoTo CleanExit

    Set loInbox = wbInbox.Worksheets("InboxReceive").ListObjects("tblInboxReceive")
    Set loLog = wbInv.Worksheets("InventoryLog").ListObjects("tblInventoryLog")
    inboxRow = FindRowByColumnValue(loInbox, "SKU", "SKU-001")
    If inboxRow = 0 Then GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(loInbox, inboxRow, "Status")) <> "PROCESSED" Then GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(loLog, 2, "EventType")) <> EVENT_TYPE_RECEIVE Then GoTo CleanExit
    If CDbl(TestPhase2Helpers.GetRowValue(loLog, 2, "QtyDelta")) <> 7 Then GoTo CleanExit

    TestReceivingRoleFlow_QueuesAndProcessesEvent = 1

CleanExit:
    TestPhase2Helpers.CloseNoSave wbRole
    TestPhase2Helpers.CloseAndDeleteWorkbook wbInbox
    TestPhase2Helpers.CloseAndDeleteWorkbook wbInv
    TestPhase2Helpers.CloseNoSave wbAuth
    TestPhase2Helpers.CloseNoSave wbCfg
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestShippingRoleFlow_QueuesAndProcessesEvent() As Long
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim wbInv As Workbook
    Dim wbInbox As Workbook
    Dim tempRoot As String
    Dim report As String
    Dim errNotes As String
    Dim eventIdOut As String
    Dim currentUserId As String
    Dim loInbox As ListObject
    Dim loLog As ListObject
    Dim wbRole As Workbook
    Dim inboxRow As Long

    On Error GoTo CleanFail
    currentUserId = modRoleEventWriter.ResolveCurrentUserId()
    tempRoot = TestPhase2Helpers.BuildUniqueTestFolder("Phase3Ship")
    Set wbCfg = TestPhase2Helpers.BuildPhase2ConfigWorkbook("WHS3", "S1", "SHIP")
    TestPhase2Helpers.SetWarehouseConfigValue wbCfg, "PathDataRoot", tempRoot
    Set wbAuth = TestPhase2Helpers.BuildPhase2AuthWorkbook("WHS3")
    TestPhase2Helpers.AddCapability wbAuth, currentUserId, "SHIP_POST", "WHS3", "S1", "ACTIVE"
    TestPhase2Helpers.AddCapability wbAuth, "svc_processor", "INBOX_PROCESS", "WHS3", "*", "ACTIVE"
    Set wbInv = TestPhase2Helpers.BuildCanonicalInventoryWorkbook("WHS3", tempRoot, Array("SKU-001"))
    Set wbInbox = TestPhase2Helpers.BuildCanonicalShipInboxWorkbook("S1", tempRoot)

    Set wbRole = Application.Workbooks.Add
    SetupShippingRoleScaffold wbRole
    If Not modShippingEventCreator.QueueShipmentsSentEventFromWorkbook(wbRole, eventIdOut, errNotes) Then GoTo CleanExit
    If modProcessor.RunBatch("WHS3", 500, report) <> 1 Then GoTo CleanExit

    Set loInbox = wbInbox.Worksheets("InboxShip").ListObjects("tblInboxShip")
    Set loLog = wbInv.Worksheets("InventoryLog").ListObjects("tblInventoryLog")
    inboxRow = FindRowByColumnValue(loInbox, "EventID", eventIdOut)
    If inboxRow = 0 Then GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(loInbox, inboxRow, "Status")) <> "PROCESSED" Then GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(loLog, 2, "EventType")) <> EVENT_TYPE_SHIP Then GoTo CleanExit
    If CDbl(TestPhase2Helpers.GetRowValue(loLog, 2, "QtyDelta")) <> -5 Then GoTo CleanExit

    TestShippingRoleFlow_QueuesAndProcessesEvent = 1

CleanExit:
    TestPhase2Helpers.CloseNoSave wbRole
    TestPhase2Helpers.CloseAndDeleteWorkbook wbInbox
    TestPhase2Helpers.CloseAndDeleteWorkbook wbInv
    TestPhase2Helpers.CloseNoSave wbAuth
    TestPhase2Helpers.CloseNoSave wbCfg
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestProductionRoleFlow_QueuesAndProcessesEvent() As Long
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook
    Dim wbInv As Workbook
    Dim wbInbox As Workbook
    Dim tempRoot As String
    Dim report As String
    Dim errNotes As String
    Dim eventIdOut As String
    Dim currentUserId As String
    Dim loInbox As ListObject
    Dim loLog As ListObject
    Dim wbRole As Workbook
    Dim inboxRow As Long

    On Error GoTo CleanFail
    currentUserId = modRoleEventWriter.ResolveCurrentUserId()
    tempRoot = TestPhase2Helpers.BuildUniqueTestFolder("Phase3Prod")
    Set wbCfg = TestPhase2Helpers.BuildPhase2ConfigWorkbook("WHP3", "S1", "PROD")
    TestPhase2Helpers.SetWarehouseConfigValue wbCfg, "PathDataRoot", tempRoot
    Set wbAuth = TestPhase2Helpers.BuildPhase2AuthWorkbook("WHP3")
    TestPhase2Helpers.AddCapability wbAuth, currentUserId, "PROD_POST", "WHP3", "S1", "ACTIVE"
    TestPhase2Helpers.AddCapability wbAuth, "svc_processor", "INBOX_PROCESS", "WHP3", "*", "ACTIVE"
    Set wbInv = TestPhase2Helpers.BuildCanonicalInventoryWorkbook("WHP3", tempRoot, Array("SKU-FG"))
    Set wbInbox = TestPhase2Helpers.BuildCanonicalProductionInboxWorkbook("S1", tempRoot)

    Set wbRole = Application.Workbooks.Add
    SetupProductionRoleScaffold wbRole
    If Not modProductionEventCreator.QueueProductionCompleteEventFromWorkbook(wbRole, eventIdOut, errNotes) Then GoTo CleanExit
    If modProcessor.RunBatch("WHP3", 500, report) <> 1 Then GoTo CleanExit

    Set loInbox = wbInbox.Worksheets("InboxProd").ListObjects("tblInboxProd")
    Set loLog = wbInv.Worksheets("InventoryLog").ListObjects("tblInventoryLog")
    inboxRow = FindRowByColumnValue(loInbox, "EventID", eventIdOut)
    If inboxRow = 0 Then GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(loInbox, inboxRow, "Status")) <> "PROCESSED" Then GoTo CleanExit
    If CStr(TestPhase2Helpers.GetRowValue(loLog, 2, "EventType")) <> EVENT_TYPE_PROD_COMPLETE Then GoTo CleanExit
    If CDbl(TestPhase2Helpers.GetRowValue(loLog, 2, "QtyDelta")) <> 8 Then GoTo CleanExit

    TestProductionRoleFlow_QueuesAndProcessesEvent = 1

CleanExit:
    TestPhase2Helpers.CloseNoSave wbRole
    TestPhase2Helpers.CloseAndDeleteWorkbook wbInbox
    TestPhase2Helpers.CloseAndDeleteWorkbook wbInv
    TestPhase2Helpers.CloseNoSave wbAuth
    TestPhase2Helpers.CloseNoSave wbCfg
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Private Sub SetupReceivingRoleScaffold(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim headers As Variant
    Dim dataArr(1 To 1, 1 To 10) As Variant

    CleanupRoleSheets wb, Array("ReceivedTally")
    Set ws = wb.Worksheets(1)
    ws.Name = "ReceivedTally"

    headers = Array("REF_NUMBER", "ITEM_CODE", "VENDORS", "VENDOR_CODE", "DESCRIPTION", "ITEM", "UOM", "QUANTITY", "LOCATION", "ROW")
    dataArr(1, 1) = "REF-001"
    dataArr(1, 2) = "SKU-001"
    dataArr(1, 3) = "Vendor A"
    dataArr(1, 4) = "V001"
    dataArr(1, 5) = "Widget desc"
    dataArr(1, 6) = "Widget"
    dataArr(1, 7) = "EA"
    dataArr(1, 8) = 7
    dataArr(1, 9) = "A1"
    CreateTable ws, "A1", "AggregateReceived", headers, dataArr
End Sub

Private Sub SetupShippingRoleScaffold(ByVal wb As Workbook)
    Dim wsShip As Worksheet
    Dim wsInv As Worksheet
    Dim aggHeaders As Variant
    Dim invHeaders As Variant
    Dim aggData(1 To 1, 1 To 4) As Variant
    Dim invData(1 To 1, 1 To 4) As Variant

    CleanupRoleSheets wb, Array("ShipmentsTally", "InventoryManagement")

    Set wsShip = wb.Worksheets(1)
    wsShip.Name = "ShipmentsTally"
    aggHeaders = Array("QUANTITY", "UOM", "ITEM", "ROW")
    aggData(1, 1) = 5
    aggData(1, 2) = "EA"
    aggData(1, 3) = "Widget"
    aggData(1, 4) = 201
    CreateTable wsShip, "A1", "AggregatePackages", aggHeaders, aggData

    Set wsInv = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsInv.Name = "InventoryManagement"
    invHeaders = Array("ROW", "ITEM_CODE", "ITEM", "SHIPMENTS")
    invData(1, 1) = 201
    invData(1, 2) = "SKU-001"
    invData(1, 3) = "Widget"
    invData(1, 4) = 5
    CreateTable wsInv, "A1", "invSys", invHeaders, invData
End Sub

Private Sub SetupProductionRoleScaffold(ByVal wb As Workbook)
    Dim wsProd As Worksheet
    Dim wsInv As Worksheet
    Dim outHeaders As Variant
    Dim invHeaders As Variant
    Dim outData(1 To 1, 1 To 4) As Variant
    Dim invData(1 To 1, 1 To 3) As Variant

    CleanupRoleSheets wb, Array("Production", "InventoryManagement")

    Set wsProd = wb.Worksheets(1)
    wsProd.Name = "Production"
    outHeaders = Array("PROCESS", "OUTPUT", "REAL OUTPUT", "ROW")
    outData(1, 1) = "Mix"
    outData(1, 2) = "Finished Good"
    outData(1, 3) = 8
    outData(1, 4) = 301
    CreateTable wsProd, "A1", "ProductionOutput", outHeaders, outData

    Set wsInv = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsInv.Name = "InventoryManagement"
    invHeaders = Array("ROW", "ITEM_CODE", "ITEM")
    invData(1, 1) = 301
    invData(1, 2) = "SKU-FG"
    invData(1, 3) = "Finished Good"
    CreateTable wsInv, "A1", "invSys", invHeaders, invData
End Sub

Private Sub CreateTable(ByVal ws As Worksheet, ByVal topLeft As String, ByVal tableName As String, ByVal headers As Variant, Optional ByVal dataRows As Variant)
    Dim startCell As Range
    Dim rowCount As Long
    Dim colCount As Long
    Dim target As Range
    Dim lo As ListObject

    Set startCell = ws.Range(topLeft)
    colCount = UBound(headers) - LBound(headers) + 1
    startCell.Resize(1, colCount).Value = headers

    rowCount = 1
    If Not IsMissing(dataRows) Then
        rowCount = UBound(dataRows, 1) + 1
        startCell.Offset(1, 0).Resize(UBound(dataRows, 1), UBound(dataRows, 2)).Value = dataRows
    End If

    Set target = startCell.Resize(rowCount, colCount)
    Set lo = ws.ListObjects.Add(xlSrcRange, target, , xlYes)
    lo.Name = tableName
End Sub

Private Sub CleanupRoleSheets(ByVal wb As Workbook, ByVal sheetNames As Variant)
    Dim i As Long
    Dim ws As Worksheet

    Application.DisplayAlerts = False
    On Error Resume Next
    For i = LBound(sheetNames) To UBound(sheetNames)
        Set ws = Nothing
        Set ws = wb.Worksheets(CStr(sheetNames(i)))
        If Not ws Is Nothing Then ws.Delete
    Next i
    On Error GoTo 0
    Application.DisplayAlerts = True
End Sub

Private Sub Tally(ByVal resultIn As Long, ByRef passed As Long, ByRef failed As Long)
    If resultIn = 1 Then
        passed = passed + 1
    Else
        failed = failed + 1
    End If
End Sub

Private Function FindRowByColumnValue(ByVal lo As ListObject, ByVal columnName As String, ByVal expectedValue As String) As Long
    Dim i As Long

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To lo.ListRows.Count
        If StrComp(CStr(TestPhase2Helpers.GetRowValue(lo, i, columnName)), expectedValue, vbTextCompare) = 0 Then
            FindRowByColumnValue = i
            Exit Function
        End If
    Next i
End Function
