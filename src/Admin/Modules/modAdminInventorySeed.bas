Attribute VB_Name = "modAdminInventorySeed"
Option Explicit

Private Const DEMO_INVENTORY_QTY As Double = 1000#

Public Function SeedDemoInventoryForWarehouse(ByVal warehouseId As String, _
                                              ByVal stationId As String, _
                                              ByVal userId As String, _
                                              ByRef report As String) As Boolean
    Dim payloadItems As Collection
    Dim payloadJson As String
    Dim eventIdOut As String
    Dim queueError As String
    Dim batchReport As String
    Dim retryReport As String
    Dim processedCount As Long
    Dim inboxReport As String
    Dim productionInboxPath As String
    Dim productionInbox As Workbook
    Dim productionInboxWasOpen As Boolean

    On Error GoTo FailSeed

    If Not EnsureDemoStationInboxes(warehouseId, stationId, inboxReport) Then
        report = inboxReport
        Exit Function
    End If

    Set payloadItems = BuildDemoInventoryPayload()
    payloadJson = modRoleEventWriter.BuildPayloadJsonFromCollection(payloadItems)
    If payloadJson = "" Or payloadJson = "[]" Then
        report = "Demo inventory payload was empty."
        Exit Function
    End If

    productionInboxPath = modRoleEventWriter.ResolveInboxWorkbookPath( _
        "INVENTORY_CREATE", warehouseId, stationId, queueError)
    If productionInboxPath = "" Then
        report = "Production inbox path could not be resolved: " & queueError
        Exit Function
    End If
    Set productionInbox = FindOpenWorkbookByPathSeed(productionInboxPath)
    productionInboxWasOpen = Not productionInbox Is Nothing
    If productionInbox Is Nothing Then
        Set productionInbox = modRoleEventWriter.OpenInboxWorkbook( _
            "INVENTORY_CREATE", warehouseId, stationId, queueError)
    End If
    If productionInbox Is Nothing Then
        report = "Production inbox could not be opened: " & queueError
        Exit Function
    End If

    If Not modRoleEventWriter.QueueInventoryCreateEvent(warehouseId, stationId, userId, payloadJson, _
                                                        "Admin demo inventory seed", _
                                                        0, productionInbox, eventIdOut, queueError, "") Then
        report = "Seed event could not be queued: " & queueError
        GoTo CleanExit
    End If

    processedCount = modProcessor.RunBatch(warehouseId, 0, batchReport)
    If processedCount < 1 And InStr(1, batchReport, "Poison=0", vbTextCompare) > 0 Then
        processedCount = modProcessor.RunBatch(warehouseId, 0, retryReport)
        batchReport = batchReport & "|Retry=" & retryReport
    End If
    If processedCount < 1 Then
        report = "Seed event was queued but not applied. " & batchReport
        GoTo CleanExit
    End If

    report = "Demo inventory seeded.|Applied=" & CStr(processedCount) & "|Processor=" & batchReport
    SeedDemoInventoryForWarehouse = True
CleanExit:
    On Error Resume Next
    If Not productionInboxWasOpen Then productionInbox.Close SaveChanges:=True
    On Error GoTo 0
    Exit Function

FailSeed:
    report = "SeedDemoInventoryForWarehouse failed: " & Err.Description
    Resume CleanExit
End Function

Private Function FindOpenWorkbookByPathSeed(ByVal fullPath As String) As Workbook
    Dim wb As Workbook

    fullPath = Trim$(fullPath)
    If fullPath = "" Then Exit Function
    For Each wb In Application.Workbooks
        If StrComp(Trim$(wb.FullName), fullPath, vbTextCompare) = 0 Then
            Set FindOpenWorkbookByPathSeed = wb
            Exit Function
        End If
    Next wb
End Function

Private Function EnsureDemoStationInboxes(ByVal warehouseId As String, _
                                          ByVal stationId As String, _
                                          ByRef report As String) As Boolean
    Dim inboxPath As String
    Dim stepReport As String

    If Not modConfig.EnsureStationInbox(warehouseId, stationId, "RECEIVE", "", inboxPath, stepReport) Then
        report = stepReport
        Exit Function
    End If
    inboxPath = ""
    If Not modConfig.EnsureStationInbox(warehouseId, stationId, "SHIP", "", inboxPath, stepReport) Then
        report = stepReport
        Exit Function
    End If
    inboxPath = ""
    If Not modConfig.EnsureStationInbox(warehouseId, stationId, "PRODUCTION", "", inboxPath, stepReport) Then
        report = stepReport
        Exit Function
    End If

    EnsureDemoStationInboxes = True
End Function

Private Function BuildDemoInventoryPayload() As Collection
    Dim rows As Collection

    Set rows = New Collection
    rows.Add modRoleEventWriter.CreateInventoryEntityPayloadItem( _
        modRoleEventWriter.CreateSystemKey(), "DEMO-RAW-BLACK-TEA", _
        DEMO_INVENTORY_QTY, "NAS-A1", "GOOD", "", "Admin demo seed")
    rows.Add modRoleEventWriter.CreateInventoryEntityPayloadItem( _
        modRoleEventWriter.CreateSystemKey(), "DEMO-SPICE-CARDAMOM", _
        DEMO_INVENTORY_QTY, "NAS-A2", "GOOD", "", "Admin demo seed")
    rows.Add modRoleEventWriter.CreateInventoryEntityPayloadItem( _
        modRoleEventWriter.CreateSystemKey(), "DEMO-PKG-TIN", _
        DEMO_INVENTORY_QTY, "NAS-P1", "GOOD", "", "Admin demo seed")
    Set BuildDemoInventoryPayload = rows
End Function
