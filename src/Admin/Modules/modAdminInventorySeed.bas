Attribute VB_Name = "modAdminInventorySeed"
Option Explicit

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
    AddDemoInventoryItem rows, "DEMO-RAW-BLACK-TEA", "Black Tea", 5000, "lbs", _
        "Loose black tea for brewing.", "raw"
    AddDemoInventoryItem rows, "DEMO-RAW-FILTERED-WATER", "Filtered Water", 20000, "lbs", _
        "Filtered brewing water.", "raw"
    AddDemoInventoryItem rows, "DEMO-RAW-CARDAMOM", "Cardamom (Decorticated)", 500, "lbs", _
        "Cardamom for chai blend.", "raw"
    AddDemoInventoryItem rows, "DEMO-RAW-BLACK-PEPPER", "Black Pepper (Whole)", 300, "lbs", _
        "Black pepper for chai blend.", "raw"
    AddDemoInventoryItem rows, "DEMO-RAW-NUTMEG", "Nutmeg (Ground)", 250, "lbs", _
        "Ground nutmeg for chai blend.", "raw"
    AddDemoInventoryItem rows, "DEMO-RAW-GINGER", "Ginger (Ground)", 250, "lbs", _
        "Ground ginger for chai blend.", "raw"
    AddDemoInventoryItem rows, "DEMO-RAW-CITRIC-ACID", "Citric Acid", 120, "lbs", _
        "Citric acid ingredient.", "raw"
    AddDemoInventoryItem rows, "DEMO-RAW-CASSIA-OIL", "Cassia Oil 340139", 80, "lbs", _
        "Cassia oil for chai blend.", "raw"
    AddDemoInventoryItem rows, "DEMO-RAW-LEMON-OIL", "Lemon Oil (5x) 34013", 80, "lbs", _
        "Lemon oil for chai blend.", "raw"
    AddDemoInventoryItem rows, "DEMO-RAW-ORANGE-OIL", "Orange Oil (Cold Press)", 80, "lbs", _
        "Orange oil for chai blend.", "raw"
    AddDemoInventoryItem rows, "DEMO-RAW-SUGAR-WHITE", "Pure Cane Sugar White Granulated", 8000, "lbs", _
        "White granulated cane sugar.", "raw"
    AddDemoInventoryItem rows, "DEMO-RAW-SUGAR-CLOUDY", "Pure Cane Sugar Cloudy White Granulated", 6000, "lbs", _
        "Cloudy white granulated cane sugar.", "raw"
    AddDemoInventoryItem rows, "DEMO-WIP-BREWED-BLACK-TEA", "Brewed Black Tea", 1200, "lbs", _
        "Intermediate brewed tea.", "wip"
    AddDemoInventoryItem rows, "DEMO-WIP-CHAI-SPICE-BLEND", "Classic Chai Spice Blend", 600, "lbs", _
        "Intermediate chai spice blend.", "wip"
    AddDemoInventoryItem rows, "DEMO-RAW-BROWN-COLOR", "Brown Color 10.5g", 100, "lbs", _
        "Brown color ingredient.", "raw"
    AddDemoInventoryItem rows, "DEMO-FG-CLASSIC-CHAI", "Black Scottie Chai Classic Concentrate", 400, "gal", _
        "Finished good concentrate for shipping.", "shippable"
    AddDemoInventoryItem rows, "DEMO-FG-12PACK-CASE", "Classic Chai 12-Pack Case", 120, "each", _
        "Finished case pack.", "shippable"
    AddDemoInventoryItem rows, "DEMO-FG-SAMPLE-BOX", "Black Scottie Sample Box", 80, "each", _
        "Sample assortment box.", "shippable"
    AddDemoInventoryItem rows, "DEMO-PKG-TIN", "Classic Chai Tin", 1000, "each", _
        "Packaging component for receiving and box-building tests.", "packaging.ship"
    AddDemoInventoryItem rows, "DEMO-PKG-SHIPPING-CARTON", "Shipping Carton Blank", 1000, "each", _
        "Corrugated carton blank for Box Designer and Box Maker tests.", "packaging.ship"
    AddDemoInventoryItem rows, "DEMO-PKG-CASE-DIVIDER", "Case Divider", 2000, "each", _
        "Divider insert for shipping-box alternatives.", "packaging.ship"
    AddDemoInventoryItem rows, "DEMO-PKG-SHIPPING-LABEL", "Shipping Label", 5000, "each", _
        "Adhesive label consumed when making a shipping box.", "packaging.ship"
    AddDemoInventoryItem rows, "DEMO-PKG-PACKING-TAPE", "Packing Tape Strip", 5000, "each", _
        "Premeasured packing-tape strip for shipping assembly tests.", "packaging.ship"
    AddDemoInventoryItem rows, "DEMO-PKG-VOID-FILL", "Void Fill Sheet", 3000, "each", _
        "Protective void fill for shipping-box alternatives.", "packaging.ship"
    Set BuildDemoInventoryPayload = rows
End Function

Private Sub AddDemoInventoryItem(ByVal rows As Collection, _
                                 ByVal itemCode As String, _
                                 ByVal itemName As String, _
                                 ByVal qty As Double, _
                                 ByVal uom As String, _
                                 ByVal description As String, _
                                 ByVal category As String)
    Dim item As Object

    Set item = modRoleEventWriter.CreateInventoryEntityPayloadItem( _
        modRoleEventWriter.CreateSystemKey(), itemCode, qty, "CLEARVIEW", _
        "GOOD", "", "Admin R1 workflow demo seed")
    item("ITEM_CODE") = itemCode
    item("ITEM") = itemName
    item("UOM") = uom
    item("DESCRIPTION") = description
    item("CATEGORY") = category
    rows.Add item
End Sub
