Attribute VB_Name = "TestDesignsDomain"
Option Explicit

Public Function TestDesignsSchema_CreatesAndValidatesAuthoritativeTables() As Long
    Dim wb As Workbook
    Dim report As String

    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail
    If Not modDesignsSchema.EnsureDesignsSchema(wb, report) Then GoTo CleanExit
    If modDesignsSchema.ValidateDesignsSchema(wb) <> "" Then GoTo CleanExit
    If FindDesignsTestTable(wb, "tblDesigns") Is Nothing Then GoTo CleanExit
    If FindDesignsTestTable(wb, "tblDesignLines") Is Nothing Then GoTo CleanExit
    If FindDesignsTestTable(wb, "tblDesignEvents") Is Nothing Then GoTo CleanExit
    If FindDesignsTestTable(wb, "tblAppliedDesignEvents") Is Nothing Then GoTo CleanExit
    If FindDesignsTestTable(wb, "tblLocks") Is Nothing Then GoTo CleanExit
    TestDesignsSchema_CreatesAndValidatesAuthoritativeTables = 1

CleanExit:
    CloseDesignsTestWorkbook wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestDesignsSchema_IsIdempotent() As Long
    Dim wb As Workbook
    Dim report As String
    Dim beforeCount As Long
    Dim afterCount As Long

    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail
    If Not modDesignsSchema.EnsureDesignsSchema(wb, report) Then GoTo CleanExit
    beforeCount = CountDesignsTestTables(wb)
    If Not modDesignsSchema.EnsureDesignsSchema(wb, report) Then GoTo CleanExit
    afterCount = CountDesignsTestTables(wb)
    If beforeCount = 5 And afterCount = beforeCount Then TestDesignsSchema_IsIdempotent = 1

CleanExit:
    CloseDesignsTestWorkbook wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestDesignsQueries_ListDesignsAndGetBOMAreReadOnly() As Long
    Dim wb As Workbook
    Dim report As String
    Dim loDesigns As ListObject
    Dim loLines As ListObject
    Dim lr As ListRow
    Dim designs As Variant
    Dim bom As Variant
    Dim designsRowsBefore As Long
    Dim linesRowsBefore As Long

    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail
    If Not modDesignsSchema.EnsureDesignsSchema(wb, report) Then GoTo CleanExit
    Set loDesigns = FindDesignsTestTable(wb, "tblDesigns")
    Set loLines = FindDesignsTestTable(wb, "tblDesignLines")

    Set lr = loDesigns.ListRows.Add
    SetDesignsTestValue loDesigns, lr.Index, "DesignId", "TEA-BLACK"
    SetDesignsTestValue loDesigns, lr.Index, "DesignVersion", "1"
    SetDesignsTestValue loDesigns, lr.Index, "DesignType", "RECIPE"
    SetDesignsTestValue loDesigns, lr.Index, "DesignName", "Brewed Black Tea"
    SetDesignsTestValue loDesigns, lr.Index, "Status", "RELEASED"

    Set lr = loLines.ListRows.Add
    SetDesignsTestValue loLines, lr.Index, "DesignId", "TEA-BLACK"
    SetDesignsTestValue loLines, lr.Index, "DesignVersion", "1"
    SetDesignsTestValue loLines, lr.Index, "LineNo", 1
    SetDesignsTestValue loLines, lr.Index, "IOType", "USED"
    SetDesignsTestValue loLines, lr.Index, "ComponentSKU", "SKU-BLACK-TEA"
    SetDesignsTestValue loLines, lr.Index, "Qty", 32.5
    SetDesignsTestValue loLines, lr.Index, "UOM", "LB"

    designsRowsBefore = loDesigns.ListRows.Count
    linesRowsBefore = loLines.ListRows.Count
    designs = modDesignsQueries.ListDesigns(wb, "RELEASED")
    bom = modDesignsQueries.GetBOM("TEA-BLACK", "1", wb)
    If IsEmpty(designs) Or Not IsArray(designs) Then GoTo CleanExit
    If IsEmpty(bom) Or Not IsArray(bom) Then GoTo CleanExit
    If CStr(designs(1, 1)) <> "TEA-BLACK" Then GoTo CleanExit
    If CStr(bom(1, 4)) <> "SKU-BLACK-TEA" Then GoTo CleanExit
    If loDesigns.ListRows.Count <> designsRowsBefore Then GoTo CleanExit
    If loLines.ListRows.Count <> linesRowsBefore Then GoTo CleanExit
    TestDesignsQueries_ListDesignsAndGetBOMAreReadOnly = 1

CleanExit:
    CloseDesignsTestWorkbook wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestDesignsDomain_DiagnosticDeclaresNoStartupMutation() As Long
    Dim diagnostic As String
    diagnostic = modDesignsInit.DiagnoseDesignsDomain()
    If InStr(1, diagnostic, "StartupMutation=False", vbTextCompare) > 0 _
       And InStr(1, diagnostic, "WHx.invSys.Data.Designs.xlsb", vbTextCompare) > 0 Then
        TestDesignsDomain_DiagnosticDeclaresNoStartupMutation = 1
    End If
End Function

Public Function TestDesignInboxSchema_CarriesDesignIdentity() As Long
    Dim wb As Workbook
    Dim lo As ListObject
    Dim report As String

    On Error GoTo CleanExit
    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modProcessor.EnsureProductionInboxSchema(wb, report) Then GoTo CleanExit
    Set lo = FindDesignsTestTable(wb, "tblInboxProd")
    If lo Is Nothing Then GoTo CleanExit
    If Not DesignsTestColumnExists(lo, "DesignId") Then GoTo CleanExit
    If Not DesignsTestColumnExists(lo, "DesignVersion") Then GoTo CleanExit
    TestDesignInboxSchema_CarriesDesignIdentity = 1

CleanExit:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
End Function

Public Function TestDesignsApply_LifecycleIsIdempotentAndRebuildable() As Long
    Dim wb As Workbook
    Dim report As String
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim createEvent As Object
    Dim releaseEvent As Object
    Dim loDesigns As ListObject
    Dim loLines As ListObject
    Dim loEvents As ListObject

    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail
    If Not modDesignsSchema.EnsureDesignsSchema(wb, report) Then GoTo CleanExit
    Set createEvent = BuildDesignsTestEvent("DES-EVT-1", "DESIGN_CREATE", "TEA-BLACK", "1", _
        "[{""DesignType"":""RECIPE"",""DesignName"":""Brewed Black Tea"",""Description"":""Concentrate""," & _
        """LineNo"":1,""Process"":1,""IOType"":""USED"",""ComponentSKU"":""SKU-BLACK-TEA""," & _
        """Qty"":32.5,""UOM"":""LB"",""Percent"":100}]")
    If Not modDesignsApply.ApplyDesignEvent(createEvent, wb, "RUN-DES-1", statusOut, errorCode, errorMessage) Then GoTo CleanExit
    If statusOut <> "APPLIED" Then GoTo CleanExit

    statusOut = "": errorCode = "": errorMessage = ""
    If Not modDesignsApply.ApplyDesignEvent(createEvent, wb, "RUN-DES-2", statusOut, errorCode, errorMessage) Then GoTo CleanExit
    If statusOut <> "SKIP_DUP" Then GoTo CleanExit

    Set releaseEvent = BuildDesignsTestEvent("DES-EVT-2", "DESIGN_RELEASE", "TEA-BLACK", "1", "")
    statusOut = "": errorCode = "": errorMessage = ""
    If Not modDesignsApply.ApplyDesignEvent(releaseEvent, wb, "RUN-DES-3", statusOut, errorCode, errorMessage) Then GoTo CleanExit

    Set loDesigns = FindDesignsTestTable(wb, "tblDesigns")
    Set loLines = FindDesignsTestTable(wb, "tblDesignLines")
    Set loEvents = FindDesignsTestTable(wb, "tblDesignEvents")
    If loEvents.ListRows.Count <> 2 Then GoTo CleanExit
    If loDesigns.ListRows.Count <> 1 Or loLines.ListRows.Count <> 1 Then GoTo CleanExit
    If CStr(loDesigns.DataBodyRange.Cells(1, loDesigns.ListColumns("Status").Index).Value) <> "RELEASED" Then GoTo CleanExit

    Do While loDesigns.ListRows.Count > 0: loDesigns.ListRows(1).Delete: Loop
    Do While loLines.ListRows.Count > 0: loLines.ListRows(1).Delete: Loop
    If Not modDesignsApply.RebuildDesignProjections(wb, report) Then GoTo CleanExit
    If loDesigns.ListRows.Count <> 1 Or loLines.ListRows.Count <> 1 Then GoTo CleanExit
    If CStr(loDesigns.DataBodyRange.Cells(1, loDesigns.ListColumns("Status").Index).Value) <> "RELEASED" Then GoTo CleanExit
    TestDesignsApply_LifecycleIsIdempotentAndRebuildable = 1

CleanExit:
    CloseDesignsTestWorkbook wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestDesignsApply_RejectsDuplicateImmutableVersion() As Long
    Dim wb As Workbook
    Dim report As String
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim evt As Object

    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail
    If Not modDesignsSchema.EnsureDesignsSchema(wb, report) Then GoTo CleanExit
    Set evt = BuildDesignsTestEvent("DES-EVT-10", "DESIGN_CREATE", "TEA-GREEN", "1", _
        "[{""DesignType"":""RECIPE"",""DesignName"":""Green Tea"",""LineNo"":1," & _
        """IOType"":""OUTPUT"",""ComponentSKU"":""SKU-GREEN-TEA"",""Qty"":1,""UOM"":""LB""}]")
    If Not modDesignsApply.ApplyDesignEvent(evt, wb, "RUN-DES-10", statusOut, errorCode, errorMessage) Then GoTo CleanExit
    Set evt = BuildDesignsTestEvent("DES-EVT-11", "DESIGN_CREATE", "TEA-GREEN", "1", _
        "[{""DesignType"":""RECIPE"",""DesignName"":""Changed Green Tea""}]")
    statusOut = "": errorCode = "": errorMessage = ""
    If modDesignsApply.ApplyDesignEvent(evt, wb, "RUN-DES-11", statusOut, errorCode, errorMessage) Then GoTo CleanExit
    If errorCode <> "DESIGN_VERSION_EXISTS" Then GoTo CleanExit
    TestDesignsApply_RejectsDuplicateImmutableVersion = 1

CleanExit:
    CloseDesignsTestWorkbook wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestDesignsApply_RebuildUsesAppliedSeqNotTableOrder() As Long
    Dim wb As Workbook
    Dim report As String
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim evt As Object
    Dim loEvents As ListObject
    Dim loDesigns As ListObject
    Dim rowOne As Variant

    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail
    If Not modDesignsSchema.EnsureDesignsSchema(wb, report) Then GoTo CleanExit
    Set evt = BuildDesignsTestEvent("DES-EVT-20", "DESIGN_CREATE", "TEA-ORDERED", "1", _
        "[{""DesignType"":""RECIPE"",""DesignName"":""Ordered Tea""}]")
    If Not modDesignsApply.ApplyDesignEvent(evt, wb, "RUN-DES-20", statusOut, errorCode, errorMessage) Then GoTo CleanExit
    Set evt = BuildDesignsTestEvent("DES-EVT-21", "DESIGN_RELEASE", "TEA-ORDERED", "1", "")
    statusOut = "": errorCode = "": errorMessage = ""
    If Not modDesignsApply.ApplyDesignEvent(evt, wb, "RUN-DES-21", statusOut, errorCode, errorMessage) Then GoTo CleanExit

    Set loEvents = FindDesignsTestTable(wb, "tblDesignEvents")
    Set loDesigns = FindDesignsTestTable(wb, "tblDesigns")
    If loEvents Is Nothing Or loDesigns Is Nothing Then GoTo CleanExit
    If loEvents.ListRows.Count <> 2 Then GoTo CleanExit
    rowOne = loEvents.DataBodyRange.Rows(1).Value
    loEvents.DataBodyRange.Rows(1).Value = loEvents.DataBodyRange.Rows(2).Value
    loEvents.DataBodyRange.Rows(2).Value = rowOne
    Do While loDesigns.ListRows.Count > 0: loDesigns.ListRows(1).Delete: Loop

    If Not modDesignsApply.RebuildDesignProjections(wb, report) Then GoTo CleanExit
    If loDesigns.ListRows.Count <> 1 Then GoTo CleanExit
    If CStr(loDesigns.DataBodyRange.Cells(1, loDesigns.ListColumns("Status").Index).Value) <> "RELEASED" Then GoTo CleanExit
    TestDesignsApply_RebuildUsesAppliedSeqNotTableOrder = 1

CleanExit:
    CloseDesignsTestWorkbook wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestDesignsApply_ImmutableVersionSurvivesProjectionLoss() As Long
    Dim wb As Workbook
    Dim report As String
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim evt As Object
    Dim loEvents As ListObject
    Dim loDesigns As ListObject
    Dim loLines As ListObject

    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail
    If Not modDesignsSchema.EnsureDesignsSchema(wb, report) Then GoTo CleanExit
    Set evt = BuildDesignsTestEvent("DES-EVT-30", "DESIGN_CREATE", "TEA-HISTORY", "1", _
        "[{""DesignType"":""RECIPE"",""DesignName"":""History Tea""}]")
    If Not modDesignsApply.ApplyDesignEvent(evt, wb, "RUN-DES-30", statusOut, errorCode, errorMessage) Then GoTo CleanExit

    Set loEvents = FindDesignsTestTable(wb, "tblDesignEvents")
    Set loDesigns = FindDesignsTestTable(wb, "tblDesigns")
    Set loLines = FindDesignsTestTable(wb, "tblDesignLines")
    Do While loDesigns.ListRows.Count > 0: loDesigns.ListRows(1).Delete: Loop
    Do While loLines.ListRows.Count > 0: loLines.ListRows(1).Delete: Loop

    Set evt = BuildDesignsTestEvent("DES-EVT-31", "DESIGN_CREATE", "TEA-HISTORY", "1", _
        "[{""DesignType"":""RECIPE"",""DesignName"":""Illegal Rewrite""}]")
    statusOut = "": errorCode = "": errorMessage = ""
    If modDesignsApply.ApplyDesignEvent(evt, wb, "RUN-DES-31", statusOut, errorCode, errorMessage) Then GoTo CleanExit
    If errorCode <> "DESIGN_VERSION_EXISTS" Then GoTo CleanExit
    If loEvents.ListRows.Count <> 1 Then GoTo CleanExit
    If loDesigns.ListRows.Count <> 1 Then GoTo CleanExit
    TestDesignsApply_ImmutableVersionSurvivesProjectionLoss = 1

CleanExit:
    CloseDesignsTestWorkbook wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestInventorySchema_RejectsMissingOrAddinAuthority() As Long
    Dim report As String

    If modInventorySchema.EnsureInventorySchema(Nothing, report) Then Exit Function
    If InStr(1, report, "Authoritative Inventory workbook was not supplied", vbTextCompare) = 0 Then Exit Function
    TestInventorySchema_RejectsMissingOrAddinAuthority = 1
End Function

Public Function TestInventoryDomain_DiagnosticDisablesLegacyDirectWrites() As Long
    Dim diagnostic As String

    diagnostic = modInventoryInit.DiagnoseInventoryDomain()
    If InStr(1, diagnostic, "LegacyDirectWrites=False", vbTextCompare) = 0 Then Exit Function
    If InStr(1, diagnostic, "UndoModel=CompensatingEvent", vbTextCompare) = 0 Then Exit Function
    TestInventoryDomain_DiagnosticDisablesLegacyDirectWrites = 1
End Function

Public Function TestInventoryDomain_LegacyLogDeletionBridgeIsNoOp() As Long
    Dim removedRows As Collection

    Set removedRows = modInventoryBridgeApi.RemoveLastBulkLogEntriesBridgeResult(5)
    If removedRows Is Nothing Then Exit Function
    If removedRows.Count <> 0 Then Exit Function
    TestInventoryDomain_LegacyLogDeletionBridgeIsNoOp = 1
End Function

Public Function TestInventoryQueries_ReadRebuiltEventLogProjection() As Long
    Dim wb As Workbook
    Dim report As String
    Dim loLog As ListObject
    Dim lr As ListRow
    Dim locations As Variant

    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail
    If Not modInventorySchema.EnsureInventorySchema(wb, report) Then GoTo CleanExit
    Set loLog = FindDesignsTestTable(wb, "tblInventoryLog")
    loLog.Parent.Unprotect

    Set lr = loLog.ListRows.Add
    SetDesignsTestValue loLog, lr.Index, "EventID", "INV-Q-1"
    SetDesignsTestValue loLog, lr.Index, "AppliedSeq", 1
    SetDesignsTestValue loLog, lr.Index, "AppliedAtUTC", Now
    SetDesignsTestValue loLog, lr.Index, "SKU", "SKU-QUERY-1"
    SetDesignsTestValue loLog, lr.Index, "QtyDelta", 10
    SetDesignsTestValue loLog, lr.Index, "Location", "CLEARVIEW"

    Set lr = loLog.ListRows.Add
    SetDesignsTestValue loLog, lr.Index, "EventID", "INV-Q-2"
    SetDesignsTestValue loLog, lr.Index, "AppliedSeq", 2
    SetDesignsTestValue loLog, lr.Index, "AppliedAtUTC", Now
    SetDesignsTestValue loLog, lr.Index, "SKU", "SKU-QUERY-1"
    SetDesignsTestValue loLog, lr.Index, "QtyDelta", -3
    SetDesignsTestValue loLog, lr.Index, "Location", "CLEARVIEW"

    If Not modInventoryApply.RebuildInventoryProjectionsForWorkbook(wb, report) Then GoTo CleanExit
    If Abs(modInventoryQueries.GetOnHandQty("SKU-QUERY-1", wb) - 7) > 0.0001 Then GoTo CleanExit
    locations = modInventoryQueries.GetLocationBalances("SKU-QUERY-1", wb)
    If IsEmpty(locations) Or Not IsArray(locations) Then GoTo CleanExit
    If CStr(locations(1, 1)) <> "CLEARVIEW" Then GoTo CleanExit
    If Abs(CDbl(locations(1, 2)) - 7) > 0.0001 Then GoTo CleanExit
    TestInventoryQueries_ReadRebuiltEventLogProjection = 1

CleanExit:
    CloseDesignsTestWorkbook wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestInventoryQueries_PickerPublishesEverySkuLocation() As Long
    Dim wb As Workbook
    Dim report As String
    Dim loCatalog As ListObject
    Dim loBalance As ListObject
    Dim loLocation As ListObject
    Dim lr As ListRow
    Dim items As Variant
    Dim r As Long
    Dim foundA1 As Boolean
    Dim foundClearview As Boolean

    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail
    If Not modInventorySchema.EnsureInventorySchema(wb, report) Then GoTo CleanExit
    Set loCatalog = FindDesignsTestTable(wb, "tblSkuCatalog")
    Set loBalance = FindDesignsTestTable(wb, "tblSkuBalance")
    Set loLocation = FindDesignsTestTable(wb, "tblLocationBalance")
    If loCatalog Is Nothing Or loBalance Is Nothing Or loLocation Is Nothing Then GoTo CleanExit
    loCatalog.Parent.Unprotect
    loBalance.Parent.Unprotect
    loLocation.Parent.Unprotect

    Set lr = loCatalog.ListRows.Add
    SetDesignsTestValue loCatalog, lr.Index, "SKU", "SKU-PICKER-1"
    SetDesignsTestValue loCatalog, lr.Index, "ROW", 96
    SetDesignsTestValue loCatalog, lr.Index, "ITEM_CODE", "SKU-PICKER-1"
    SetDesignsTestValue loCatalog, lr.Index, "ITEM", "Malawi Fine Cut Black Tea"
    SetDesignsTestValue loCatalog, lr.Index, "UOM", "LB"
    SetDesignsTestValue loCatalog, lr.Index, "LOCATION", "A1"

    Set lr = loBalance.ListRows.Add
    SetDesignsTestValue loBalance, lr.Index, "SKU", "SKU-PICKER-1"
    SetDesignsTestValue loBalance, lr.Index, "QtyOnHand", 3175

    Set lr = loLocation.ListRows.Add
    SetDesignsTestValue loLocation, lr.Index, "SKU", "SKU-PICKER-1"
    SetDesignsTestValue loLocation, lr.Index, "Location", "A1"
    SetDesignsTestValue loLocation, lr.Index, "QtyOnHand", 1000

    Set lr = loLocation.ListRows.Add
    SetDesignsTestValue loLocation, lr.Index, "SKU", "SKU-PICKER-1"
    SetDesignsTestValue loLocation, lr.Index, "Location", "CLEARVIEW"
    SetDesignsTestValue loLocation, lr.Index, "QtyOnHand", 2175

    items = modInventoryBridgeApi.ListInventoryPickerItemsBridgeResult("Malawi", wb)
    If IsEmpty(items) Or Not IsArray(items) Then GoTo CleanExit
    For r = LBound(items, 1) To UBound(items, 1)
        If CStr(items(r, 1)) = "96" And CDbl(items(r, 4)) = 3175 Then
            If StrComp(CStr(items(r, 5)), "A1", vbTextCompare) = 0 Then foundA1 = True
            If StrComp(CStr(items(r, 5)), "CLEARVIEW", vbTextCompare) = 0 Then foundClearview = True
        End If
    Next r
    If foundA1 And foundClearview Then TestInventoryQueries_PickerPublishesEverySkuLocation = 1

CleanExit:
    CloseDesignsTestWorkbook wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestInventoryApply_ShipRejectsNegativeInventory() As Long
    Dim wb As Workbook
    Dim seedEvent As Object
    Dim shipEvent As Object
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim payloadJson As String
    Dim loLog As ListObject

    Set wb = TestPhase2Helpers.BuildPhase2InventoryWorkbook("WH-NEG-SHIP", Array("SKU-NEG-SHIP"), False)
    Set seedEvent = TestPhase2Helpers.CreateReceiveEvent( _
        "INV-NEG-SHIP-SEED", "WH-NEG-SHIP", "S-NEG", "U-NEG", "SKU-NEG-SHIP", 5, "CLEARVIEW", "seed")
    payloadJson = TestPhase2Helpers.BuildPayloadJson( _
        TestPhase2Helpers.CreatePayloadItem(1, "SKU-NEG-SHIP", 6, "CLEARVIEW", "overdraw"))
    Set shipEvent = TestPhase2Helpers.CreatePayloadEvent( _
        "INV-NEG-SHIP-APPLY", EVENT_TYPE_SHIP, "WH-NEG-SHIP", "S-NEG", "U-NEG", payloadJson)

    On Error GoTo CleanFail
    If Not modInventoryApply.ApplyEvent(seedEvent, wb, "RUN-NEG-SHIP-SEED", statusOut, errorCode, errorMessage) Then GoTo CleanExit
    statusOut = "": errorCode = "": errorMessage = ""
    If modInventoryApply.ApplyEvent(shipEvent, wb, "RUN-NEG-SHIP", statusOut, errorCode, errorMessage) Then GoTo CleanExit
    If errorCode <> "INSUFFICIENT_INVENTORY" Then GoTo CleanExit
    If InStr(1, errorMessage, EVENT_TYPE_SHIP, vbTextCompare) = 0 Then GoTo CleanExit
    Set loLog = FindDesignsTestTable(wb, "tblInventoryLog")
    If loLog Is Nothing Or loLog.ListRows.Count <> 1 Then GoTo CleanExit
    TestInventoryApply_ShipRejectsNegativeInventory = 1

CleanExit:
    CloseDesignsTestWorkbook wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestInventoryApply_ProductionConsumeRejectsNegativeInventory() As Long
    Dim wb As Workbook
    Dim consumeEvent As Object
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim payloadJson As String
    Dim loLog As ListObject

    Set wb = TestPhase2Helpers.BuildPhase2InventoryWorkbook("WH-NEG-PROD", Array("SKU-NEG-PROD"), False)
    payloadJson = TestPhase2Helpers.BuildPayloadJson( _
        TestPhase2Helpers.CreatePayloadItem(1, "SKU-NEG-PROD", 1, "CLEARVIEW", "overdraw", "USED"))
    Set consumeEvent = TestPhase2Helpers.CreatePayloadEvent( _
        "INV-NEG-PROD-APPLY", EVENT_TYPE_PROD_CONSUME, "WH-NEG-PROD", "S-NEG", "U-NEG", payloadJson)

    On Error GoTo CleanFail
    If modInventoryApply.ApplyEvent(consumeEvent, wb, "RUN-NEG-PROD", statusOut, errorCode, errorMessage) Then GoTo CleanExit
    If errorCode <> "INSUFFICIENT_INVENTORY" Then GoTo CleanExit
    If InStr(1, errorMessage, EVENT_TYPE_PROD_CONSUME, vbTextCompare) = 0 Then GoTo CleanExit
    Set loLog = FindDesignsTestTable(wb, "tblInventoryLog")
    If loLog Is Nothing Then GoTo CleanExit
    If Not loLog.DataBodyRange Is Nothing Then
        If loLog.ListRows.Count > 0 Then GoTo CleanExit
    End If
    TestInventoryApply_ProductionConsumeRejectsNegativeInventory = 1

CleanExit:
    CloseDesignsTestWorkbook wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Private Function FindDesignsTestTable(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        On Error Resume Next
        Set FindDesignsTestTable = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not FindDesignsTestTable Is Nothing Then Exit Function
    Next ws
End Function

Private Function DesignsTestColumnExists(ByVal lo As ListObject, ByVal columnName As String) As Boolean
    On Error Resume Next
    DesignsTestColumnExists = Not lo.ListColumns(columnName) Is Nothing
    On Error GoTo 0
End Function

Private Function BuildDesignsTestEvent(ByVal eventId As String, ByVal eventType As String, _
                                       ByVal designId As String, ByVal designVersion As String, _
                                       ByVal payloadJson As String) As Object
    Dim evt As Object
    Set evt = CreateObject("Scripting.Dictionary")
    evt.CompareMode = vbTextCompare
    evt("EventID") = eventId
    evt("EventType") = eventType
    evt("CreatedAtUTC") = Now
    evt("WarehouseId") = "WH-DES"
    evt("StationId") = "S-DES"
    evt("UserId") = "U-DES"
    evt("DesignId") = designId
    evt("DesignVersion") = designVersion
    evt("PayloadJson") = payloadJson
    evt("SourceInbox") = "TEST"
    Set BuildDesignsTestEvent = evt
End Function

Private Function CountDesignsTestTables(ByVal wb As Workbook) As Long
    Dim ws As Worksheet
    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        CountDesignsTestTables = CountDesignsTestTables + ws.ListObjects.Count
    Next ws
End Function

Private Sub SetDesignsTestValue(ByVal lo As ListObject, ByVal rowIndex As Long, _
                                ByVal columnName As String, ByVal valueOut As Variant)
    lo.DataBodyRange.Cells(rowIndex, lo.ListColumns(columnName).Index).Value = valueOut
End Sub

Private Sub CloseDesignsTestWorkbook(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub
    Application.DisplayAlerts = False
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
End Sub
