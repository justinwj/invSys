Attribute VB_Name = "TestPhase2Helpers"
Option Explicit

Public Function BuildPhase2ConfigWorkbook(ByVal whId As String, ByVal stId As String, _
                                          Optional ByVal roleDefault As String = "RECEIVE", _
                                          Optional ByVal processorServiceUserId As String = "svc_processor") As Workbook
    Dim wb As Workbook
    Dim wsWh As Worksheet
    Dim wsSt As Worksheet
    Dim p As String

    Set wb = Application.Workbooks.Add
    Set wsWh = wb.Worksheets(1)
    wsWh.Name = "WarehouseConfig"
    Set wsSt = wb.Worksheets.Add(After:=wsWh)
    wsSt.Name = "StationConfig"

    wsWh.Range("A1").Resize(1, 20).Value = Array( _
        "WarehouseId", "WarehouseName", "Timezone", "DefaultLocation", _
        "BatchSize", "LockTimeoutMinutes", "HeartbeatIntervalSeconds", "MaxLockHoldMinutes", _
        "SnapshotCadence", "BackupCadence", "PathDataRoot", "PathBackupRoot", "PathSharePointRoot", _
        "DesignsEnabled", "PoisonRetryMax", "AuthCacheTTLSeconds", "ProcessorServiceUserId", _
        "FF_DesignsEnabled", "FF_OutlookAlerts", "FF_AutoSnapshot")
    wsWh.Range("A2").Resize(1, 20).Value = Array( _
        whId, "Main Warehouse", "UTC", "A1", _
        500, 3, 30, 2, _
        "PER_BATCH", "DAILY", "C:\invSys\" & whId & "\", "C:\invSys\Backups\" & whId & "\", "", _
        False, 3, 300, processorServiceUserId, _
        False, False, True)
    wsWh.ListObjects.Add(xlSrcRange, wsWh.Range("A1:T2"), , xlYes).Name = "tblWarehouseConfig"

    wsSt.Range("A1").Resize(1, 4).Value = Array("StationId", "WarehouseId", "StationName", "RoleDefault")
    wsSt.Range("A2").Resize(1, 4).Value = Array(stId, whId, Environ$("COMPUTERNAME"), roleDefault)
    wsSt.ListObjects.Add(xlSrcRange, wsSt.Range("A1:D2"), , xlYes).Name = "tblStationConfig"

    p = Environ$("TEMP") & "\" & whId & ".invSys.Config.test.xlsx"
    SaveWorkbookAsTestFile wb, p, 51
    Set BuildPhase2ConfigWorkbook = wb
End Function

Public Function BuildPhase2AuthWorkbook(ByVal whId As String, _
                                        Optional ByVal processorServiceUserId As String = "svc_processor") As Workbook
    Dim wb As Workbook
    Dim wsUsers As Worksheet
    Dim wsCaps As Worksheet
    Dim p As String

    Set wb = Application.Workbooks.Add
    Set wsUsers = wb.Worksheets(1)
    wsUsers.Name = "Users"
    Set wsCaps = wb.Worksheets.Add(After:=wsUsers)
    wsCaps.Name = "Capabilities"

    wsUsers.Range("A1").Resize(1, 6).Value = Array("UserId", "DisplayName", "PinHash", "Status", "ValidFrom", "ValidTo")
    wsUsers.Range("A2").Resize(1, 6).Value = Array("user1", "User One", "", "Active", "", "")
    wsUsers.Range("A3").Resize(1, 6).Value = Array("user2", "User Two", "", "Active", "", "")
    wsUsers.Range("A4").Resize(1, 6).Value = Array(processorServiceUserId, "Processor Service", "", "Active", "", "")
    wsUsers.ListObjects.Add(xlSrcRange, wsUsers.Range("A1:F4"), , xlYes).Name = "tblUsers"

    wsCaps.Range("A1").Resize(1, 7).Value = Array("UserId", "Capability", "WarehouseId", "StationId", "Status", "ValidFrom", "ValidTo")
    wsCaps.Range("A2").Resize(1, 7).Value = Array("", "", "", "", "", "", "")
    wsCaps.ListObjects.Add(xlSrcRange, wsCaps.Range("A1:G2"), , xlYes).Name = "tblCapabilities"

    p = Environ$("TEMP") & "\" & whId & ".invSys.Auth.test.xlsx"
    SaveWorkbookAsTestFile wb, p, 51
    Set BuildPhase2AuthWorkbook = wb
End Function

Public Function BuildPhase2InventoryWorkbook(ByVal whId As String, Optional ByVal skuList As Variant) As Workbook
    Dim wb As Workbook
    Dim wsSku As Worksheet
    Dim loSku As ListObject
    Dim lastRow As Long
    Dim i As Long
    Dim p As String

    Set wb = Application.Workbooks.Add
    wb.Worksheets(1).Name = "InventoryLog"
    Call modInventorySchema.EnsureInventorySchema(wb)
    DeleteAllTableRows wb.Worksheets("InventoryLog").ListObjects("tblInventoryLog"), True
    DeleteAllTableRows wb.Worksheets("AppliedEvents").ListObjects("tblAppliedEvents"), True
    DeleteAllTableRows wb.Worksheets("Locks").ListObjects("tblLocks"), True

    If Not IsMissing(skuList) Then
        Set wsSku = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        wsSku.Name = "SkuCatalog"
        wsSku.Range("A1").Value = "SKU"
        lastRow = 1
        If IsArray(skuList) Then
            For i = LBound(skuList) To UBound(skuList)
                lastRow = lastRow + 1
                wsSku.Cells(lastRow, 1).Value = skuList(i)
            Next i
        ElseIf CStr(skuList) <> "" Then
            lastRow = 2
            wsSku.Cells(lastRow, 1).Value = CStr(skuList)
        End If
        If lastRow = 1 Then lastRow = 2
        Set loSku = wsSku.ListObjects.Add(xlSrcRange, wsSku.Range("A1:A" & CStr(lastRow)), , xlYes)
        loSku.Name = "tblSkuCatalog"
    End If

    p = Environ$("TEMP") & "\" & whId & ".invSys.Data.Inventory.test.xlsx"
    SaveWorkbookAsTestFile wb, p, 51
    Set BuildPhase2InventoryWorkbook = wb
End Function

Public Function BuildPhase2InboxWorkbook(Optional ByVal stationId As String = "S1") As Workbook
    Dim wb As Workbook
    Dim report As String
    Dim p As String

    Set wb = Application.Workbooks.Add
    wb.Worksheets(1).Name = "InboxReceive"
    Call modProcessor.EnsureReceiveInboxSchema(wb, report)
    DeleteAllTableRows wb.Worksheets("InboxReceive").ListObjects("tblInboxReceive"), False

    p = Environ$("TEMP") & "\invSys.Inbox.Receiving." & stationId & ".test.xlsx"
    SaveWorkbookAsTestFile wb, p, 51
    Set BuildPhase2InboxWorkbook = wb
End Function

Public Sub AddCapability(ByVal wb As Workbook, _
                         ByVal userId As String, _
                         ByVal capability As String, _
                         ByVal whId As String, _
                         ByVal stId As String, _
                         ByVal status As String, _
                         Optional ByVal validFrom As String = "", _
                         Optional ByVal validTo As String = "")
    Dim lo As ListObject
    Dim r As ListRow

    Set lo = wb.Worksheets("Capabilities").ListObjects("tblCapabilities")
    EnsureTableSheetEditable lo, "tblCapabilities"
    Set r = lo.ListRows.Add
    r.Range.Cells(1, lo.ListColumns("UserId").Index).Value = userId
    r.Range.Cells(1, lo.ListColumns("Capability").Index).Value = capability
    r.Range.Cells(1, lo.ListColumns("WarehouseId").Index).Value = whId
    r.Range.Cells(1, lo.ListColumns("StationId").Index).Value = stId
    r.Range.Cells(1, lo.ListColumns("Status").Index).Value = status
    r.Range.Cells(1, lo.ListColumns("ValidFrom").Index).Value = validFrom
    r.Range.Cells(1, lo.ListColumns("ValidTo").Index).Value = validTo
End Sub

Public Sub AddInboxReceiveRow(ByVal wb As Workbook, _
                              ByVal eventId As String, _
                              ByVal createdAtUtc As Variant, _
                              ByVal whId As String, _
                              ByVal stId As String, _
                              ByVal userId As String, _
                              ByVal sku As String, _
                              ByVal qty As Double, _
                              Optional ByVal locationVal As String = "", _
                              Optional ByVal noteVal As String = "")
    Dim lo As ListObject
    Dim r As ListRow

    Set lo = wb.Worksheets("InboxReceive").ListObjects("tblInboxReceive")
    EnsureTableSheetEditable lo, "tblInboxReceive"
    Set r = lo.ListRows.Add
    SetTableRowValue lo, r.Index, "EventID", eventId
    SetTableRowValue lo, r.Index, "CreatedAtUTC", createdAtUtc
    SetTableRowValue lo, r.Index, "WarehouseId", whId
    SetTableRowValue lo, r.Index, "StationId", stId
    SetTableRowValue lo, r.Index, "UserId", userId
    SetTableRowValue lo, r.Index, "SKU", sku
    SetTableRowValue lo, r.Index, "Qty", qty
    SetTableRowValue lo, r.Index, "Location", locationVal
    SetTableRowValue lo, r.Index, "Note", noteVal
    SetTableRowValue lo, r.Index, "Status", "NEW"
    SetTableRowValue lo, r.Index, "RetryCount", 0
End Sub

Public Function CreateReceiveEvent(ByVal eventId As String, _
                                   ByVal whId As String, _
                                   ByVal stId As String, _
                                   ByVal userId As String, _
                                   ByVal sku As String, _
                                   ByVal qty As Double, _
                                   Optional ByVal locationVal As String = "", _
                                   Optional ByVal noteVal As String = "", _
                                   Optional ByVal createdAtUtc As Variant = Empty, _
                                   Optional ByVal sourceInbox As String = "test-inbox") As Object
    Dim evt As Object

    Set evt = CreateObject("Scripting.Dictionary")
    evt.CompareMode = vbTextCompare
    evt("EventID") = eventId
    evt("CreatedAtUTC") = IIf(IsEmpty(createdAtUtc), Now, createdAtUtc)
    evt("WarehouseId") = whId
    evt("StationId") = stId
    evt("UserId") = userId
    evt("SKU") = sku
    evt("Qty") = qty
    evt("Location") = locationVal
    evt("Note") = noteVal
    evt("SourceInbox") = sourceInbox
    Set CreateReceiveEvent = evt
End Function

Public Function TableExists(ByVal wb As Workbook, ByVal tableName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    For Each ws In wb.Worksheets
        If Not ws.ListObjects(tableName) Is Nothing Then
            TableExists = True
            Exit Function
        End If
    Next ws
    On Error GoTo 0
End Function

Public Function GetRowValue(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As Variant
    Dim idx As Long
    idx = lo.ListColumns(columnName).Index
    GetRowValue = lo.DataBodyRange.Cells(rowIndex, idx).Value
End Function

Public Sub CloseNoSave(ByVal wb As Workbook)
    Dim p As String
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    p = wb.FullName
    wb.Close SaveChanges:=False
    If InStr(1, p, ".test.", vbTextCompare) > 0 Then
        If Len(Dir$(p)) > 0 Then Kill p
    End If
    On Error GoTo 0
End Sub

Private Sub SaveWorkbookAsTestFile(ByVal wb As Workbook, ByVal pathOut As String, ByVal fileFormat As Long)
    On Error Resume Next
    Kill pathOut
    On Error GoTo 0
    wb.SaveAs Filename:=pathOut, FileFormat:=fileFormat
End Sub

Private Sub SetTableRowValue(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String, ByVal valueOut As Variant)
    lo.DataBodyRange.Cells(rowIndex, lo.ListColumns(columnName).Index).Value = valueOut
End Sub

Private Sub DeleteAllTableRows(ByVal lo As ListObject, ByVal reprotectAfter As Boolean)
    On Error Resume Next
    lo.Parent.Unprotect
    On Error GoTo 0
    Do While lo.ListRows.Count > 0
        lo.ListRows(lo.ListRows.Count).Delete
    Loop
    If reprotectAfter Then
        On Error Resume Next
        lo.Parent.Protect UserInterfaceOnly:=True
        On Error GoTo 0
    End If
End Sub

Private Sub EnsureTableSheetEditable(ByVal lo As ListObject, ByVal tableName As String)
    If lo Is Nothing Then Exit Sub
    If Not lo.Parent.ProtectContents Then Exit Sub

    On Error Resume Next
    lo.Parent.Unprotect
    On Error GoTo 0

    If lo.Parent.ProtectContents Then
        Err.Raise vbObjectError + 2601, "TestPhase2Helpers.EnsureTableSheetEditable", _
                  "Worksheet '" & lo.Parent.Name & "' is protected and could not be unprotected before writing to " & tableName & "."
    End If
End Sub
