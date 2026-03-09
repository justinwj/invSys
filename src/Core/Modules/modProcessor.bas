Attribute VB_Name = "modProcessor"
Option Explicit

Public Const INBOX_STATUS_NEW As String = "NEW"
Public Const INBOX_STATUS_PROCESSED As String = "PROCESSED"
Public Const INBOX_STATUS_SKIP_DUP As String = "SKIP_DUP"
Public Const INBOX_STATUS_POISON As String = "POISON"

Public Function RunBatch(Optional ByVal warehouseId As String = "", _
                         Optional ByVal batchSize As Long = 0, _
                         Optional ByRef report As String = "") As Long
    On Error GoTo FailRun

    Dim inventoryWb As Workbook
    Dim inboxWorkbooks As Collection
    Dim inboxWb As Variant
    Dim loInbox As ListObject
    Dim rowIndex As Long
    Dim runId As String
    Dim message As String
    Dim serviceUserId As String
    Dim skipDupCount As Long
    Dim poisonCount As Long
    Dim heartbeatSeconds As Long
    Dim lastHeartbeat As Date
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim evt As Object
    Dim lockHeld As Boolean

    If Not EnsurePhase2Context(warehouseId, report) Then Exit Function

    warehouseId = modConfig.GetString("WarehouseId", warehouseId)
    If warehouseId = "" Then
        report = "WarehouseId not resolved."
        Exit Function
    End If

    serviceUserId = modConfig.GetString("ProcessorServiceUserId", "svc_processor")
    If serviceUserId = "" Then serviceUserId = "svc_processor"

    If batchSize <= 0 Then batchSize = modConfig.GetLong("BatchSize", 500)
    If batchSize <= 0 Then batchSize = 500

    heartbeatSeconds = modConfig.GetLong("HeartbeatIntervalSeconds", 30)
    If heartbeatSeconds <= 0 Then heartbeatSeconds = 30

    If Not modAuth.CanPerform("INBOX_PROCESS", serviceUserId, warehouseId, modConfig.GetString("StationId", ""), "PROCESSOR", "PROCESSOR-RUN") Then
        report = "Processor service identity lacks INBOX_PROCESS."
        Exit Function
    End If

    Set inventoryWb = modInventoryApply.ResolveInventoryWorkbook(warehouseId)
    If inventoryWb Is Nothing Then
        report = "Inventory workbook not found."
        Exit Function
    End If

    If Not modLockManager.AcquireLock("INVENTORY", warehouseId, serviceUserId, modConfig.GetString("StationId", ""), inventoryWb, runId, message) Then
        report = message
        Exit Function
    End If
    lockHeld = True
    lastHeartbeat = Now

    Set inboxWorkbooks = ResolveReceiveInboxWorkbooks()
    For Each inboxWb In inboxWorkbooks
        If Not EnsureReceiveInboxSchema(inboxWb) Then GoTo ContinueInbox

        Set loInbox = FindListObjectByNameProcessor(inboxWb, "tblInboxReceive")
        If loInbox Is Nothing Then GoTo ContinueInbox
        If loInbox.DataBodyRange Is Nothing Then GoTo ContinueInbox

        For rowIndex = 1 To loInbox.ListRows.Count
            If RunBatch >= batchSize Then Exit For
            If Not IsProcessableInboxRow(loInbox, rowIndex, warehouseId) Then GoTo ContinueRow

            Set evt = BuildInboxEvent(loInbox, rowIndex, inboxWb.Name)
            If evt Is Nothing Then
                UpdateInboxRowStatus loInbox, rowIndex, INBOX_STATUS_POISON, "INVALID_EVENT", "Unable to read inbox row."
                poisonCount = poisonCount + 1
                GoTo MaybeHeartbeat
            End If

            If Not modAuth.CanPerform("RECEIVE_POST", GetDictionaryString(evt, "UserId"), GetDictionaryString(evt, "WarehouseId"), GetDictionaryString(evt, "StationId"), "PROCESSOR_VALIDATE", GetDictionaryString(evt, "EventID")) Then
                UpdateInboxRowStatus loInbox, rowIndex, INBOX_STATUS_POISON, "AUTH_DENIED", "Event creator lacks RECEIVE_POST capability."
                poisonCount = poisonCount + 1
                GoTo MaybeHeartbeat
            End If

            statusOut = vbNullString
            errorCode = vbNullString
            errorMessage = vbNullString

            If modInventoryApply.ApplyReceiveEvent(evt, inventoryWb, runId, statusOut, errorCode, errorMessage) Then
                Select Case UCase$(statusOut)
                    Case APPLY_STATUS_APPLIED
                        UpdateInboxRowStatus loInbox, rowIndex, INBOX_STATUS_PROCESSED
                        RunBatch = RunBatch + 1
                    Case APPLY_STATUS_SKIP_DUP
                        UpdateInboxRowStatus loInbox, rowIndex, INBOX_STATUS_SKIP_DUP
                        skipDupCount = skipDupCount + 1
                    Case Else
                        UpdateInboxRowStatus loInbox, rowIndex, INBOX_STATUS_POISON, "UNKNOWN_APPLY_STATUS", "Unknown apply status."
                        poisonCount = poisonCount + 1
                End Select
            Else
                UpdateInboxRowStatus loInbox, rowIndex, INBOX_STATUS_POISON, errorCode, errorMessage
                poisonCount = poisonCount + 1
            End If

MaybeHeartbeat:
            If DateDiff("s", lastHeartbeat, Now) >= heartbeatSeconds Then
                Call modLockManager.UpdateHeartbeat("INVENTORY", runId, inventoryWb)
                lastHeartbeat = Now
            End If

ContinueRow:
        Next rowIndex

        If RunBatch >= batchSize Then Exit For
ContinueInbox:
    Next inboxWb

    report = "Applied=" & CStr(RunBatch) & "; SkipDup=" & CStr(skipDupCount) & "; Poison=" & CStr(poisonCount) & "; RunId=" & runId

CleanExit:
    If lockHeld Then Call modLockManager.ReleaseLock("INVENTORY", runId, inventoryWb)
    Exit Function

FailRun:
    report = "RunBatch failed: " & Err.Description
    Resume CleanExit
End Function

Public Function EnsureReceiveInboxSchema(Optional ByVal targetWb As Workbook = Nothing, _
                                         Optional ByRef report As String = "") As Boolean
    On Error GoTo FailEnsure

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim headers As Variant
    Dim startCell As Range
    Dim dataRange As Range
    Dim i As Long

    If targetWb Is Nothing Then
        Set wb = ResolveSingleReceiveInboxWorkbook()
    Else
        Set wb = targetWb
    End If
    If wb Is Nothing Then
        report = "Inbox workbook not found."
        Exit Function
    End If

    headers = Array("EventID", "ParentEventId", "UndoOfEventId", "CreatedAtUTC", "WarehouseId", "StationId", _
                    "UserId", "SKU", "Qty", "Location", "Note", "Status", "RetryCount", "ErrorCode", _
                    "ErrorMessage", "FailedAtUTC")

    Set ws = EnsureWorksheetProcessor(wb, "InboxReceive")
    SetSheetProtectionProcessor ws, False
    On Error Resume Next
    Set lo = ws.ListObjects("tblInboxReceive")
    On Error GoTo 0

    If lo Is Nothing Then
        Set startCell = GetNextTableStartCellProcessor(ws)
        For i = LBound(headers) To UBound(headers)
            startCell.Offset(0, i - LBound(headers)).Value = headers(i)
        Next i

        Set dataRange = ws.Range(startCell, startCell.Offset(1, UBound(headers) - LBound(headers)))
        Set lo = ws.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
        lo.Name = "tblInboxReceive"
    End If

    For i = LBound(headers) To UBound(headers)
        EnsureListColumnProcessor lo, CStr(headers(i))
    Next i

    EnsureTableHasRowProcessor lo
    report = "OK"
    EnsureReceiveInboxSchema = True
    SetSheetProtectionProcessor ws, True
    Exit Function

FailEnsure:
    On Error Resume Next
    If Not ws Is Nothing Then SetSheetProtectionProcessor ws, True
    On Error GoTo 0
    report = "EnsureReceiveInboxSchema failed: " & Err.Description
End Function

Private Function EnsurePhase2Context(ByVal warehouseId As String, ByRef report As String) As Boolean
    If Not modConfig.LoadConfig(warehouseId, "") Then
        report = "Config load failed: " & modConfig.Validate()
        Exit Function
    End If

    If Not modAuth.LoadAuth(modConfig.GetString("WarehouseId", warehouseId)) Then
        report = "Auth load failed: " & modAuth.ValidateAuth()
        Exit Function
    End If

    EnsurePhase2Context = True
End Function

Private Function ResolveReceiveInboxWorkbooks() As Collection
    Dim wb As Workbook
    Set ResolveReceiveInboxWorkbooks = New Collection

    For Each wb In Application.Workbooks
        If IsReceiveInboxWorkbookName(wb.Name) Then
            ResolveReceiveInboxWorkbooks.Add wb
        End If
    Next wb

    If ResolveReceiveInboxWorkbooks.Count > 0 Then Exit Function

    For Each wb In Application.Workbooks
        If WorkbookHasListObjectProcessor(wb, "tblInboxReceive") Then
            ResolveReceiveInboxWorkbooks.Add wb
        End If
    Next wb
End Function

Private Function ResolveSingleReceiveInboxWorkbook() As Workbook
    Dim workbooks As Collection
    Set workbooks = ResolveReceiveInboxWorkbooks()
    If workbooks.Count > 0 Then Set ResolveSingleReceiveInboxWorkbook = workbooks(1)
End Function

Private Function BuildInboxEvent(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal workbookName As String) As Object
    Dim evt As Object
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    Set evt = CreateObject("Scripting.Dictionary")
    evt.CompareMode = vbTextCompare
    evt("EventID") = GetCellByColumnProcessor(lo, rowIndex, "EventID")
    evt("ParentEventId") = GetCellByColumnProcessor(lo, rowIndex, "ParentEventId")
    evt("UndoOfEventId") = GetCellByColumnProcessor(lo, rowIndex, "UndoOfEventId")
    evt("CreatedAtUTC") = GetCellByColumnProcessor(lo, rowIndex, "CreatedAtUTC")
    evt("WarehouseId") = GetCellByColumnProcessor(lo, rowIndex, "WarehouseId")
    evt("StationId") = GetCellByColumnProcessor(lo, rowIndex, "StationId")
    evt("UserId") = GetCellByColumnProcessor(lo, rowIndex, "UserId")
    evt("SKU") = GetCellByColumnProcessor(lo, rowIndex, "SKU")
    evt("Qty") = GetCellByColumnProcessor(lo, rowIndex, "Qty")
    evt("Location") = GetCellByColumnProcessor(lo, rowIndex, "Location")
    evt("Note") = GetCellByColumnProcessor(lo, rowIndex, "Note")
    evt("SourceInbox") = workbookName & ":tblInboxReceive"
    Set BuildInboxEvent = evt
End Function

Private Function IsProcessableInboxRow(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal warehouseId As String) As Boolean
    Dim statusVal As String
    Dim eventId As String
    Dim rowWarehouse As String

    eventId = SafeTrimProcessor(GetCellByColumnProcessor(lo, rowIndex, "EventID"))
    If eventId = "" Then Exit Function

    statusVal = UCase$(SafeTrimProcessor(GetCellByColumnProcessor(lo, rowIndex, "Status")))
    If statusVal <> "" And statusVal <> INBOX_STATUS_NEW Then Exit Function

    rowWarehouse = SafeTrimProcessor(GetCellByColumnProcessor(lo, rowIndex, "WarehouseId"))
    If warehouseId <> "" And rowWarehouse <> "" Then
        If StrComp(warehouseId, rowWarehouse, vbTextCompare) <> 0 Then Exit Function
    End If

    IsProcessableInboxRow = True
End Function

Private Sub UpdateInboxRowStatus(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal newStatus As String, _
                                 Optional ByVal errorCode As String = "", Optional ByVal errorMessage As String = "")
    Dim retryCount As Long
    If lo Is Nothing Then Exit Sub

    SetSheetProtectionProcessor lo.Parent, False

    SetCellByColumnProcessor lo, rowIndex, "Status", newStatus

    Select Case UCase$(newStatus)
        Case INBOX_STATUS_POISON
            retryCount = 0
            If IsNumeric(GetCellByColumnProcessor(lo, rowIndex, "RetryCount")) Then
                retryCount = CLng(GetCellByColumnProcessor(lo, rowIndex, "RetryCount"))
            End If
            SetCellByColumnProcessor lo, rowIndex, "RetryCount", retryCount + 1
            SetCellByColumnProcessor lo, rowIndex, "ErrorCode", errorCode
            SetCellByColumnProcessor lo, rowIndex, "ErrorMessage", errorMessage
            SetCellByColumnProcessor lo, rowIndex, "FailedAtUTC", Now
        Case Else
            SetCellByColumnProcessor lo, rowIndex, "ErrorCode", vbNullString
            SetCellByColumnProcessor lo, rowIndex, "ErrorMessage", vbNullString
            SetCellByColumnProcessor lo, rowIndex, "FailedAtUTC", vbNullString
    End Select

    SetSheetProtectionProcessor lo.Parent, True
End Sub

Private Function GetDictionaryString(ByVal d As Object, ByVal key As String) As String
    On Error Resume Next
    GetDictionaryString = SafeTrimProcessor(d(key))
    On Error GoTo 0
End Function

Private Function IsReceiveInboxWorkbookName(ByVal wbName As String) As Boolean
    Dim n As String
    n = LCase$(wbName)
    IsReceiveInboxWorkbookName = (n Like "invsys.inbox.receiving.*.xlsb") Or _
                                 (n Like "invsys.inbox.receiving.*.xlsx") Or _
                                 (n Like "invsys.inbox.receiving.*.xlsm")
End Function

Private Sub EnsureListColumnProcessor(ByVal lo As ListObject, ByVal columnName As String)
    If GetColumnIndexProcessor(lo, columnName) > 0 Then Exit Sub
    lo.ListColumns.Add lo.ListColumns.Count + 1
    lo.ListColumns(lo.ListColumns.Count).Name = columnName
End Sub

Private Sub EnsureTableHasRowProcessor(ByVal lo As ListObject)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then lo.ListRows.Add
End Sub

Private Function EnsureWorksheetProcessor(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureWorksheetProcessor = wb.Worksheets(sheetName)
    On Error GoTo 0

    If EnsureWorksheetProcessor Is Nothing Then
        Set EnsureWorksheetProcessor = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        EnsureWorksheetProcessor.Name = sheetName
    End If
End Function

Private Function GetNextTableStartCellProcessor(ByVal ws As Worksheet) As Range
    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then
        Set GetNextTableStartCellProcessor = ws.Range("A1")
    Else
        Set GetNextTableStartCellProcessor = ws.Cells(ws.Rows.Count, 1).End(xlUp).Offset(2, 0)
    End If
End Function

Private Function WorkbookHasListObjectProcessor(ByVal wb As Workbook, ByVal tableName As String) As Boolean
    WorkbookHasListObjectProcessor = Not (FindListObjectByNameProcessor(wb, tableName) Is Nothing)
End Function

Private Function FindListObjectByNameProcessor(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    On Error Resume Next
    For Each ws In wb.Worksheets
        Set FindListObjectByNameProcessor = ws.ListObjects(tableName)
        If Not FindListObjectByNameProcessor Is Nothing Then Exit Function
    Next ws
    On Error GoTo 0
End Function

Private Function GetCellByColumnProcessor(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As Variant
    Dim idx As Long
    idx = GetColumnIndexProcessor(lo, columnName)
    If idx = 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    GetCellByColumnProcessor = lo.DataBodyRange.Cells(rowIndex, idx).Value
End Function

Private Sub SetCellByColumnProcessor(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String, ByVal valueOut As Variant)
    Dim idx As Long
    idx = GetColumnIndexProcessor(lo, columnName)
    If idx = 0 Then Exit Sub
    lo.DataBodyRange.Cells(rowIndex, idx).Value = valueOut
End Sub

Private Function GetColumnIndexProcessor(ByVal lo As ListObject, ByVal columnName As String) As Long
    Dim i As Long
    If lo Is Nothing Then Exit Function
    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).Name, columnName, vbTextCompare) = 0 Then
            GetColumnIndexProcessor = i
            Exit Function
        End If
    Next i
End Function

Private Function SafeTrimProcessor(ByVal valueIn As Variant) As String
    On Error Resume Next
    SafeTrimProcessor = Trim$(CStr(valueIn))
End Function

Private Sub SetSheetProtectionProcessor(ByVal ws As Worksheet, ByVal protectAfter As Boolean)
    If ws Is Nothing Then Exit Sub
    If protectAfter Then
        On Error Resume Next
        ws.Protect UserInterfaceOnly:=True
        On Error GoTo 0
    Else
        If Not ws.ProtectContents Then Exit Sub
        On Error Resume Next
        ws.Unprotect
        On Error GoTo 0
        If ws.ProtectContents Then
            Err.Raise vbObjectError + 2401, "modProcessor.SetSheetProtectionProcessor", _
                      "Worksheet '" & ws.Name & "' is protected and could not be unprotected. " & _
                      "Excel automation cannot update inbox tables while the sheet remains protected."
        End If
    End If
End Sub
