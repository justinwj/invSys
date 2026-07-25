Attribute VB_Name = "modDesignsApply"
Option Explicit

Public Const DESIGN_STATUS_APPLIED As String = "APPLIED"
Public Const DESIGN_STATUS_SKIP_DUP As String = "SKIP_DUP"

Public Const DESIGN_EVENT_CREATE As String = "DESIGN_CREATE"
Public Const DESIGN_EVENT_RELEASE As String = "DESIGN_RELEASE"
Public Const DESIGN_EVENT_OBSOLETE As String = "DESIGN_OBSOLETE"

Public Function ApplyDesignEvent(ByVal evt As Object, _
                                 Optional ByVal designsWb As Workbook = Nothing, _
                                 Optional ByVal runId As String = "", _
                                 Optional ByRef statusOut As String = "", _
                                 Optional ByRef errorCode As String = "", _
                                 Optional ByRef errorMessage As String = "") As Boolean
    On Error GoTo FailApply

    Dim wb As Workbook
    Dim loEvents As ListObject
    Dim loApplied As ListObject
    Dim eventId As String
    Dim eventType As String
    Dim designId As String
    Dim designVersion As String
    Dim payloadJson As String
    Dim occurredAt As Date
    Dim appliedAt As Date
    Dim appliedSeq As Long
    Dim report As String
    Dim lr As ListRow

    Set wb = modDesignsRuntime.ResolveDesignsWorkbook(GetEventTextDesign(evt, "WarehouseId"), designsWb, report)
    If wb Is Nothing Then
        errorCode = "DESIGNS_WORKBOOK_NOT_FOUND"
        errorMessage = report
        Exit Function
    End If
    If Not modDesignsSchema.EnsureDesignsSchema(wb, report) Then
        errorCode = "DESIGNS_SCHEMA_INVALID"
        errorMessage = report
        Exit Function
    End If

    Set loEvents = FindDesignsApplyTable(wb, "tblDesignEvents")
    Set loApplied = FindDesignsApplyTable(wb, "tblAppliedDesignEvents")
    eventId = GetEventTextDesign(evt, "EventID")
    eventType = UCase$(GetEventTextDesign(evt, "EventType"))
    designId = GetEventTextDesign(evt, "DesignId")
    designVersion = GetEventTextDesign(evt, "DesignVersion")
    payloadJson = GetEventTextDesign(evt, "PayloadJson")

    If eventId = "" Or designId = "" Or designVersion = "" Then
        errorCode = "INVALID_DESIGN_EVENT"
        errorMessage = "EventID, DesignId, and DesignVersion are required."
        Exit Function
    End If
    If GetEventTextDesign(evt, "WarehouseId") = "" _
       Or GetEventTextDesign(evt, "StationId") = "" _
       Or GetEventTextDesign(evt, "UserId") = "" Then
        errorCode = "INVALID_DESIGN_EVENT"
        errorMessage = "WarehouseId, StationId, and UserId are required."
        Exit Function
    End If
    If Not TryGetEventDateDesign(evt, "CreatedAtUTC", occurredAt) Then
        errorCode = "INVALID_DESIGN_EVENT"
        errorMessage = "CreatedAtUTC is required and must be a valid date."
        Exit Function
    End If
    If AppliedDesignEventExists(loApplied, eventId) Then
        statusOut = DESIGN_STATUS_SKIP_DUP
        ApplyDesignEvent = True
        Exit Function
    End If
    ' Lifecycle decisions are derived from authoritative event history, never
    ' from whatever projection rows happen to be present in the workbook.
    If Not RebuildDesignProjections(wb, report) Then
        errorCode = "DESIGN_PROJECTION_FAILED"
        errorMessage = report
        Exit Function
    End If
    If Not ValidateLifecycleTransition(wb, eventType, designId, designVersion, payloadJson, errorCode, errorMessage) Then
        Exit Function
    End If

    appliedAt = Now
    appliedSeq = NextDesignAppliedSeq(loApplied)
    If runId = "" Then runId = "DESIGN-RUN-" & Format$(appliedAt, "yyyymmddhhnnss")

    Set lr = loEvents.ListRows.Add
    SetDesignTableValue loEvents, lr.Index, "EventID", eventId
    SetDesignTableValue loEvents, lr.Index, "UndoOfEventId", GetEventTextDesign(evt, "UndoOfEventId")
    SetDesignTableValue loEvents, lr.Index, "AppliedSeq", appliedSeq
    SetDesignTableValue loEvents, lr.Index, "EventType", eventType
    SetDesignTableValue loEvents, lr.Index, "OccurredAtUTC", occurredAt
    SetDesignTableValue loEvents, lr.Index, "AppliedAtUTC", appliedAt
    SetDesignTableValue loEvents, lr.Index, "WarehouseId", GetEventTextDesign(evt, "WarehouseId")
    SetDesignTableValue loEvents, lr.Index, "StationId", GetEventTextDesign(evt, "StationId")
    SetDesignTableValue loEvents, lr.Index, "UserId", GetEventTextDesign(evt, "UserId")
    SetDesignTableValue loEvents, lr.Index, "DesignId", designId
    SetDesignTableValue loEvents, lr.Index, "DesignVersion", designVersion
    SetDesignTableValue loEvents, lr.Index, "PayloadJson", payloadJson
    SetDesignTableValue loEvents, lr.Index, "Note", GetEventTextDesign(evt, "Note")

    Set lr = loApplied.ListRows.Add
    SetDesignTableValue loApplied, lr.Index, "EventID", eventId
    SetDesignTableValue loApplied, lr.Index, "UndoOfEventId", GetEventTextDesign(evt, "UndoOfEventId")
    SetDesignTableValue loApplied, lr.Index, "AppliedSeq", appliedSeq
    SetDesignTableValue loApplied, lr.Index, "AppliedAtUTC", appliedAt
    SetDesignTableValue loApplied, lr.Index, "RunId", runId
    SetDesignTableValue loApplied, lr.Index, "SourceInbox", GetEventTextDesign(evt, "SourceInbox")
    SetDesignTableValue loApplied, lr.Index, "Status", DESIGN_STATUS_APPLIED

    If Not RebuildDesignProjections(wb, report) Then
        errorCode = "DESIGN_PROJECTION_FAILED"
        errorMessage = report
        Exit Function
    End If
    SaveDesignsWorkbook wb
    statusOut = DESIGN_STATUS_APPLIED
    ApplyDesignEvent = True
    Exit Function

FailApply:
    If errorCode = "" Then errorCode = "DESIGN_APPLY_EXCEPTION"
    If errorMessage = "" Then errorMessage = CStr(Err.Number) & ": " & Err.Description
End Function

Public Function RebuildDesignProjections(ByVal designsWb As Workbook, _
                                         Optional ByRef report As String = "") As Boolean
    On Error GoTo FailRebuild

    Dim loEvents As ListObject
    Dim loDesigns As ListObject
    Dim loLines As ListObject
    Dim values As Variant
    Dim r As Long
    Dim replayOrder As Variant
    Dim orderIndex As Long
    Dim errorMessage As String

    If designsWb Is Nothing Then
        report = "An explicit authoritative Designs workbook is required."
        Exit Function
    End If
    If designsWb.IsAddin Then
        report = "An explicit authoritative Designs workbook is required."
        Exit Function
    End If
    If Not modDesignsSchema.EnsureDesignsSchema(designsWb, report) Then Exit Function
    Set loEvents = FindDesignsApplyTable(designsWb, "tblDesignEvents")
    Set loDesigns = FindDesignsApplyTable(designsWb, "tblDesigns")
    Set loLines = FindDesignsApplyTable(designsWb, "tblDesignLines")
    ClearDesignTableRows loDesigns
    ClearDesignTableRows loLines

    If Not loEvents.DataBodyRange Is Nothing Then
        values = loEvents.DataBodyRange.Value
        replayOrder = DesignEventReplayOrder(loEvents, values, errorMessage)
        If IsEmpty(replayOrder) Then
            report = "Projection replay order is invalid: " & errorMessage
            Exit Function
        End If
        For orderIndex = LBound(replayOrder) To UBound(replayOrder)
            r = CLng(replayOrder(orderIndex))
            If Not ReplayDesignEvent(loEvents, values, r, loDesigns, loLines, errorMessage) Then
                report = "Projection replay failed at event row " & CStr(r) & ": " & errorMessage
                Exit Function
            End If
        Next orderIndex
    End If
    report = "OK"
    RebuildDesignProjections = True
    Exit Function

FailRebuild:
    report = "RebuildDesignProjections failed: " & Err.Description
End Function

Private Function DesignEventReplayOrder(ByVal loEvents As ListObject, ByVal values As Variant, _
                                        ByRef errorMessage As String) As Variant
    Dim rowCount As Long
    Dim order() As Long
    Dim seqs() As Long
    Dim seen As Object
    Dim i As Long
    Dim j As Long
    Dim seqValue As Variant
    Dim swapValue As Long

    rowCount = UBound(values, 1)
    If rowCount <= 0 Then Exit Function
    ReDim order(1 To rowCount)
    ReDim seqs(1 To rowCount)
    Set seen = CreateObject("Scripting.Dictionary")

    For i = 1 To rowCount
        seqValue = ReadDesignTableValue(loEvents, values, i, "AppliedSeq")
        If Not IsNumeric(seqValue) Then
            errorMessage = "AppliedSeq is required at event row " & CStr(i) & "."
            Exit Function
        End If
        seqs(i) = CLng(seqValue)
        If seqs(i) <= 0 Then
            errorMessage = "AppliedSeq must be positive at event row " & CStr(i) & "."
            Exit Function
        End If
        If seen.Exists(CStr(seqs(i))) Then
            errorMessage = "Duplicate AppliedSeq " & CStr(seqs(i)) & "."
            Exit Function
        End If
        seen.Add CStr(seqs(i)), True
        order(i) = i
    Next i

    For i = 1 To rowCount - 1
        For j = i + 1 To rowCount
            If seqs(order(j)) < seqs(order(i)) Then
                swapValue = order(i)
                order(i) = order(j)
                order(j) = swapValue
            End If
        Next j
    Next i
    DesignEventReplayOrder = order
End Function

Private Function ValidateLifecycleTransition(ByVal wb As Workbook, ByVal eventType As String, _
                                             ByVal designId As String, ByVal designVersion As String, _
                                             ByVal payloadJson As String, ByRef errorCode As String, _
                                             ByRef errorMessage As String) As Boolean
    Dim currentStatus As String
    Dim payload As Collection

    currentStatus = CurrentDesignStatus(wb, designId, designVersion)
    Select Case eventType
        Case DESIGN_EVENT_CREATE
            If currentStatus <> "" Then
                errorCode = "DESIGN_VERSION_EXISTS"
                errorMessage = "DesignId and DesignVersion are immutable once created."
                Exit Function
            End If
            Set payload = ParseDesignPayload(payloadJson, errorMessage)
            If payload Is Nothing Then
                errorCode = "INVALID_DESIGN_PAYLOAD"
                If errorMessage = "" Then errorMessage = "DESIGN_CREATE requires at least one payload row."
                Exit Function
            End If
            If payload.Count = 0 Then
                errorCode = "INVALID_DESIGN_PAYLOAD"
                errorMessage = "DESIGN_CREATE requires at least one payload row."
                Exit Function
            End If
            If PayloadText(payload(1), "DesignType") = "" Or PayloadText(payload(1), "DesignName") = "" Then
                errorCode = "INVALID_DESIGN_PAYLOAD"
                errorMessage = "DESIGN_CREATE payload requires DesignType and DesignName."
                Exit Function
            End If
        Case DESIGN_EVENT_RELEASE
            If StrComp(currentStatus, "DRAFT", vbTextCompare) <> 0 Then
                errorCode = "INVALID_DESIGN_TRANSITION"
                errorMessage = "Only a DRAFT design version can be released."
                Exit Function
            End If
        Case DESIGN_EVENT_OBSOLETE
            If currentStatus = "" Or StrComp(currentStatus, "OBSOLETE", vbTextCompare) = 0 Then
                errorCode = "INVALID_DESIGN_TRANSITION"
                errorMessage = "Only an existing, non-obsolete design version can be obsoleted."
                Exit Function
            End If
        Case Else
            errorCode = "INVALID_DESIGN_EVENT_TYPE"
            errorMessage = "Unsupported design event type: " & eventType
            Exit Function
    End Select
    ValidateLifecycleTransition = True
End Function

Private Function ReplayDesignEvent(ByVal loEvents As ListObject, ByVal values As Variant, _
                                   ByVal rowIndex As Long, ByVal loDesigns As ListObject, _
                                   ByVal loLines As ListObject, ByRef errorMessage As String) As Boolean
    Dim eventType As String
    Dim designId As String
    Dim designVersion As String
    Dim userId As String
    Dim eventId As String
    Dim appliedAt As Variant
    Dim payload As Collection
    Dim item As Variant
    Dim lr As ListRow
    Dim targetRow As Long

    eventType = UCase$(ReadDesignTableText(loEvents, values, rowIndex, "EventType"))
    designId = ReadDesignTableText(loEvents, values, rowIndex, "DesignId")
    designVersion = ReadDesignTableText(loEvents, values, rowIndex, "DesignVersion")
    userId = ReadDesignTableText(loEvents, values, rowIndex, "UserId")
    eventId = ReadDesignTableText(loEvents, values, rowIndex, "EventID")
    appliedAt = ReadDesignTableValue(loEvents, values, rowIndex, "AppliedAtUTC")

    Select Case eventType
        Case DESIGN_EVENT_CREATE
            Set payload = ParseDesignPayload(ReadDesignTableText(loEvents, values, rowIndex, "PayloadJson"), errorMessage)
            If payload Is Nothing Then Exit Function
            If payload.Count = 0 Then Exit Function
            Set lr = loDesigns.ListRows.Add
            SetDesignTableValue loDesigns, lr.Index, "DesignId", designId
            SetDesignTableValue loDesigns, lr.Index, "DesignVersion", designVersion
            SetDesignTableValue loDesigns, lr.Index, "DesignType", PayloadText(payload(1), "DesignType")
            SetDesignTableValue loDesigns, lr.Index, "DesignName", PayloadText(payload(1), "DesignName")
            SetDesignTableValue loDesigns, lr.Index, "Description", PayloadText(payload(1), "Description")
            SetDesignTableValue loDesigns, lr.Index, "Status", "DRAFT"
            SetDesignTableValue loDesigns, lr.Index, "EffectiveFromUTC", PayloadText(payload(1), "EffectiveFromUTC")
            SetDesignTableValue loDesigns, lr.Index, "EffectiveToUTC", PayloadText(payload(1), "EffectiveToUTC")
            SetDesignTableValue loDesigns, lr.Index, "CreatedAtUTC", ReadDesignTableValue(loEvents, values, rowIndex, "OccurredAtUTC")
            SetDesignTableValue loDesigns, lr.Index, "CreatedByUserId", userId
            SetDesignTableValue loDesigns, lr.Index, "SourceEventID", eventId
            For Each item In payload
                If PayloadHasLine(item) Then AddProjectedDesignLine loLines, designId, designVersion, item
            Next item
        Case DESIGN_EVENT_RELEASE
            targetRow = FindProjectedDesignRow(loDesigns, designId, designVersion)
            If targetRow = 0 Then
                errorMessage = "Release target not found."
                Exit Function
            End If
            SetDesignTableValue loDesigns, targetRow, "Status", "RELEASED"
            SetDesignTableValue loDesigns, targetRow, "ReleasedAtUTC", appliedAt
            SetDesignTableValue loDesigns, targetRow, "ReleasedByUserId", userId
        Case DESIGN_EVENT_OBSOLETE
            targetRow = FindProjectedDesignRow(loDesigns, designId, designVersion)
            If targetRow = 0 Then
                errorMessage = "Obsolete target not found."
                Exit Function
            End If
            SetDesignTableValue loDesigns, targetRow, "Status", "OBSOLETE"
            SetDesignTableValue loDesigns, targetRow, "ObsoletedAtUTC", appliedAt
            SetDesignTableValue loDesigns, targetRow, "ObsoletedByUserId", userId
        Case Else
            errorMessage = "Unsupported event type " & eventType
            Exit Function
    End Select
    ReplayDesignEvent = True
End Function

Private Sub AddProjectedDesignLine(ByVal lo As ListObject, ByVal designId As String, _
                                   ByVal designVersion As String, ByVal item As Object)
    Dim lr As ListRow
    Set lr = lo.ListRows.Add
    SetDesignTableValue lo, lr.Index, "DesignId", designId
    SetDesignTableValue lo, lr.Index, "DesignVersion", designVersion
    SetDesignTableValue lo, lr.Index, "LineNo", PayloadValue(item, "LineNo")
    SetDesignTableValue lo, lr.Index, "Process", PayloadText(item, "Process")
    SetDesignTableValue lo, lr.Index, "IOType", PayloadText(item, "IOType")
    SetDesignTableValue lo, lr.Index, "ComponentSKU", PayloadText(item, "ComponentSKU")
    SetDesignTableValue lo, lr.Index, "ComponentDesignId", PayloadText(item, "ComponentDesignId")
    SetDesignTableValue lo, lr.Index, "ComponentDesignVersion", PayloadText(item, "ComponentDesignVersion")
    SetDesignTableValue lo, lr.Index, "Qty", PayloadValue(item, "Qty")
    SetDesignTableValue lo, lr.Index, "UOM", PayloadText(item, "UOM")
    SetDesignTableValue lo, lr.Index, "Percent", PayloadValue(item, "Percent")
    SetDesignTableValue lo, lr.Index, "Instruction", PayloadText(item, "Instruction")
End Sub

Private Function PayloadHasLine(ByVal item As Object) As Boolean
    PayloadHasLine = (PayloadText(item, "ComponentSKU") <> "" _
                      Or PayloadText(item, "ComponentDesignId") <> "" _
                      Or PayloadText(item, "IOType") <> "")
End Function

Private Function CurrentDesignStatus(ByVal wb As Workbook, ByVal designId As String, _
                                     ByVal designVersion As String) As String
    Dim lo As ListObject
    Dim values As Variant
    Dim r As Long
    Set lo = FindDesignsApplyTable(wb, "tblDesigns")
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    values = lo.DataBodyRange.Value
    For r = 1 To UBound(values, 1)
        If StrComp(ReadDesignTableText(lo, values, r, "DesignId"), designId, vbTextCompare) = 0 _
           And StrComp(ReadDesignTableText(lo, values, r, "DesignVersion"), designVersion, vbTextCompare) = 0 Then
            CurrentDesignStatus = ReadDesignTableText(lo, values, r, "Status")
            Exit Function
        End If
    Next r
End Function

Private Function FindProjectedDesignRow(ByVal lo As ListObject, ByVal designId As String, _
                                        ByVal designVersion As String) As Long
    Dim values As Variant
    Dim r As Long
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    values = lo.DataBodyRange.Value
    For r = 1 To UBound(values, 1)
        If StrComp(ReadDesignTableText(lo, values, r, "DesignId"), designId, vbTextCompare) = 0 _
           And StrComp(ReadDesignTableText(lo, values, r, "DesignVersion"), designVersion, vbTextCompare) = 0 Then
            FindProjectedDesignRow = r
            Exit Function
        End If
    Next r
End Function

Private Function AppliedDesignEventExists(ByVal lo As ListObject, ByVal eventId As String) As Boolean
    Dim values As Variant
    Dim r As Long
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    values = lo.DataBodyRange.Value
    For r = 1 To UBound(values, 1)
        If StrComp(ReadDesignTableText(lo, values, r, "EventID"), eventId, vbTextCompare) = 0 Then
            AppliedDesignEventExists = True
            Exit Function
        End If
    Next r
End Function

Private Function NextDesignAppliedSeq(ByVal lo As ListObject) As Long
    Dim values As Variant
    Dim r As Long
    Dim seq As Long
    If lo Is Nothing Then
        NextDesignAppliedSeq = 1
        Exit Function
    End If
    If lo.DataBodyRange Is Nothing Then
        NextDesignAppliedSeq = 1
        Exit Function
    End If
    values = lo.DataBodyRange.Value
    For r = 1 To UBound(values, 1)
        If IsNumeric(ReadDesignTableValue(lo, values, r, "AppliedSeq")) Then
            If CLng(ReadDesignTableValue(lo, values, r, "AppliedSeq")) > seq Then _
                seq = CLng(ReadDesignTableValue(lo, values, r, "AppliedSeq"))
        End If
    Next r
    NextDesignAppliedSeq = seq + 1
End Function

Private Sub ClearDesignTableRows(ByVal lo As ListObject)
    If lo Is Nothing Then Exit Sub
    Do While lo.ListRows.Count > 0
        lo.ListRows(lo.ListRows.Count).Delete
    Loop
End Sub

Private Sub SaveDesignsWorkbook(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub
    If wb.ReadOnly Then Exit Sub
    If Trim$(wb.Path) = "" Then Exit Sub
    wb.Save
End Sub

Private Function FindDesignsApplyTable(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        On Error Resume Next
        Set FindDesignsApplyTable = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not FindDesignsApplyTable Is Nothing Then Exit Function
    Next ws
End Function

Private Sub SetDesignTableValue(ByVal lo As ListObject, ByVal rowIndex As Long, _
                                ByVal columnName As String, ByVal valueOut As Variant)
    lo.DataBodyRange.Cells(rowIndex, lo.ListColumns(columnName).Index).Value = valueOut
End Sub

Private Function ReadDesignTableValue(ByVal lo As ListObject, ByVal values As Variant, _
                                      ByVal rowIndex As Long, ByVal columnName As String) As Variant
    ReadDesignTableValue = values(rowIndex, lo.ListColumns(columnName).Index)
End Function

Private Function ReadDesignTableText(ByVal lo As ListObject, ByVal values As Variant, _
                                     ByVal rowIndex As Long, ByVal columnName As String) As String
    Dim valueIn As Variant
    valueIn = ReadDesignTableValue(lo, values, rowIndex, columnName)
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Then Exit Function
    ReadDesignTableText = Trim$(CStr(valueIn))
End Function

Private Function GetEventTextDesign(ByVal evt As Object, ByVal key As String) As String
    On Error GoTo CleanExit
    If evt Is Nothing Then Exit Function
    If Not evt.Exists(key) Then Exit Function
    If IsNull(evt(key)) Or IsEmpty(evt(key)) Or IsError(evt(key)) Then Exit Function
    GetEventTextDesign = Trim$(CStr(evt(key)))
CleanExit:
End Function

Private Function TryGetEventDateDesign(ByVal evt As Object, ByVal key As String, _
                                       ByRef dateOut As Date) As Boolean
    Dim textIn As String
    On Error GoTo CleanExit
    textIn = GetEventTextDesign(evt, key)
    If textIn = "" Or Not IsDate(textIn) Then Exit Function
    dateOut = CDate(textIn)
    TryGetEventDateDesign = True
CleanExit:
End Function

Private Function PayloadText(ByVal item As Object, ByVal key As String) As String
    Dim valueIn As Variant
    On Error GoTo CleanExit
    If item Is Nothing Then Exit Function
    If Not item.Exists(key) Then Exit Function
    valueIn = item(key)
    If IsNull(valueIn) Or IsEmpty(valueIn) Or IsError(valueIn) Then Exit Function
    PayloadText = Trim$(CStr(valueIn))
CleanExit:
End Function

Private Function PayloadValue(ByVal item As Object, ByVal key As String) As Variant
    On Error GoTo CleanExit
    If item Is Nothing Then Exit Function
    If Not item.Exists(key) Then Exit Function
    PayloadValue = item(key)
CleanExit:
End Function

Private Function ParseDesignPayload(ByVal jsonText As String, ByRef errorMessage As String) As Collection
    Dim pos As Long
    Dim item As Object
    Set ParseDesignPayload = New Collection
    pos = 1
    SkipJsonWhitespaceDesign jsonText, pos
    If Mid$(jsonText, pos, 1) <> "[" Then
        errorMessage = "PayloadJson must start with '['."
        Set ParseDesignPayload = Nothing
        Exit Function
    End If
    pos = pos + 1
    Do
        SkipJsonWhitespaceDesign jsonText, pos
        If Mid$(jsonText, pos, 1) = "]" Then
            pos = pos + 1
            Exit Do
        End If
        Set item = ParseJsonObjectDesign(jsonText, pos, errorMessage)
        If item Is Nothing Then
            Set ParseDesignPayload = Nothing
            Exit Function
        End If
        ParseDesignPayload.Add item
        SkipJsonWhitespaceDesign jsonText, pos
        If Mid$(jsonText, pos, 1) = "," Then
            pos = pos + 1
        ElseIf Mid$(jsonText, pos, 1) <> "]" Then
            errorMessage = "PayloadJson array is missing a comma separator."
            Set ParseDesignPayload = Nothing
            Exit Function
        End If
    Loop
    SkipJsonWhitespaceDesign jsonText, pos
    If pos <= Len(jsonText) Then
        errorMessage = "PayloadJson contains unexpected trailing characters."
        Set ParseDesignPayload = Nothing
    End If
End Function

Private Function ParseJsonObjectDesign(ByVal jsonText As String, ByRef pos As Long, _
                                       ByRef errorMessage As String) As Object
    Dim item As Object
    Dim key As String
    Set item = CreateObject("Scripting.Dictionary")
    item.CompareMode = vbTextCompare
    SkipJsonWhitespaceDesign jsonText, pos
    If Mid$(jsonText, pos, 1) <> "{" Then
        errorMessage = "PayloadJson object must start with '{'."
        Exit Function
    End If
    pos = pos + 1
    Do
        SkipJsonWhitespaceDesign jsonText, pos
        If Mid$(jsonText, pos, 1) = "}" Then
            pos = pos + 1
            Set ParseJsonObjectDesign = item
            Exit Function
        End If
        key = ParseJsonStringDesign(jsonText, pos, errorMessage)
        If errorMessage <> "" Then Exit Function
        SkipJsonWhitespaceDesign jsonText, pos
        If Mid$(jsonText, pos, 1) <> ":" Then
            errorMessage = "PayloadJson object is missing ':' after key '" & key & "'."
            Exit Function
        End If
        pos = pos + 1
        item(key) = ParseJsonValueDesign(jsonText, pos, errorMessage)
        If errorMessage <> "" Then Exit Function
        SkipJsonWhitespaceDesign jsonText, pos
        If Mid$(jsonText, pos, 1) = "," Then
            pos = pos + 1
        ElseIf Mid$(jsonText, pos, 1) = "}" Then
            pos = pos + 1
            Set ParseJsonObjectDesign = item
            Exit Function
        Else
            errorMessage = "PayloadJson object is missing a comma separator."
            Exit Function
        End If
    Loop
End Function

Private Function ParseJsonValueDesign(ByVal jsonText As String, ByRef pos As Long, _
                                      ByRef errorMessage As String) As Variant
    Dim token As String
    Dim startAt As Long
    SkipJsonWhitespaceDesign jsonText, pos
    If Mid$(jsonText, pos, 1) = Chr$(34) Then
        ParseJsonValueDesign = ParseJsonStringDesign(jsonText, pos, errorMessage)
        Exit Function
    End If
    startAt = pos
    Do While pos <= Len(jsonText) And InStr(1, ",}]" & vbCr & vbLf & vbTab & " ", Mid$(jsonText, pos, 1), vbBinaryCompare) = 0
        pos = pos + 1
    Loop
    token = Trim$(Mid$(jsonText, startAt, pos - startAt))
    Select Case LCase$(token)
        Case "true": ParseJsonValueDesign = True
        Case "false": ParseJsonValueDesign = False
        Case "null": ParseJsonValueDesign = vbNullString
        Case Else
            If IsNumeric(token) Then
                ParseJsonValueDesign = CDbl(token)
            Else
                errorMessage = "Unsupported value in PayloadJson at position " & CStr(startAt) & "."
            End If
    End Select
End Function

Private Function ParseJsonStringDesign(ByVal jsonText As String, ByRef pos As Long, _
                                       ByRef errorMessage As String) As String
    Dim ch As String
    Dim esc As String
    If Mid$(jsonText, pos, 1) <> Chr$(34) Then
        errorMessage = "Expected string value in PayloadJson."
        Exit Function
    End If
    pos = pos + 1
    Do While pos <= Len(jsonText)
        ch = Mid$(jsonText, pos, 1)
        pos = pos + 1
        If ch = Chr$(34) Then Exit Function
        If ch = "\" Then
            If pos > Len(jsonText) Then
                errorMessage = "Incomplete escape sequence in PayloadJson."
                Exit Function
            End If
            esc = Mid$(jsonText, pos, 1)
            pos = pos + 1
            Select Case esc
                Case Chr$(34), "\", "/": ParseJsonStringDesign = ParseJsonStringDesign & esc
                Case "n": ParseJsonStringDesign = ParseJsonStringDesign & vbLf
                Case "r": ParseJsonStringDesign = ParseJsonStringDesign & vbCr
                Case "t": ParseJsonStringDesign = ParseJsonStringDesign & vbTab
                Case Else
                    errorMessage = "Unsupported escape sequence in PayloadJson."
                    Exit Function
            End Select
        Else
            ParseJsonStringDesign = ParseJsonStringDesign & ch
        End If
    Loop
    errorMessage = "Unterminated string in PayloadJson."
End Function

Private Sub SkipJsonWhitespaceDesign(ByVal jsonText As String, ByRef pos As Long)
    Do While pos <= Len(jsonText)
        If InStr(1, " " & vbTab & vbCr & vbLf, Mid$(jsonText, pos, 1), vbBinaryCompare) = 0 Then Exit Do
        pos = pos + 1
    Loop
End Sub
