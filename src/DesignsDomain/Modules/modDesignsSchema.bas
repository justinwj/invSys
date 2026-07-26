Attribute VB_Name = "modDesignsSchema"
Option Explicit

Private Const SHEET_DESIGNS As String = "Designs"
Private Const SHEET_LINES As String = "DesignLines"
Private Const SHEET_EVENTS As String = "DesignEvents"
Private Const SHEET_APPLIED As String = "AppliedDesignEvents"
Private Const SHEET_LOCKS As String = "Locks"

Private Const TABLE_DESIGNS As String = "tblDesigns"
Private Const TABLE_LINES As String = "tblDesignLines"
Private Const TABLE_EVENTS As String = "tblDesignEvents"
Private Const TABLE_APPLIED As String = "tblAppliedDesignEvents"
Private Const TABLE_LOCKS As String = "tblLocks"

Public Function EnsureDesignsSchema(Optional ByVal targetWb As Workbook = Nothing, _
                                    Optional ByRef report As String = "") As Boolean
    On Error GoTo FailEnsure

    Dim wb As Workbook
    Set wb = targetWb
    If wb Is Nothing Then
        report = "Authoritative Designs workbook was not supplied."
        Exit Function
    End If
    If wb.IsAddin Then
        report = "Designs schema cannot be created inside an XLAM."
        Exit Function
    End If

    EnsureDesignsTable wb
    EnsureDesignLinesTable wb
    EnsureDesignEventsTable wb
    EnsureAppliedDesignEventsTable wb
    EnsureDesignLocksTable wb
    report = "OK"
    EnsureDesignsSchema = True
    Exit Function

FailEnsure:
    report = "EnsureDesignsSchema failed: " & Err.Description
End Function

Public Function ValidateDesignsSchema(ByVal targetWb As Workbook) As String
    If targetWb Is Nothing Then
        ValidateDesignsSchema = "Authoritative Designs workbook was not supplied."
        Exit Function
    End If

    ValidateDesignsSchema = ValidateRequiredTable(targetWb, TABLE_DESIGNS, _
        Array("DesignId", "DesignVersion", "DesignType", "DesignName", "Status", "CreatedAtUTC", "CreatedByUserId"))
    If ValidateDesignsSchema <> "" Then Exit Function
    ValidateDesignsSchema = ValidateRequiredTable(targetWb, TABLE_LINES, _
        Array("DesignId", "DesignVersion", "LineNo", "IOType", "ComponentSKU", "Qty", "UOM"))
    If ValidateDesignsSchema <> "" Then Exit Function
    ValidateDesignsSchema = ValidateRequiredTable(targetWb, TABLE_EVENTS, _
        Array("EventID", "AppliedSeq", "EventType", "WarehouseId", "DesignId", "DesignVersion", "PayloadJson"))
    If ValidateDesignsSchema <> "" Then Exit Function
    ValidateDesignsSchema = ValidateRequiredTable(targetWb, TABLE_APPLIED, _
        Array("EventID", "AppliedSeq", "AppliedAtUTC", "RunId", "SourceInbox", "Status"))
    If ValidateDesignsSchema <> "" Then Exit Function
    ValidateDesignsSchema = ValidateRequiredTable(targetWb, TABLE_LOCKS, _
        Array("LockName", "OwnerStationId", "OwnerUserId", "RunId", "AcquiredAtUTC", "ExpiresAtUTC", "HeartbeatAtUTC", "Status"))
End Function

Private Sub EnsureDesignsTable(ByVal wb As Workbook)
    EnsureTable wb, SHEET_DESIGNS, TABLE_DESIGNS, Array( _
        "DesignId", "DesignVersion", "DesignType", "DesignName", "Description", "Status", _
        "EffectiveFromUTC", "EffectiveToUTC", "CreatedAtUTC", "CreatedByUserId", _
        "ReleasedAtUTC", "ReleasedByUserId", "ObsoletedAtUTC", "ObsoletedByUserId", "SourceEventID")
End Sub

Private Sub EnsureDesignLinesTable(ByVal wb As Workbook)
    EnsureTable wb, SHEET_LINES, TABLE_LINES, Array( _
        "DesignId", "DesignVersion", "LineNo", "Process", "IOType", "ComponentSKU", _
        "ComponentDesignId", "ComponentDesignVersion", "Qty", "UOM", "Percent", "Instruction")
End Sub

Private Sub EnsureDesignEventsTable(ByVal wb As Workbook)
    EnsureTable wb, SHEET_EVENTS, TABLE_EVENTS, Array( _
        "EventID", "UndoOfEventId", "AppliedSeq", "EventType", "OccurredAtUTC", "AppliedAtUTC", _
        "WarehouseId", "StationId", "UserId", "DesignId", "DesignVersion", "PayloadJson", "Note")
End Sub

Private Sub EnsureAppliedDesignEventsTable(ByVal wb As Workbook)
    EnsureTable wb, SHEET_APPLIED, TABLE_APPLIED, Array( _
        "EventID", "UndoOfEventId", "AppliedSeq", "AppliedAtUTC", "RunId", "SourceInbox", "Status")
End Sub

Private Sub EnsureDesignLocksTable(ByVal wb As Workbook)
    EnsureTable wb, SHEET_LOCKS, TABLE_LOCKS, Array( _
        "LockName", "OwnerStationId", "OwnerUserId", "RunId", "AcquiredAtUTC", _
        "ExpiresAtUTC", "HeartbeatAtUTC", "Status")
End Sub

Private Sub EnsureTable(ByVal wb As Workbook, ByVal sheetName As String, _
                        ByVal tableName As String, ByVal headers As Variant)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim i As Long
    Dim nextColumn As Long
    Dim dataRange As Range

    Set ws = EnsureWorksheet(wb, sheetName)
    Set lo = FindTable(wb, tableName)
    If lo Is Nothing Then
        For i = LBound(headers) To UBound(headers)
            ws.Cells(1, i - LBound(headers) + 1).Value = CStr(headers(i))
        Next i
        Set dataRange = ws.Range(ws.Cells(1, 1), ws.Cells(2, UBound(headers) - LBound(headers) + 1))
        Set lo = ws.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
        lo.Name = tableName
        If Not lo.DataBodyRange Is Nothing Then lo.ListRows(1).Delete
    Else
        For i = LBound(headers) To UBound(headers)
            If ColumnIndex(lo, CStr(headers(i))) = 0 Then
                nextColumn = lo.ListColumns.Count + 1
                lo.ListColumns.Add nextColumn
                lo.ListColumns(nextColumn).Name = CStr(headers(i))
            End If
        Next i
    End If
    FormatDesignIdentityColumns lo
End Sub

Private Sub FormatDesignIdentityColumns(ByVal lo As ListObject)
    Dim columnName As String
    Dim lc As ListColumn

    If lo Is Nothing Then Exit Sub
    For Each lc In lo.ListColumns
        columnName = UCase$(Trim$(lc.Name))
        Select Case columnName
            Case "EVENTID", "UNDOOFEVENTID", "WAREHOUSEID", "STATIONID", "USERID", _
                 "DESIGNID", "DESIGNVERSION", "COMPONENTSKU", "COMPONENTDESIGNID", _
                 "COMPONENTDESIGNVERSION", "SOURCEEVENTID", "RUNID", "OWNERSTATIONID", _
                 "OWNERUSERID"
                lc.Range.NumberFormat = "@"
        End Select
    Next lc
End Sub

Private Function EnsureWorksheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureWorksheet = wb.Worksheets(sheetName)
    On Error GoTo 0
    If EnsureWorksheet Is Nothing Then
        Set EnsureWorksheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        EnsureWorksheet.Name = sheetName
    End If
End Function

Private Function FindTable(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        On Error Resume Next
        Set FindTable = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not FindTable Is Nothing Then Exit Function
    Next ws
End Function

Private Function ColumnIndex(ByVal lo As ListObject, ByVal columnName As String) As Long
    Dim i As Long
    If lo Is Nothing Then Exit Function
    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).Name, columnName, vbTextCompare) = 0 Then
            ColumnIndex = i
            Exit Function
        End If
    Next i
End Function

Private Function ValidateRequiredTable(ByVal wb As Workbook, ByVal tableName As String, _
                                       ByVal headers As Variant) As String
    Dim lo As ListObject
    Dim i As Long

    Set lo = FindTable(wb, tableName)
    If lo Is Nothing Then
        ValidateRequiredTable = tableName & " not found."
        Exit Function
    End If
    For i = LBound(headers) To UBound(headers)
        If ColumnIndex(lo, CStr(headers(i))) = 0 Then
            ValidateRequiredTable = tableName & " missing column " & CStr(headers(i)) & "."
            Exit Function
        End If
    Next i
End Function
