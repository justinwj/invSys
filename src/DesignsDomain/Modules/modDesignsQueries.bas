Attribute VB_Name = "modDesignsQueries"
Option Explicit

Public Function ListDesigns(Optional ByVal designsWb As Workbook = Nothing, _
                            Optional ByVal statusFilter As String = "") As Variant
    On Error GoTo FailQuery

    Dim wb As Workbook
    Dim lo As ListObject
    Dim src As Variant
    Dim result() As Variant
    Dim trimmed() As Variant
    Dim r As Long
    Dim c As Long
    Dim outRow As Long
    Dim report As String

    Set wb = modDesignsRuntime.ResolveDesignsWorkbook("", designsWb, report)
    If wb Is Nothing Then Exit Function
    Set lo = FindDesignsTableQuery(wb, "tblDesigns")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function

    src = lo.DataBodyRange.Value
    ReDim result(1 To UBound(src, 1), 1 To 6)
    For r = 1 To UBound(src, 1)
        If statusFilter = "" Or StrComp(ReadDesignCellQuery(lo, src, r, "Status"), statusFilter, vbTextCompare) = 0 Then
            outRow = outRow + 1
            result(outRow, 1) = ReadDesignCellQuery(lo, src, r, "DesignId")
            result(outRow, 2) = ReadDesignCellQuery(lo, src, r, "DesignVersion")
            result(outRow, 3) = ReadDesignCellQuery(lo, src, r, "DesignType")
            result(outRow, 4) = ReadDesignCellQuery(lo, src, r, "DesignName")
            result(outRow, 5) = ReadDesignCellQuery(lo, src, r, "Description")
            result(outRow, 6) = ReadDesignCellQuery(lo, src, r, "Status")
        End If
    Next r
    If outRow = 0 Then Exit Function
    ReDim trimmed(1 To outRow, 1 To 6)
    For r = 1 To outRow
        For c = 1 To 6
            trimmed(r, c) = result(r, c)
        Next c
    Next r
    ListDesigns = trimmed
    Exit Function

FailQuery:
    ListDesigns = Empty
End Function

Public Function GetBOM(ByVal designId As String, ByVal designVersion As String, _
                       Optional ByVal designsWb As Workbook = Nothing) As Variant
    On Error GoTo FailQuery

    Dim wb As Workbook
    Dim lo As ListObject
    Dim src As Variant
    Dim result() As Variant
    Dim trimmed() As Variant
    Dim r As Long
    Dim c As Long
    Dim outRow As Long
    Dim report As String

    Set wb = modDesignsRuntime.ResolveDesignsWorkbook("", designsWb, report)
    If wb Is Nothing Then Exit Function
    Set lo = FindDesignsTableQuery(wb, "tblDesignLines")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    src = lo.DataBodyRange.Value
    ReDim result(1 To UBound(src, 1), 1 To 10)
    For r = 1 To UBound(src, 1)
        If DesignIdsMatchQuery(ReadDesignCellQuery(lo, src, r, "DesignId"), designId) _
           And StrComp(ReadDesignCellQuery(lo, src, r, "DesignVersion"), Trim$(designVersion), vbTextCompare) = 0 Then
            outRow = outRow + 1
            result(outRow, 1) = ReadDesignCellQuery(lo, src, r, "LineNo")
            result(outRow, 2) = ReadDesignCellQuery(lo, src, r, "Process")
            result(outRow, 3) = ReadDesignCellQuery(lo, src, r, "IOType")
            result(outRow, 4) = ReadDesignCellQuery(lo, src, r, "ComponentSKU")
            result(outRow, 5) = ReadDesignCellQuery(lo, src, r, "ComponentDesignId")
            result(outRow, 6) = ReadDesignCellQuery(lo, src, r, "ComponentDesignVersion")
            result(outRow, 7) = ReadDesignCellQuery(lo, src, r, "Qty")
            result(outRow, 8) = ReadDesignCellQuery(lo, src, r, "UOM")
            result(outRow, 9) = ReadDesignCellQuery(lo, src, r, "Percent")
            result(outRow, 10) = ReadDesignCellQuery(lo, src, r, "Instruction")
        End If
    Next r
    If outRow = 0 Then Exit Function
    SortBomRowsByLineNo result, outRow
    ReDim trimmed(1 To outRow, 1 To 10)
    For r = 1 To outRow
        For c = 1 To 10
            trimmed(r, c) = result(r, c)
        Next c
    Next r
    GetBOM = trimmed
    Exit Function

FailQuery:
    GetBOM = Empty
End Function

Private Sub SortBomRowsByLineNo(ByRef values As Variant, ByVal rowCount As Long)
    Dim i As Long
    Dim j As Long
    Dim c As Long
    Dim leftLine As Double
    Dim rightLine As Double
    Dim swapValue As Variant

    For i = 1 To rowCount - 1
        For j = i + 1 To rowCount
            leftLine = DesignLineSortValueQuery(values(i, 1), i)
            rightLine = DesignLineSortValueQuery(values(j, 1), j)
            If rightLine < leftLine Then
                For c = 1 To 10
                    swapValue = values(i, c)
                    values(i, c) = values(j, c)
                    values(j, c) = swapValue
                Next c
            End If
        Next j
    Next i
End Sub

Private Function DesignLineSortValueQuery(ByVal lineValue As Variant, ByVal fallbackRow As Long) As Double
    If IsNumeric(lineValue) Then
        DesignLineSortValueQuery = CDbl(lineValue)
    Else
        DesignLineSortValueQuery = 1000000000# + fallbackRow
    End If
End Function

Public Function GetBOMForStatus(ByVal designId As String, _
                                ByVal designVersion As String, _
                                ByVal requiredStatus As String, _
                                Optional ByVal designsWb As Workbook = Nothing) As Variant
    On Error GoTo FailQuery

    Dim wb As Workbook
    Dim lo As ListObject
    Dim src As Variant
    Dim r As Long
    Dim report As String

    designId = Trim$(designId)
    designVersion = Trim$(designVersion)
    requiredStatus = UCase$(Trim$(requiredStatus))
    If designId = "" Or designVersion = "" Or requiredStatus = "" Then Exit Function

    Set wb = modDesignsRuntime.ResolveDesignsWorkbook("", designsWb, report)
    If wb Is Nothing Then Exit Function
    Set lo = FindDesignsTableQuery(wb, "tblDesigns")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function

    src = lo.DataBodyRange.Value
    For r = 1 To UBound(src, 1)
        If DesignIdsMatchQuery(ReadDesignCellQuery(lo, src, r, "DesignId"), designId) _
           And StrComp(ReadDesignCellQuery(lo, src, r, "DesignVersion"), designVersion, vbTextCompare) = 0 Then
            If StrComp(ReadDesignCellQuery(lo, src, r, "Status"), requiredStatus, vbTextCompare) <> 0 Then Exit Function
            GetBOMForStatus = GetBOM(designId, designVersion, wb)
            Exit Function
        End If
    Next r
    Exit Function

FailQuery:
    GetBOMForStatus = Empty
End Function

Private Function DesignIdsMatchQuery(ByVal leftId As String, ByVal rightId As String) As Boolean
    DesignIdsMatchQuery = (StrComp(CanonicalDesignIdQuery(leftId), _
                                   CanonicalDesignIdQuery(rightId), vbTextCompare) = 0)
End Function

Private Function CanonicalDesignIdQuery(ByVal valueIn As String) As String
    Dim textValue As String
    Dim numericValue As Long

    textValue = UCase$(Trim$(valueIn))
    If textValue = "" Then Exit Function
    If Len(textValue) <= 3 And IsNumeric(textValue) Then
        numericValue = CLng(CDbl(textValue))
        If numericValue >= 0 And numericValue <= 999 Then
            CanonicalDesignIdQuery = Right$("000" & CStr(numericValue), 3)
            Exit Function
        End If
    End If
    CanonicalDesignIdQuery = textValue
End Function

Private Function FindDesignsTableQuery(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        On Error Resume Next
        Set FindDesignsTableQuery = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not FindDesignsTableQuery Is Nothing Then Exit Function
    Next ws
End Function

Private Function ReadDesignCellQuery(ByVal lo As ListObject, ByVal values As Variant, _
                                     ByVal rowIndex As Long, ByVal columnName As String) As String
    Dim columnIndex As Long
    On Error Resume Next
    columnIndex = lo.ListColumns(columnName).Index
    On Error GoTo 0
    If columnIndex = 0 Then Exit Function
    If IsError(values(rowIndex, columnIndex)) Or IsNull(values(rowIndex, columnIndex)) Then Exit Function
    ReadDesignCellQuery = CStr(values(rowIndex, columnIndex))
End Function
