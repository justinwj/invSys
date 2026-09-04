Attribute VB_Name = "modListBoxTableExport"
Option Explicit

Public Function ExportVisibleListBoxToNewTable(ByVal sourceList As MSForms.ListBox, _
                                               ByVal headers As Variant, _
                                               ByRef report As String) As Boolean
    Dim exportWb As Workbook
    Dim exportWs As Worksheet
    Dim exportTable As ListObject
    Dim visibleColumns As Collection
    Dim widths As Variant
    Dim sourceColumn As Long
    Dim targetColumn As Long
    Dim rowIndex As Long
    Dim tableRange As Range
    Dim headerText As String

    If sourceList Is Nothing Then
        report = "The named ListBox is not open."
        Exit Function
    End If
    If Not IsArray(headers) Then
        report = "The named ListBox has no declared export headings."
        Exit Function
    End If
    widths = Split(sourceList.ColumnWidths, ";")
    If UBound(headers) - LBound(headers) + 1 <> sourceList.ColumnCount Then
        report = "The named ListBox export schema does not match its displayed columns."
        Exit Function
    End If
    Set visibleColumns = New Collection
    For sourceColumn = 0 To sourceList.ColumnCount - 1
        If sourceColumn <= UBound(widths) Then
            If Val(Replace$(Trim$(CStr(widths(sourceColumn))), "pt", "")) > 0 Then
                headerText = Trim$(CStr(headers(LBound(headers) + sourceColumn)))
                If headerText <> "" Then visibleColumns.Add sourceColumn
            End If
        End If
    Next sourceColumn
    If visibleColumns.Count = 0 Then
        report = "The named ListBox has no visible exportable columns."
        Exit Function
    End If

    Set exportWb = Application.Workbooks.Add(xlWBATWorksheet)
    Set exportWs = exportWb.Worksheets(1)
    exportWs.Name = "ListBox Export"
    For targetColumn = 1 To visibleColumns.Count
        sourceColumn = CLng(visibleColumns(targetColumn))
        exportWs.Cells(1, targetColumn).Value = headers(LBound(headers) + sourceColumn)
    Next targetColumn
    For rowIndex = 0 To sourceList.ListCount - 1
        For targetColumn = 1 To visibleColumns.Count
            sourceColumn = CLng(visibleColumns(targetColumn))
            exportWs.Cells(rowIndex + 2, targetColumn).Value = sourceList.List(rowIndex, sourceColumn)
        Next targetColumn
    Next rowIndex
    Set tableRange = exportWs.Range(exportWs.Cells(1, 1), _
        exportWs.Cells(Application.Max(2, sourceList.ListCount + 1), visibleColumns.Count))
    Set exportTable = exportWs.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
    exportTable.Name = "tblListBoxExport"
    exportWs.Columns.AutoFit
    exportWs.Activate
    report = "Exported " & CStr(sourceList.ListCount) & " row(s) from " & sourceList.Name & _
             " to " & exportTable.Name & " in a new unsaved workbook."
    ExportVisibleListBoxToNewTable = True
End Function
