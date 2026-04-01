Attribute VB_Name = "modHqAggregator"
Option Explicit

Private Const SHEET_GLOBAL_SNAPSHOT As String = "GlobalInventorySnapshot"
Private Const TABLE_GLOBAL_SNAPSHOT As String = "tblGlobalInventorySnapshot"
Private Const TABLE_WAREHOUSE_SNAPSHOT As String = "tblInventorySnapshot"

Public Function RunHQAggregation(Optional ByVal sharePointRoot As String = "", _
                                 Optional ByVal outputPath As String = "", _
                                 Optional ByRef report As String = "") As Boolean
    Dim snapshotsFolder As String

    If sharePointRoot = "" Then
        If Not modConfig.LoadConfig("", "") Then
            report = "Config load failed: " & modConfig.Validate()
            Exit Function
        End If
        sharePointRoot = modConfig.GetString("PathSharePointRoot", "")
    End If
    If Trim$(sharePointRoot) = "" Then
        report = "PathSharePointRoot not configured."
        Exit Function
    End If

    snapshotsFolder = NormalizeFolderPathHq(sharePointRoot) & "Snapshots"
    If Trim$(outputPath) = "" Then outputPath = NormalizeFolderPathHq(sharePointRoot) & "Global\invSys.Global.InventorySnapshot.xlsb"
    RunHQAggregation = GenerateGlobalSnapshotFromFolder(snapshotsFolder, outputPath, report)
End Function

Public Function GenerateGlobalSnapshotFromFolder(ByVal snapshotsFolder As String, _
                                                 ByVal outputPath As String, _
                                                 Optional ByRef report As String = "") As Boolean
    On Error GoTo FailAggregate

    Dim fileName As String
    Dim tempFolder As String
    Dim tempFile As String
    Dim wbSnap As Workbook
    Dim globalRows As Object
    Dim key As String
    Dim lo As ListObject
    Dim i As Long

    If Trim$(snapshotsFolder) = "" Then
        report = "Snapshots folder is required."
        Exit Function
    End If

    Set globalRows = CreateObject("Scripting.Dictionary")
    globalRows.CompareMode = vbTextCompare
    tempFolder = Environ$("TEMP") & "\invSysHQ_" & Format$(Now, "yyyymmdd_hhnnss")
    CreateFolderRecursiveHq tempFolder

    fileName = Dir$(NormalizeFolderPathHq(snapshotsFolder) & "*.invSys.Snapshot.Inventory.xls*")
    Do While fileName <> ""
        tempFile = NormalizeFolderPathHq(tempFolder) & fileName
        On Error Resume Next
        Kill tempFile
        On Error GoTo FailAggregate
        FileCopy NormalizeFolderPathHq(snapshotsFolder) & fileName, tempFile

        Set wbSnap = Application.Workbooks.Open(tempFile, ReadOnly:=True)
        Set lo = FindListObjectByNameHq(wbSnap, TABLE_WAREHOUSE_SNAPSHOT)
        If Not lo Is Nothing Then
            For i = 1 To lo.ListRows.Count
                If SafeTrimHq(GetCellByColumnHq(lo, i, "SKU")) <> "" Then
                    key = SafeTrimHq(GetCellByColumnHq(lo, i, "WarehouseId")) & "|" & SafeTrimHq(GetCellByColumnHq(lo, i, "SKU"))
                    MergeSnapshotRow globalRows, key, lo, i, fileName
                End If
            Next i
        End If
        wbSnap.Close SaveChanges:=False
        Set wbSnap = Nothing

        fileName = Dir$
    Loop

    WriteGlobalSnapshotWorkbook outputPath, globalRows
    report = "Rows=" & CStr(globalRows.Count)
    GenerateGlobalSnapshotFromFolder = True
    Exit Function

FailAggregate:
    On Error Resume Next
    If Not wbSnap Is Nothing Then wbSnap.Close SaveChanges:=False
    On Error GoTo 0
    report = "GenerateGlobalSnapshotFromFolder failed: " & Err.Description
End Function

Private Sub MergeSnapshotRow(ByVal globalRows As Object, _
                             ByVal key As String, _
                             ByVal lo As ListObject, _
                             ByVal rowIndex As Long, _
                             ByVal sourceFile As String)
    Dim entry As Object
    Dim currentDate As Variant
    Dim existingDate As Variant

    If globalRows.Exists(key) Then
        Set entry = globalRows(key)
        currentDate = GetCellByColumnHq(lo, rowIndex, "LastAppliedAtUTC")
        existingDate = entry("LastAppliedAtUTC")
        If IsDate(currentDate) And IsDate(existingDate) Then
            If CDate(currentDate) <= CDate(existingDate) Then Exit Sub
        End If
    Else
        Set entry = CreateObject("Scripting.Dictionary")
        entry.CompareMode = vbTextCompare
        globalRows.Add key, entry
    End If

    entry("WarehouseId") = GetCellByColumnHq(lo, rowIndex, "WarehouseId")
    entry("SKU") = GetCellByColumnHq(lo, rowIndex, "SKU")
    entry("QtyOnHand") = GetCellByColumnHq(lo, rowIndex, "QtyOnHand")
    entry("LastAppliedAtUTC") = GetCellByColumnHq(lo, rowIndex, "LastAppliedAtUTC")
    entry("SourceSnapshot") = sourceFile
End Sub

Private Sub WriteGlobalSnapshotWorkbook(ByVal outputPath As String, ByVal globalRows As Object)
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim headers As Variant
    Dim startCell As Range
    Dim i As Long
    Dim key As Variant
    Dim rowIndex As Long

    EnsureFolderForFileHq outputPath
    CloseWorkbookByFullNameHq outputPath
    On Error Resume Next
    Kill outputPath
    On Error GoTo 0

    Set wb = Application.Workbooks.Add
    headers = Array("WarehouseId", "SKU", "QtyOnHand", "LastAppliedAtUTC", "SourceSnapshot")
    Set ws = wb.Worksheets(1)
    ws.Name = SHEET_GLOBAL_SNAPSHOT
    Set startCell = ws.Range("A1")
    For i = LBound(headers) To UBound(headers)
        startCell.Offset(0, i - LBound(headers)).Value = headers(i)
    Next i

    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(startCell, startCell.Offset(1, UBound(headers) - LBound(headers))), , xlYes)
    lo.Name = TABLE_GLOBAL_SNAPSHOT
    If lo.DataBodyRange Is Nothing Then lo.ListRows.Add
    DeleteAllRowsHq lo

    For Each key In globalRows.Keys
        lo.ListRows.Add
        rowIndex = lo.ListRows.Count
        SetTableRowValueHq lo, rowIndex, "WarehouseId", globalRows(key)("WarehouseId")
        SetTableRowValueHq lo, rowIndex, "SKU", globalRows(key)("SKU")
        SetTableRowValueHq lo, rowIndex, "QtyOnHand", globalRows(key)("QtyOnHand")
        SetTableRowValueHq lo, rowIndex, "LastAppliedAtUTC", globalRows(key)("LastAppliedAtUTC")
        SetTableRowValueHq lo, rowIndex, "SourceSnapshot", globalRows(key)("SourceSnapshot")
    Next key

    wb.SaveAs Filename:=outputPath, FileFormat:=50
    wb.Close SaveChanges:=True
End Sub

Private Function FindListObjectByNameHq(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    On Error Resume Next
    For Each ws In wb.Worksheets
        Set FindListObjectByNameHq = ws.ListObjects(tableName)
        If Not FindListObjectByNameHq Is Nothing Then Exit Function
    Next ws
    On Error GoTo 0
End Function

Private Function GetCellByColumnHq(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As Variant
    Dim idx As Long
    idx = GetColumnIndexHq(lo, columnName)
    If idx = 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    GetCellByColumnHq = lo.DataBodyRange.Cells(rowIndex, idx).Value
End Function

Private Sub SetTableRowValueHq(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String, ByVal valueOut As Variant)
    Dim idx As Long
    idx = GetColumnIndexHq(lo, columnName)
    If idx = 0 Then Exit Sub
    lo.DataBodyRange.Cells(rowIndex, idx).Value = valueOut
End Sub

Private Function GetColumnIndexHq(ByVal lo As ListObject, ByVal columnName As String) As Long
    Dim i As Long
    If lo Is Nothing Then Exit Function
    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).Name, columnName, vbTextCompare) = 0 Then
            GetColumnIndexHq = i
            Exit Function
        End If
    Next i
End Function

Private Sub DeleteAllRowsHq(ByVal lo As ListObject)
    If lo Is Nothing Then Exit Sub
    Do While lo.ListRows.Count > 0
        lo.ListRows(lo.ListRows.Count).Delete
    Loop
End Sub

Private Function NormalizeFolderPathHq(ByVal folderPath As String) As String
    folderPath = Trim$(folderPath)
    If folderPath = "" Then Exit Function
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    NormalizeFolderPathHq = folderPath
End Function

Private Function SafeTrimHq(ByVal valueIn As Variant) As String
    On Error Resume Next
    SafeTrimHq = Trim$(CStr(valueIn))
End Function

Private Sub EnsureFolderForFileHq(ByVal filePath As String)
    Dim folderPath As String
    Dim sepPos As Long

    sepPos = InStrRev(filePath, "\")
    If sepPos <= 0 Then Exit Sub
    folderPath = Left$(filePath, sepPos - 1)
    CreateFolderRecursiveHq folderPath
End Sub

Private Sub CreateFolderRecursiveHq(ByVal folderPath As String)
    Dim parentPath As String
    Dim sepPos As Long

    folderPath = Trim$(folderPath)
    If folderPath = "" Then Exit Sub
    If Len(Dir$(folderPath, vbDirectory)) > 0 Then Exit Sub
    If Right$(folderPath, 1) = "\" Then folderPath = Left$(folderPath, Len(folderPath) - 1)

    sepPos = InStrRev(folderPath, "\")
    If sepPos > 0 Then
        parentPath = Left$(folderPath, sepPos - 1)
        If Right$(parentPath, 1) = ":" Then parentPath = parentPath & "\"
        If parentPath <> "" And Len(Dir$(parentPath, vbDirectory)) = 0 Then CreateFolderRecursiveHq parentPath
    End If
    If Len(Dir$(folderPath, vbDirectory)) = 0 Then MkDir folderPath
End Sub

Private Sub CloseWorkbookByFullNameHq(ByVal fullNameIn As String)
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullNameIn, vbTextCompare) = 0 Then
            On Error Resume Next
            wb.Close SaveChanges:=False
            On Error GoTo 0
            Exit For
        End If
    Next wb
End Sub
