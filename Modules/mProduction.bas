Attribute VB_Name = "mProduction"
Option Explicit

' Production system core module (wiring + helpers).

Private Const SHEET_PRODUCTION As String = "Production"
Private Const SHEET_TEMPLATES As String = "TemplatesTable"

Private Const TABLE_RECIPE_CHOOSER As String = "RC_RecipeChoose"
Private Const TABLE_INV_PALETTE_GENERATED As String = "InventoryPalette_generated"

Private mRowCountCache As Object

Public Sub InitializeProductionUI()
    ' Placeholder for future UI setup (buttons, toggles, etc.).
End Sub

' ===== Worksheet event entry points =====
Public Sub HandleProductionSelectionChange(ByVal Target As Range)
    If Target Is Nothing Then Exit Sub
    If Not IsOnProductionSheet(Target) Then Exit Sub
    Dim router As New cPickerRouter
    router.HandleSelectionChange Target
End Sub

Public Sub HandleProductionBeforeDoubleClick(ByVal Target As Range, ByRef Cancel As Boolean)
    If Target Is Nothing Then Exit Sub
    If Not IsOnProductionSheet(Target) Then Exit Sub
    Dim router As New cPickerRouter
    If router.HandleBeforeDoubleClick(Target, Cancel) Then Exit Sub
End Sub

Public Sub HandleProductionChange(ByVal Target As Range)
    If Target Is Nothing Then Exit Sub
    If Not IsOnProductionSheet(Target) Then Exit Sub

    Dim lo As ListObject
    On Error Resume Next
    Set lo = Target.ListObject
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub

    If IsPaletteTable(lo) Then
        EnsureRowCountCache
        Dim key As String: key = lo.Name
        Dim newCount As Long: newCount = ListObjectRowCount(lo)
        Dim oldCount As Long
        If mRowCountCache.Exists(key) Then oldCount = CLng(mRowCountCache(key))
        If newCount > oldCount Then
            Dim bandMgr As New cTableBandManager
            bandMgr.Init lo.Parent
            bandMgr.ExpandBandForTable lo, (newCount - oldCount)
        End If
        mRowCountCache(key) = newCount
    End If
End Sub

' ===== Band/table helpers =====
Private Sub EnsureRowCountCache()
    If mRowCountCache Is Nothing Then
        Set mRowCountCache = CreateObject("Scripting.Dictionary")
    End If
End Sub

Private Function IsOnProductionSheet(ByVal Target As Range) As Boolean
    On Error Resume Next
    IsOnProductionSheet = (Target.Worksheet.Name = SHEET_PRODUCTION)
    On Error GoTo 0
End Function

Private Function IsPaletteTable(lo As ListObject) As Boolean
    If lo Is Nothing Then Exit Function
    Dim nm As String: nm = LCase$(lo.Name)
    If nm = LCase$(TABLE_INV_PALETTE_GENERATED) Then
        IsPaletteTable = True
    ElseIf nm Like "proc_*_palette" Then
        IsPaletteTable = True
    End If
End Function

Private Function ListObjectRowCount(lo As ListObject) As Long
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    ListObjectRowCount = lo.DataBodyRange.Rows.Count
End Function

' ===== Generic helpers =====
Public Function GetProductionSheet() As Worksheet
    Set GetProductionSheet = SheetExists(SHEET_PRODUCTION)
End Function

Public Function SheetExists(nameOrCode As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, nameOrCode, vbTextCompare) = 0 _
           Or StrComp(ws.CodeName, nameOrCode, vbTextCompare) = 0 Then
            Set SheetExists = ws
            Exit Function
        End If
    Next ws
End Function

Public Function GetListObject(ws As Worksheet, tableName As String) As ListObject
    On Error Resume Next
    Set GetListObject = ws.ListObjects(tableName)
    On Error GoTo 0
End Function

Public Function ColumnIndex(lo As ListObject, colName As String) As Long
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If StrComp(lc.Name, colName, vbTextCompare) = 0 Then
            ColumnIndex = lc.Index
            Exit Function
        End If
    Next lc
    ColumnIndex = 0
End Function

