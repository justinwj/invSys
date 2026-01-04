VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPickerRouter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mSearch As cDynItemSearch

Public Sub HandleSelectionChange(ByVal Target As Range)
    ' Placeholder: no automatic picker on selection to avoid noisy popups.
End Sub

Public Function HandleBeforeDoubleClick(ByVal Target As Range, ByRef Cancel As Boolean) As Boolean
    If Target Is Nothing Then Exit Function
    Dim lo As ListObject
    On Error Resume Next
    Set lo = Target.ListObject
    On Error GoTo 0
    If lo Is Nothing Then Exit Function

    Dim colName As String
    colName = ColumnNameFromCell(lo, Target)
    If colName = "" Then Exit Function

    If IsItemPickerTable(lo) And (LCase$(colName) = "item" Or LCase$(colName) = "items") Then
        EnsureSearch
        mSearch.ShowForCell Target
        Cancel = True
        HandleBeforeDoubleClick = True
        Exit Function
    End If

    If LCase$(lo.Name) = "rc_recipechoose" And LCase$(colName) = "recipe" Then
        MsgBox "Recipe picker is not wired yet.", vbInformation
        Cancel = True
        HandleBeforeDoubleClick = True
        Exit Function
    End If
End Function

Private Sub EnsureSearch()
    If mSearch Is Nothing Then Set mSearch = New cDynItemSearch
End Sub

Private Function IsItemPickerTable(ByVal lo As ListObject) As Boolean
    If lo Is Nothing Then Exit Function
    Dim nm As String: nm = LCase$(lo.Name)
    If nm = "inventorypalette_generated" Then
        IsItemPickerTable = True
    ElseIf nm = "ip_chooseitem" Then
        IsItemPickerTable = True
    ElseIf nm Like "proc_*_palette" Then
        IsItemPickerTable = True
    End If
End Function

Private Function ColumnNameFromCell(lo As ListObject, targetCell As Range) As String
    If lo Is Nothing Then Exit Function
    If targetCell Is Nothing Then Exit Function
    Dim colIdx As Long
    colIdx = targetCell.Column - lo.Range.Column + 1
    If colIdx < 1 Or colIdx > lo.ListColumns.Count Then Exit Function
    ColumnNameFromCell = lo.ListColumns(colIdx).Name
End Function

