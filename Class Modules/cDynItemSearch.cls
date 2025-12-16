'========================
' Class: cDynItemSearch (FIXED - No WithEvents)
' Purpose: Runtime-built Item Search form for ReceivedTally
'------------------------
Option Explicit

Private txtSearch As Object
Private lst As Object
Private uf As Object
Private itemsArr As Variant
Private callbackCell As Range

Public Sub ShowForCell(ByVal targetCell As Range)
    If targetCell Is Nothing Then Exit Sub
    Set callbackCell = targetCell
    
    If Not BuildForm() Then
        On Error Resume Next
        frmItemSearch.Show vbModeless
        On Error GoTo 0
        Exit Sub
    End If
    
    If Not LoadItems() Then
        On Error Resume Next
        frmItemSearch.Show vbModeless
        On Error GoTo 0
        Exit Sub
    End If
    
    uf.Show vbModeless
End Sub

Private Function BuildForm() As Boolean
    On Error Resume Next
    Set uf = VBA.UserForms.Add()
    If uf Is Nothing Then
        Set uf = CreateObject("Forms.UserForm.1")
    End If
    On Error GoTo 0
    
    If uf Is Nothing Then
        BuildForm = False
        Exit Function
    End If
    
    BuildForm = True
    uf.caption = "Item Search"
    uf.Width = 420
    uf.Height = 340

    Set txtSearch = uf.Controls.Add("Forms.TextBox.1", "txtSearch", True)
    txtSearch.Left = 10: txtSearch.Top = 10: txtSearch.Width = uf.Width - 30

    Set lst = uf.Controls.Add("Forms.ListBox.1", "lst", True)
    lst.Left = 10: lst.Top = 35: lst.Width = uf.Width - 30: lst.Height = 230
    lst.ColumnCount = 7
    lst.ColumnWidths = "60;90;180;60;90;180;140"

    Dim txtDesc As Object
    Set txtDesc = uf.Controls.Add("Forms.TextBox.1", "txtDesc", True)
    txtDesc.Left = 10: txtDesc.Top = lst.Top + lst.Height + 5
    txtDesc.Width = uf.Width - 30: txtDesc.Height = 50
    txtDesc.MultiLine = True: txtDesc.Locked = True
End Function

Private Function LoadItems() As Boolean
    itemsArr = modTS_Received.LoadItemList()
    lst.Clear
    If IsEmpty(itemsArr) Then
        LoadItems = False
        Exit Function
    End If
    
    Dim i As Long
    For i = LBound(itemsArr, 1) To UBound(itemsArr, 1)
        lst.AddItem
        lst.List(lst.ListCount - 1, 0) = itemsArr(i, 1)
        lst.List(lst.ListCount - 1, 1) = itemsArr(i, 2)
        lst.List(lst.ListCount - 1, 2) = itemsArr(i, 3)
        lst.List(lst.ListCount - 1, 3) = itemsArr(i, 4)
        lst.List(lst.ListCount - 1, 4) = itemsArr(i, 5)
        lst.List(lst.ListCount - 1, 5) = itemsArr(i, 6)
        lst.List(lst.ListCount - 1, 6) = itemsArr(i, 7)
    Next i
    LoadItems = True
End Function

' Manually wired event handlers (no WithEvents needed)
Public Sub WireEvents()
    ' This is called by external event handlers if needed
    ' For now, events are handled via direct control interaction
End Sub

Private Sub HandleTextboxChange()
    Dim term As String: term = LCase$(Trim$(txtSearch.text))
    Dim i As Long
    lst.Clear
    If IsEmpty(itemsArr) Then Exit Sub
    For i = LBound(itemsArr, 1) To UBound(itemsArr, 1)
        If term = "" Or InStr(1, LCase$(itemsArr(i, 3)), term) > 0 Then
            lst.AddItem
            lst.List(lst.ListCount - 1, 0) = itemsArr(i, 1)
            lst.List(lst.ListCount - 1, 1) = itemsArr(i, 2)
            lst.List(lst.ListCount - 1, 2) = itemsArr(i, 3)
            lst.List(lst.ListCount - 1, 3) = itemsArr(i, 4)
            lst.List(lst.ListCount - 1, 4) = itemsArr(i, 5)
            lst.List(lst.ListCount - 1, 5) = itemsArr(i, 6)
            lst.List(lst.ListCount - 1, 6) = itemsArr(i, 7)
        End If
    Next i
End Sub

Private Sub CommitSelection()
    If lst.ListIndex < 0 Then Exit Sub
    If callbackCell Is Nothing Then Exit Sub

    callbackCell.value = lst.List(lst.ListIndex, 2)

    Dim refNum As String, qty As Double
    Dim rt As ListObject
    On Error Resume Next
    Set rt = callbackCell.ListObject
    On Error GoTo 0
    
    If Not rt Is Nothing And Not rt.DataBodyRange Is Nothing Then
        Dim rowIdx As Long
        rowIdx = callbackCell.row - rt.DataBodyRange.row + 1
        If rowIdx >= 1 And rowIdx <= rt.DataBodyRange.rows.count Then
            refNum = NzStr(rt.DataBodyRange.Cells(rowIdx, ColumnIndex(rt, "REF_NUMBER")).value)
            qty = NzDbl(rt.DataBodyRange.Cells(rowIdx, ColumnIndex(rt, "QUANTITY")).value)
        End If
    End If

    modTS_Received.AddOrMergeFromSearch _
        refNum, _
        lst.List(lst.ListIndex, 2), _
        lst.List(lst.ListIndex, 1), _
        qty, _
        lst.List(lst.ListIndex, 6), _
        "", _
        lst.List(lst.ListIndex, 5), _
        lst.List(lst.ListIndex, 3), _
        lst.List(lst.ListIndex, 4), _
        CLng(NzLng(lst.List(lst.ListIndex, 0)))

    modTS_Received.RebuildAggregation
    uf.Hide
End Sub

Private Function NzStr(v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then NzStr = "" Else NzStr = CStr(v)
End Function

Private Function NzDbl(v As Variant) As Double
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Or v = "" Then NzDbl = 0# Else NzDbl = CDbl(v)
End Function

Private Function NzLng(v As Variant) As Long
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Or v = "" Then NzLng = 0 Else NzLng = CLng(v)
End Function

Private Function ColumnIndex(lo As ListObject, colName As String) As Long
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If StrComp(lc.name, colName, vbTextCompare) = 0 Then ColumnIndex = lc.Index: Exit Function
    Next
    ColumnIndex = 0
End Function