VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShipmentsTally"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ShipmentsTally worksheet class module
Option Explicit

Private isShowingForm As Boolean

' This event fires whenever a cell is selected - by clicking, tabbing, or arrow keys
Option Explicit

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Only process single cells
    If Target.Cells.Count > 1 Then Exit Sub
    
    ' Skip if the form is already open
    If modGlobals.IsFormOpen("frmItemSearch") Then Exit Sub
    
    On Error Resume Next
    ' Check if we're in the ITEMS column of ShipmentsTally table
    Dim tbl As ListObject
    Set tbl = Me.ListObjects("ShipmentsTally")
    
    If Not tbl Is Nothing Then
        ' Direct check if this cell is in the ITEMS column
        If Target.Column = tbl.ListColumns("ITEMS").Range.Column And _
           Target.Row > tbl.HeaderRowRange.Row Then
            
            ' Set selected cell and show form
            Set gSelectedCell = Target
            modGlobals.PauseTimer
            frmItemSearch.Show vbModeless
        End If
    End If
    On Error GoTo 0
End Sub

' Also handle BeforeDoubleClick as backup
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    On Error Resume Next
    Dim tbl As ListObject
    Set tbl = Me.ListObjects("ShipmentsTally")
    If tbl Is Nothing Then Exit Sub
    
    ' Check if it's in the ITEMS column
    If Target.Column = tbl.ListColumns("ITEMS").Range.Column Then
        If Target.Row > tbl.HeaderRowRange.Row Then
            Set gSelectedCell = Target
            frmItemSearch.Show vbModeless
            Cancel = True
        End If
    End If
End Sub

' Add right-click handler
Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
    On Error Resume Next
    Dim tbl As ListObject
    Set tbl = Me.ListObjects("ShipmentsTally")
    If tbl Is Nothing Then Exit Sub
    
    ' Get the ITEMS column range
    Dim itemsColIndex As Long
    For itemsColIndex = 1 To tbl.ListColumns.Count
        If UCase(tbl.ListColumns(itemsColIndex).Name) = "ITEMS" Then Exit For
    Next itemsColIndex
    
    ' If we didn't find the ITEMS column
    If itemsColIndex > tbl.ListColumns.Count Then Exit Sub
    
    ' Check if the target cell is in the data area of the ITEMS column
    If Target.Column = tbl.ListColumns(itemsColIndex).Range.Column Then
        If Target.Row > tbl.HeaderRowRange.Row Then
            ' Show the form on right-click in ITEMS column
            Debug.Print "ShipmentsTally: Right-click in ITEMS column, showing form"
            Set gSelectedCell = Target
            frmItemSearch.Show vbModeless
            Cancel = True ' Prevent default context menu
        End If
    End If
    On Error GoTo 0
End Sub


Private Function IsFormLoaded(formName As String) As Boolean
    Dim frm As Object
    IsFormLoaded = False
    
    On Error Resume Next
    For Each frm In UserForms
        If frm.Name = formName Then
            IsFormLoaded = True
            Exit For
        End If
    Next frm
    On Error GoTo 0
End Function

