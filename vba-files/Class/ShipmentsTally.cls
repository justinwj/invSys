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
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Prevent recursive calls
    If isShowingForm Then Exit Sub
    
    ' Simplified version with minimal error sources
    If Target.Cells.Count > 1 Then Exit Sub
    
    On Error GoTo ErrorHandler
    Dim tbl As ListObject
    Set tbl = Me.ListObjects("ShipmentsTally") ' or "ReceivedTally" for that sheet
    
    ' Only proceed if we found the table
    If tbl Is Nothing Then Exit Sub
    
    ' Direct check for ITEMS column by name 
    Dim itemsCol As ListColumn
    Set itemsCol = tbl.ListColumns("ITEMS")
    
    ' Check if this is a cell in the ITEMS column's data range
    If Not Intersect(Target, itemsCol.DataBodyRange) Is Nothing Then
        ' Prevent recursive calls
        isShowingForm = True
        
        Debug.Print "Cell in ITEMS column selected"
        Set gSelectedCell = Target
        
        If Not IsFormLoaded("frmItemSearch") Then
            frmItemSearch.Show vbModeless
        End If
        
        isShowingForm = False
    End If
    
    Exit Sub
ErrorHandler:
    isShowingForm = False
    Debug.Print "Error in Worksheet_SelectionChange: " & Err.Description
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

