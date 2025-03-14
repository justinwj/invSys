<<<<<<< HEAD
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShipmentsTally"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'========================
' OrdersTally worksheet module)
' ========================
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim lo As ListObject
    Dim hdrCell As Range
    Dim itemsHeader As String
    Dim itemsCol As Long
    Dim found As Boolean
    
    itemsHeader = "ITEMS"
    found = False
    
    ' Check if the selected cell is part of a table.
    On Error Resume Next
    Set lo = Target.ListObject
    On Error GoTo 0
    
    If Not lo Is Nothing Then
        ' Do nothing if the header row is clicked.
        If Not Intersect(Target, lo.HeaderRowRange) Is Nothing Then Exit Sub
        
        ' Loop through the header row to find the column labeled "ITEMS"
        For Each hdrCell In lo.HeaderRowRange
            If Trim(UCase(hdrCell.Value)) = UCase(itemsHeader) Then
                itemsCol = hdrCell.Column
                found = True
                Exit For
            End If
        Next hdrCell
        
        ' If the target cell is in the "ITEMS" column, launch the search tool.
        If found And Target.Column = itemsCol Then
            Set gSelectedCell = Target.Cells(1, 1)
            Dim searchForm As New frmItemSearch
            With searchForm
                .StartUpPosition = 0  ' Manual positioning.
                .Left = Application.ActiveWindow.PointsToScreenPixelsX(Target.Left)
                .Top = Application.ActiveWindow.PointsToScreenPixelsY(Target.Top + Target.Height)
                .Show vbModeless
            End With
            Exit Sub
        End If
    End If
    
    ' If selection moves away from an ITEMS cell, attempt to commit and close the search form.
    On Error Resume Next
    frmItemSearch.CommitSelectionAndClose
    On Error GoTo 0
End Sub

=======
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

' This event fires whenever a cell is selected - by clicking, tabbing, or arrow keys
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Only activate for single cell selection
    If Target.Cells.Count > 1 Then Exit Sub
    
    On Error Resume Next
    Dim tbl As ListObject
    Set tbl = Me.ListObjects("ShipmentsTally")
    If tbl Is Nothing Then Exit Sub
    
    ' Check if it's in the ITEMS column
    If Target.Column = tbl.ListColumns("ITEMS").Range.Column Then
        If Target.Row > tbl.HeaderRowRange.Row Then
            Set gSelectedCell = Target
            frmItemSearch.Show vbModeless
        End If
    End If
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

>>>>>>> 0e2f3dc (Refactored a lot, tally system does not work as intended but big button works)
