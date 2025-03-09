VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReceivedTally"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'========================
' ReceivedTally worksheet module)
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
