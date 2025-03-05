VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InventoryManagement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
 Private Sub Worksheet_Change(ByVal Target As Range)
    Dim tbl As ListObject
    Dim totalInvCol As ListColumn, lastEditedCol As ListColumn
    Dim emptyRows As Range
    Dim changedCell As Range
    Dim prefix As String
    Dim idNumber As Long
    Dim existingCodes As Object
    Dim cell As Range

    On Error GoTo ErrorHandler

    ' Reference the inventory table
    Set tbl = Me.ListObjects("invSys") ' Replace with the actual table name
    If tbl Is Nothing Then Exit Sub

    ' ----- 1. Update LAST EDITED Column -----
    Set totalInvCol = tbl.ListColumns("TOTAL INV")
    Set lastEditedCol = tbl.ListColumns("LAST EDITED")
    
    If Not Intersect(Target, totalInvCol.DataBodyRange) Is Nothing Then
        Application.EnableEvents = False
        For Each changedCell In Intersect(Target, totalInvCol.DataBodyRange)
            If Not IsEmpty(changedCell.Value) Then
                tbl.ListRows(changedCell.row - tbl.HeaderRowRange.row).Range(lastEditedCol.Index).Value = Now
            End If
        Next changedCell
    End If

    ' ----- 2. Generate Unique Item_Code for New Rows -----
    Set existingCodes = CreateObject("Scripting.Dictionary")
    For Each cell In tbl.ListColumns("Item_Code").DataBodyRange
        If Not IsEmpty(cell.Value) Then existingCodes(cell.Value) = True
    Next cell

    For Each cell In tbl.ListColumns("Item_Code").DataBodyRange
        If IsEmpty(cell.Value) Then
            prefix = "ITEM-"
            idNumber = 1
            Do While existingCodes.Exists(prefix & Format(idNumber, "0000"))
                idNumber = idNumber + 1
            Loop
            cell.Value = prefix & Format(idNumber, "0000")
            existingCodes.Add cell.Value, True
        End If
    Next cell

    ' ----- 3. Clear Item_Code on Row Deletion -----
    On Error Resume Next
    Set emptyRows = tbl.ListColumns("Item_Code").DataBodyRange.SpecialCells(xlCellTypeBlanks)
    On Error GoTo 0

    If Not emptyRows Is Nothing Then
        emptyRows.ClearContents
    End If

ExitHandler:
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    Debug.Print "Error: " & Err.Description
    Resume ExitHandler
End Sub

Sub UpdateUndoRedoStatus()

End Sub

