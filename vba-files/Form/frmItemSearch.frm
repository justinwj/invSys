VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmItemSearch 
   Caption         =   "Item Search"
   ClientHeight    =   3645
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5400
   OleObjectBlob   =   "frmItemSearch.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmItemSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FullItemList As Variant

Private Sub UserForm_MouseScroll()
    EnableMouseScroll frmItemSearch
End Sub

Private Sub UserForm_Activate()
    Me.txtBox.SetFocus
    ' Place the caret at the beginning with no selection.
    Me.txtBox.SelStart = 0
    Me.txtBox.SelLength = 0
End Sub

Private Sub UserForm_Initialize()
    ' Load the full list of items from the invSys table.
    FullItemList = modTS_Data.LoadItemList()
    ' Populate lstBox with the full list.
    Call PopulateListBox(FullItemList)
    
    ' Pre-populate txtBox with the current cell value (if any) without selecting it.
    If Not gSelectedCell Is Nothing Then
        If Not IsEmpty(gSelectedCell.Value) Then
            Me.txtBox.Text = CStr(gSelectedCell.Value)
            ' Place the caret at the beginning without selecting text.
            Me.txtBox.SelStart = 0
            Me.txtBox.SelLength = 0
        End If
    End If
    
    ' Attempt to match the current txtBox content without altering it.
    Call txtBox_Change
    EnableMouseScroll Me
End Sub

' Instead of filtering out non-matching items, just move the highlighter to the nearest match.
Private Sub txtBox_Change()
    Dim searchText As String
    Dim i As Long
    
    searchText = LCase(Me.txtBox.Text)
    
    ' If the text box is empty, clear the list box selection and exit.
    If Len(Trim(searchText)) = 0 Then
        Me.lstBox.ListIndex = -1
        Exit Sub
    End If
    
    ' Iterate through the list box items to find the first match.
    For i = 0 To Me.lstBox.ListCount - 1
        If InStr(1, LCase(Me.lstBox.List(i)), searchText) > 0 Then
            Me.lstBox.ListIndex = i
            Exit Sub
        End If
    Next i
    
    ' If no match is found, deselect any item.
    Me.lstBox.ListIndex = -1
End Sub

' When the user clicks on an item in the list box, update txtBox immediately.
Private Sub lstBox_Click()
    ' Simply highlight the item in the listbox.
    ' Do not update txtBox.Text here so that user typing is not interfered with.
    ' The chosen listbox value will be used during CommitSelectionAndClose.
    ' (Optionally, you might want to visually indicate the selection if needed.)
End Sub

' Commit the selection if the user presses Tab or Enter.
Private Sub txtBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
        CommitSelectionAndClose
        KeyCode = 0  ' Prevent default handling.
    End If
End Sub

' Double-clicking on an item also commits the selection.
Private Sub lstBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    CommitSelectionAndClose
End Sub

' On exit, commit the value:
' - If txtBox is empty, clear the cell.
' - If an item is highlighted in lstBox, use that value.
' - Otherwise, use the text entered in txtBox.
Public Sub CommitSelectionAndClose()
    Dim chosenValue As String
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim orderCol As Long, itemsCol As Long
    Dim currentRowIndex As Long, prevRowOrder As Variant

    ' Decide on the value to commit. If nothing in txtBox then clear the cell.
    If Trim(Me.txtBox.Text) = "" Then
        If Not gSelectedCell Is Nothing Then gSelectedCell.ClearContents
    Else
        ' Use the listbox-selected value if available; otherwise, use the textbox text.
        If Me.lstBox.ListIndex <> -1 Then
            chosenValue = Me.lstBox.Value
        Else
            chosenValue = Me.txtBox.Text
        End If
        gSelectedCell.Value = chosenValue
        
        ' If the active cell is part of a table, check if it belongs to the ITEMS column.
        Set ws = gSelectedCell.Worksheet
        If Not gSelectedCell.ListObject Is Nothing Then
            Set tbl = gSelectedCell.ListObject
            orderCol = modTS_Data.GetColumnIndexByHeader("ORDER_NUMBER")
            itemsCol = modTS_Data.GetColumnIndexByHeader("ITEMS")
            
            ' Confirm that gSelectedCell is in the ITEMS column
            If gSelectedCell.Column = tbl.Range.Cells(1, itemsCol).Column Then
                ' Calculate which data row (1-based) we are in.
                currentRowIndex = gSelectedCell.Row - tbl.HeaderRowRange.Row
                If currentRowIndex > 1 Then
                    ' Get the ORDER_NUMBER from the previous row in the table.
                    prevRowOrder = tbl.DataBodyRange.Cells(currentRowIndex - 1, orderCol).Value
                    If Not IsEmpty(prevRowOrder) Then
                        ' Copy the ORDER_NUMBER value to the current row.
                        tbl.DataBodyRange.Cells(currentRowIndex, orderCol).Value = prevRowOrder
                    End If
                End If
            End If
        End If
    End If
    Unload Me
End Sub

' Populate the list box with the full list of items.
Private Sub PopulateListBox(itemArray As Variant)
    Dim i As Long
    Me.lstBox.Clear
    For i = LBound(itemArray) To UBound(itemArray)
        Me.lstBox.AddItem itemArray(i)
    Next i
End Sub

Private Sub UserForm_Deactivate()
    CommitSelectionAndClose
End Sub
