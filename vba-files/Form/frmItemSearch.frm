VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmItemSearch 
   Caption         =   "Item Search"
   ClientHeight    =   4950
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
    Dim i As Long, matchIndex As Long
    
    searchText = LCase(Me.txtBox.Text)
    
    ' If the text box is empty, clear the list box selection and exit.
    If Len(Trim(searchText)) = 0 Then
        Me.lstBox.ListIndex = -1
        Exit Sub
    End If
    
    ' Iterate through the list box items to find the first match.
    For i = 0 To Me.lstBox.ListCount - 1
        If InStr(1, LCase(Me.lstBox.List(i)), searchText) > 0 Then
            matchIndex = i
            Me.lstBox.ListIndex = matchIndex
            
            ' FIX 1: Better centering calculation
            Dim visibleItems As Long, centerPos As Long
            visibleItems = Int(Me.lstBox.Height / 15)  ' Approx height per item
            centerPos = Application.Max(0, matchIndex - Int(visibleItems / 2))
            Me.lstBox.TopIndex = centerPos
            Exit Sub
        End If
    Next i
    
    ' If no match is found, deselect any item.
    Me.lstBox.ListIndex = -1
End Sub

' When the user clicks on an item in the list box
Private Sub lstBox_Click()
    ' Keep the item highlighted but don't update the search text
End Sub

' Commit the selection if the user presses Tab or Enter in textbox
Private Sub txtBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
        CommitSelectionAndClose
        KeyCode = 0  ' Prevent default handling
    End If
End Sub

' FIX 2: Ensure Enter key works when list box has focus
Private Sub lstBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = vbKeyReturn Then
        CommitSelectionAndClose
        KeyAscii = 0  ' Prevent default handling
    End If
End Sub

' Also handle key down for Enter in list box
Private Sub lstBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
        CommitSelectionAndClose
        KeyCode = 0  ' Prevent default handling
    End If
End Sub

' Double-clicking on an item also commits the selection.
Private Sub lstBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    CommitSelectionAndClose
End Sub

' FIX 3: Improved CommitSelectionAndClose to fix ORDER_NUMBER copying, more robust empty check
Public Sub CommitSelectionAndClose()
    Static isRunning As Boolean
    
    ' Prevent recursive calls or double-execution
    If isRunning Then Exit Sub
    isRunning = True
    
    Dim chosenValue As String
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim currentRowIndex As Long
    Dim prevRowOrder As Variant, currentOrderValue As Variant
    
    ' Decide on the value to commit
    If Me.lstBox.ListIndex <> -1 Then
        chosenValue = Me.lstBox.List(Me.lstBox.ListIndex)
    ElseIf Trim(Me.txtBox.Text) <> "" Then
        chosenValue = Me.txtBox.Text
    Else
        If Not gSelectedCell Is Nothing Then gSelectedCell.ClearContents
        isRunning = False
        Unload Me
        Exit Sub
    End If
    
    ' Apply the chosen value
    If Not gSelectedCell Is Nothing Then
        gSelectedCell.Value = chosenValue
        
        ' Get the table and check if we're modifying the ITEMS column
        If Not gSelectedCell.ListObject Is Nothing Then
            Set tbl = gSelectedCell.ListObject
            
            ' Find column indexes directly
            Dim orderCol As Long, itemsCol As Long
            
            On Error Resume Next
            ' Get columns by their exact names
            orderCol = WorksheetFunction.Match("ORDER_NUMBER", tbl.HeaderRowRange, 0)
            itemsCol = WorksheetFunction.Match("ITEMS", tbl.HeaderRowRange, 0)
            On Error GoTo 0
            
            ' Only proceed if both columns exist
            If orderCol > 0 And itemsCol > 0 Then
                ' Check if we're in the ITEMS column
                If gSelectedCell.Column = tbl.HeaderRowRange.Cells(1, itemsCol).Column Then
                    ' Get data row number (1-based)
                    currentRowIndex = gSelectedCell.row - tbl.HeaderRowRange.row
                    
                    ' If not in first row, check if we need to copy ORDER_NUMBER
                    If currentRowIndex > 1 Then
                        ' Get current ORDER_NUMBER cell value
                        currentOrderValue = tbl.DataBodyRange.Cells(currentRowIndex, orderCol).Value
                        
                        ' More robust empty check - check if truly empty or just whitespace
                        If IsEmpty(currentOrderValue) Or Trim(CStr(currentOrderValue)) = "" Then
                            ' Get value from previous row
                            prevRowOrder = tbl.DataBodyRange.Cells(currentRowIndex - 1, orderCol).Value
                            
                            ' Copy to current row only if previous row has a value
                            If Not IsEmpty(prevRowOrder) Then
                                tbl.DataBodyRange.Cells(currentRowIndex, orderCol).Value = prevRowOrder
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    isRunning = False
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
