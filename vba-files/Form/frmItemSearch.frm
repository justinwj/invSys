VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmItemSearch 
   Caption         =   "Item Search"
   ClientHeight    =   5085
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6480
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
Private LastSearchText As String
Private LastSearchTime As Double
Private SearchFirstCharIndex() As Long  ' Array to store first character indexes
Private Const MIN_SEARCH_INTERVAL As Double = 0.2  ' Minimum seconds between searches

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
    ' Add variable declaration
    Dim i As Long
    
    ' Set up the list box columns
    Me.lstBox.ColumnCount = 4  ' ITEM_CODE, ROW#, ITEM, LOCATION
    Me.lstBox.ColumnWidths = "70;40;150;80"
    
    ' Load inventory items
    Dim items As Variant
    items = modTS_Data.LoadItemList()
    
    ' Populate list box with items
    If Not IsEmpty(items) Then
        For i = LBound(items, 1) To UBound(items, 1)
            Me.lstBox.AddItem items(i, 0)  ' ITEM_CODE
            Me.lstBox.List(Me.lstBox.ListCount - 1, 1) = items(i, 1)  ' ROW#
            Me.lstBox.List(Me.lstBox.ListCount - 1, 2) = items(i, 2)  ' ITEM name
            Me.lstBox.List(Me.lstBox.ListCount - 1, 3) = items(i, 3)  ' LOCATION
        Next i
    End If
    
    ' Configure the description textbox for word wrapping
    Me.txtBox2.MultiLine = True
    Me.txtBox2.WordWrap = True
    Me.txtBox2.Locked = True ' Make it read-only
    Me.txtBox2.BackColor = RGB(255, 255, 255) ' Keep white background even when locked
    
    ' Populate lstBox with the full list.
    Call PopulateListBox(FullItemList)
    
    ' Create first character index for faster searching
    BuildFirstCharIndex
    
    ' Pre-populate txtBox with the current cell value (if any) without selecting it.
    If Not gSelectedCell Is Nothing Then
        If Not IsEmpty(gSelectedCell.Value) Then
            Me.txtBox.text = CStr(gSelectedCell.Value)
            ' Place the caret at the beginning without selecting text.
            Me.txtBox.SelStart = 0
            Me.txtBox.SelLength = 0
        End If
    End If
    
    ' Attempt to match the current txtBox content without altering it.
    Call txtBox_Change
    EnableMouseScroll Me
End Sub

' Build an index of where each first character appears in the list for faster searching
Private Sub BuildFirstCharIndex()
    Dim i As Long, char As String
    Dim dict As Object
    
    ' Create a dictionary to track the first occurrence of each character
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Initialize the array with -1 (not found)
    ReDim SearchFirstCharIndex(0 To 255)
    For i = 0 To 255
        SearchFirstCharIndex(i) = -1
    Next i
    
    ' Go through the list and record the first occurrence of each first character
    For i = 0 To Me.lstBox.ListCount - 1
        If Me.lstBox.List(i, 2) <> "" Then
            char = UCase(Left$(Me.lstBox.List(i, 2), 1))
            ' Only record the first occurrence
            If Asc(char) <= 255 And SearchFirstCharIndex(Asc(char)) = -1 Then
                SearchFirstCharIndex(Asc(char)) = i
            End If
        End If
    Next i
End Sub

' Update the txtBox_Change event with better error handling for fast typing
Private Sub txtBox_Change()
    Dim currentTime As Double
    Dim searchText As String, firstChar As String
    Dim i As Long, matchIndex As Long, startIndex As Long
    Dim visibleItems As Long, centerPos As Long
    
    ' Get current time and search text
    currentTime = Timer
    searchText = LCase(Trim(Me.txtBox.Text))
    
    ' Only search if:
    ' 1. Search text has changed significantly, OR
    ' 2. Enough time has passed since last search, OR
    ' 3. Text is empty or very short
    If searchText <> LastSearchText And _
       (currentTime - LastSearchTime >= MIN_SEARCH_INTERVAL Or _
        Len(searchText) <= 2) Then
        
        ' Update tracking variables
        LastSearchTime = currentTime
        LastSearchText = searchText
        
        ' If the text box is empty, clear the list box selection and exit
        If Len(searchText) = 0 Then
            Me.lstBox.ListIndex = -1
            Exit Sub
        End If
        
        ' Get the first character and find its index position
        On Error Resume Next
        firstChar = UCase(Left$(searchText, 1))
        If Len(firstChar) > 0 And Asc(firstChar) <= 255 Then
            startIndex = SearchFirstCharIndex(Asc(firstChar))
        Else
            startIndex = 0
        End If
        On Error GoTo 0
        
        ' If first character not indexed, start from beginning
        If startIndex = -1 Then startIndex = 0
        
        ' Optimized search strategy
        matchIndex = -1
        
        On Error Resume Next
        ' First pass: Search from the first character index position
        For i = startIndex To Me.lstBox.ListCount - 1
            If InStr(1, LCase(Me.lstBox.List(i, 2)), searchText) > 0 Then
                matchIndex = i
                Exit For
            End If
        Next i
        
        ' Second pass: If not found and we started from a specific index, 
        ' search from beginning to that index
        If matchIndex = -1 And startIndex > 0 Then
            For i = 0 To startIndex - 1
                If InStr(1, LCase(Me.lstBox.List(i, 2)), searchText) > 0 Then
                    matchIndex = i
                    Exit For
                End If
            Next i
        End If
        On Error GoTo 0
        
        ' Update UI with results
        If matchIndex <> -1 Then
            Me.lstBox.ListIndex = matchIndex
            
            ' FIXED: Better centering calculation with error handling
            On Error Resume Next
            ' Calculate visible items - ensure it's at least 1
            visibleItems = Int(Me.lstBox.Height / 15)  ' Approx height per item
            If visibleItems < 1 Then visibleItems = 1
            
            ' Safe calculation for center position
            If matchIndex > Int(visibleItems / 2) Then
                centerPos = matchIndex - Int(visibleItems / 2)
            Else
                centerPos = 0
            End If
            
            ' Set top index safely
            Me.lstBox.TopIndex = centerPos
            On Error GoTo 0
            
            ' Update description
            UpdateDescription
        Else
            Me.lstBox.ListIndex = -1
            Me.txtBox2.Text = ""
        End If
    End If
End Sub

' When the user clicks on an item in the list box
Private Sub lstBox_Click()
    ' Keep the item highlighted but don't update the search text
    UpdateDescription
End Sub

' Add handler for keyboard navigation in list box
Private Sub lstBox_Change()
    UpdateDescription
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
    
    ' Prevent recursive calls
    If isRunning Then Exit Sub
    isRunning = True
    
    Dim chosenValue As String
    Dim chosenItemCode As String
    Dim chosenRowNum As String  ' Use actual ROW#
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim currentRowIndex As Long
    
    ' Get selected values
    If Me.lstBox.ListIndex <> -1 Then
        chosenItemCode = Me.lstBox.List(Me.lstBox.ListIndex, 0)  ' ITEM_CODE
        chosenRowNum = Me.lstBox.List(Me.lstBox.ListIndex, 1)    ' ROW#
        chosenValue = Me.lstBox.List(Me.lstBox.ListIndex, 2)     ' Item name
    ElseIf Trim(Me.txtBox.Text) <> "" Then
        chosenValue = Me.txtBox.Text
        chosenItemCode = ""
        chosenRowNum = ""
    Else
        If Not gSelectedCell Is Nothing Then gSelectedCell.ClearContents
        isRunning = False
        Unload Me
        Exit Sub
    End If
    
    ' Apply the chosen value to the cell
    If Not gSelectedCell Is Nothing Then
        ' Set the visible item name in the cell
        gSelectedCell.Value = chosenValue
        
        ' Store reference data in cell comment
        On Error Resume Next
        If gSelectedCell.Comment Is Nothing Then
            gSelectedCell.AddComment
        End If
        gSelectedCell.Comment.Text "ITEM_CODE: " & chosenItemCode & vbCrLf & _
                                  "ROW#: " & chosenRowNum
        gSelectedCell.Comment.Visible = False
        On Error GoTo 0
        
        ' Update the UOM in the current row if needed
        On Error Resume Next
        Set ws = gSelectedCell.Worksheet
        
        If ws.Name = "ShipmentsTally" Or ws.Name = "ReceivedTally" Then
            If ws.Name = "ShipmentsTally" Then
                Set tbl = ws.ListObjects("ShipmentsTally")
            Else
                Set tbl = ws.ListObjects("ReceivedTally")
            End If
            
            If Not tbl Is Nothing Then
                Dim itemsCol As Long, uomCol As Long, rowNumCol As Long
                
                ' Find column indexes
                For itemsCol = 1 To tbl.ListColumns.Count
                    If UCase(tbl.ListColumns(itemsCol).Name) = "ITEMS" Then
                        If gSelectedCell.Column = tbl.ListColumns(itemsCol).Range.Column Then
                            ' Found the ITEMS column and we're in it
                            currentRowIndex = gSelectedCell.Row - tbl.HeaderRowRange.Row
                            
                            ' If we have a valid row
                            If currentRowIndex > 0 Then
                                ' Find UOM column
                                For uomCol = 1 To tbl.ListColumns.Count
                                    If UCase(tbl.ListColumns(uomCol).Name) = "UOM" Then
                                        ' Get UOM using both item name and ROW#
                                        Dim itemUOM As String
                                        itemUOM = modTS_Data.GetItemUOMByRowNum(chosenRowNum, chosenItemCode, chosenValue)
                                        
                                        ' Set UOM
                                        tbl.DataBodyRange(currentRowIndex, uomCol).Value = itemUOM
                                        Exit For
                                    End If
                                Next uomCol
                                
                                ' Store ROW# in hidden column if it exists
                                For rowNumCol = 1 To tbl.ListColumns.Count
                                    If UCase(tbl.ListColumns(rowNumCol).Name) = "ROW#" Then
                                        tbl.DataBodyRange(currentRowIndex, rowNumCol).Value = chosenRowNum
                                        Exit For
                                    End If
                                Next rowNumCol
                            End If
                            Exit For
                        End If
                    End If
                Next itemsCol
            End If
        End If
        On Error GoTo 0
    End If
    
    isRunning = False
    Unload Me
End Sub

' Populate the list box with the full list of items.
Private Sub PopulateListBox(itemArray As Variant)
    Dim i As Long
    Me.lstBox.Clear
    
    ' Check if itemArray is properly initialized
    If IsEmpty(itemArray) Or Not IsArray(itemArray) Then Exit Sub
    
    For i = LBound(itemArray, 1) To UBound(itemArray, 1)
        Me.lstBox.AddItem ""
        ' Match the array indices with how data is loaded
        Me.lstBox.List(Me.lstBox.ListCount - 1, 0) = itemArray(i, 0)  ' ITEM_CODE
        Me.lstBox.List(Me.lstBox.ListCount - 1, 1) = itemArray(i, 1)  ' ROW#
        Me.lstBox.List(Me.lstBox.ListCount - 1, 2) = itemArray(i, 2)  ' ITEM name
        Me.lstBox.List(Me.lstBox.ListCount - 1, 3) = itemArray(i, 3)  ' LOCATION
    Next i
End Sub

' Helper function to update the description in txtBox2
Private Sub UpdateDescription()
    ' Clear existing description
    Me.txtBox2.text = ""
    
    ' If an item is selected in the main listbox
    If Me.lstBox.ListIndex <> -1 Then
        ' Get the selected index
        Dim selectedIndex As Integer
        selectedIndex = Me.lstBox.ListIndex
        
        ' Get the description for this item from the FullItemList
        ' Add 1 because ListBox is 0-based but array is 1-based
        If selectedIndex + 1 <= UBound(FullItemList, 1) Then
            ' Set the description text
            Me.txtBox2.text = FullItemList(selectedIndex + 1, 4)
        End If
    End If
End Sub

