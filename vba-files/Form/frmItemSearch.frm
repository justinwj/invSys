VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmItemSearch 
   Caption         =   "Item Search"
   ClientHeight    =   3645
   ClientLeft      =   120
   ClientTop       =   470
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
    ' Ensure the text box is active for immediate typing.
    Me.txtBox.SetFocus
End Sub

Private Sub UserForm_Initialize()
    ' Load the full list of items from the invSys table.
    FullItemList = modTS_Data.LoadItemList()
    ' Populate lstBox with the full list.
    Call PopulateListBox(FullItemList)
    
    ' Pre-populate txtBox with the current cell value (if any) and select all text.
    If Not gSelectedCell Is Nothing Then
        If Not IsEmpty(gSelectedCell.Value) Then
            Me.txtBox.Text = CStr(gSelectedCell.Value)
            Me.txtBox.SelStart = 0
            Me.txtBox.SelLength = Len(Me.txtBox.Text)
        End If
    End If
    
    ' Attempt to match the current txtBox content.
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
    If Me.lstBox.ListIndex <> -1 Then
        Me.txtBox.Text = Me.lstBox.Value
        Me.txtBox.SelStart = 0
        Me.txtBox.SelLength = Len(Me.txtBox.Text)
    End If
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
    If Trim(Me.txtBox.Text) = "" Then
        If Not gSelectedCell Is Nothing Then gSelectedCell.ClearContents
    Else
        If Me.lstBox.ListIndex <> -1 Then
            gSelectedCell.Value = Me.lstBox.Value
        Else
            gSelectedCell.Value = Me.txtBox.Text
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
