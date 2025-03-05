Attribute VB_Name = "modUserAuth"
Sub SignOut_Click()
    ' Call HandleSignOut directly from modUserAuth
    HandleSignOut
End Sub

Public Sub HandleSignOut()
    Dim ws As Worksheet

    ' Check if the workbook is closing
    If Application.Workbooks.count = 0 Then Exit Sub

    ' Set reference to the inventory sheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
    On Error GoTo 0

    ' If sheet not found, exit to avoid errors
    If ws Is Nothing Then Exit Sub

    ' Lock inventory sheet to prevent unauthorized edits
    ' ws.Protect password:="yourPIN", UserInterfaceOnly:=True, DrawingObjects:=False

    ' Prevent reopening frmLogin if the workbook is closing
    If Not Application.Workbooks.count = 0 Then
        frmLogin.txtUsername.Value = ""
        frmLogin.txtPIN.Value = ""
        frmLogin.Show
    End If
End Sub
Public Sub HideLoginForm()
    Unload frmLogin
End Sub
Public Sub LoadRolesIntoComboBox(cmb As ComboBox)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim roleCell As Range
    
    ' Set reference to tblRoles table in UserCredentials sheet
    Set ws = ThisWorkbook.Sheets("UserCredentials")
    Set tbl = ws.ListObjects("tblRoles")

    ' Clear existing values
    cmb.Clear

    ' Populate dropdown with role names
    For Each roleCell In tbl.ListColumns("Roles").DataBodyRange
        cmb.AddItem roleCell.Value
    Next roleCell
End Sub

Public Sub LoadUsersIntoComboBox(cmb As ComboBox)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim userCell As Range
    
    ' Set reference to UserCredentials table
    Set ws = ThisWorkbook.Sheets("UserCredentials")
    Set tbl = ws.ListObjects("UserCredentials")

    ' Clear existing values
    cmb.Clear

    ' Populate dropdown with usernames
    For Each userCell In tbl.ListColumns("USERNAME").DataBodyRange
        cmb.AddItem userCell.Value
    Next userCell
End Sub

