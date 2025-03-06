VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    Dim ws As Worksheet

    ' Unprotect all sheets on startup except for frmLogin control
    For Each ws In ThisWorkbook.Sheets
        On Error Resume Next
        ws.Unprotect ' No password needed if we don't know it
        On Error GoTo 0
    Next ws

    ' Force login form to lock access
    ' frmLogin.Show
End Sub
Sub CheckForHiddenPasswords()
    Dim nm As Name
    For Each nm In ThisWorkbook.Names
        Debug.Print nm.Name & " = " & nm.RefersTo
    Next nm
End Sub





