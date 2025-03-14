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
    ' Make sure Excel events are enabled
    Application.EnableEvents = True

    ' Set up all handlers
    modTS_Data.SetupAllHandlers
    
    ' Add the big search buttons
    modTS_Data.AddBigSearchButton
    
    ' Display message to user about using F4 key
    MsgBox "To search for items:" & vbCrLf & _
           "1. Click the blue 'SEARCH ITEMS' button" & vbCrLf & _
           "2. Or press F4 when cursor is in ITEMS column" & vbCrLf & _
           "3. Or double-click in the ITEMS column", _
           vbInformation, "Item Search Available"
End Sub

' Add these event handlers to ensure sheet events work
Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    ' When a tally sheet is activated, ensure events are enabled
    If Sh.Name = "ShipmentsTally" Or Sh.Name = "ReceivedTally" Then
        ' Force events enabled
        Application.EnableEvents = True
        
        ' Debug message
        Debug.Print "Sheet activated: " & Sh.Name & ", Events enabled: " & Application.EnableEvents
    End If
End Sub

Sub CheckForHiddenPasswords()
    Dim nm As Name
    For Each nm In ThisWorkbook.Names
        Debug.Print nm.Name & " = " & nm.RefersTo
    Next nm
End Sub

Public Sub InitializeGlobalVariables()
    ' Make sure the gSelectedCell variable is available
    On Error Resume Next
    Set gSelectedCell = Nothing
    On Error GoTo 0
End Sub