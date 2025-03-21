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
    
    ' Add the big search buttons (these work well)
    modTS_Data.AddBigSearchButton
    
    ' Show user instructions
    MsgBox "To search for items:" & vbCrLf & _
           "1. Click the blue 'SEARCH ITEMS' button" & vbCrLf & _
           "2. Press F3 to open the search form with the current cell" & vbCrLf & _
           "3. Try clicking directly in ITEMS cells", _
           vbInformation, "Item Search Available"
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    ' Re-enable events when activating sheets
    Application.EnableEvents = True
End Sub