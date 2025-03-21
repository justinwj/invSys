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
    
    ' Create a backup keyboard shortcut for when events don't work
    On Error Resume Next
    Application.OnKey "{F3}", "modGlobals.OpenItemSearchForCurrentCell"
    On Error GoTo 0
    
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
    
    ' Force reset of event handlers
    If Sh.Name = "ShipmentsTally" Or Sh.Name = "ReceivedTally" Then
        ' Force event reconnection
        Dim tempSetting As Boolean
        tempSetting = Application.EnableEvents
        Application.EnableEvents = False
        Application.EnableEvents = True
        Application.EnableEvents = tempSetting
        
        Debug.Print "Sheet " & Sh.Name & " activated - events reset"
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

Public Sub CheckActiveCell()
    On Error Resume Next
    
    ' Skip processing if timer is paused
    If gTimerPaused Then
        Debug.Print "CheckActiveCell: Timer is paused, skipping checks"
        GoTo ScheduleNext
    End If
    
    ' Only run if we're on a tally sheet and no form is already showing
    If ActiveSheet.Name = "ShipmentsTally" Or ActiveSheet.Name = "ReceivedTally" Then
        ' Don't show the form if it's already visible
        If Not IsFormLoaded("frmItemSearch") Then
            ' Rest of your existing code...
        End If
    End If

ScheduleNext:
    ' Schedule next check
    Application.OnTime Now + TimeSerial(0, 0, 1), "CheckActiveCell"
End Sub

' Helper function to check if a form is already loaded
Function IsFormLoaded(formName As String) As Boolean
    Dim frm As Object
    IsFormLoaded = False
    
    On Error Resume Next
    For Each frm In UserForms
        If frm.Name = formName Then
            IsFormLoaded = True
            Exit For
        End If
    Next frm
    On Error GoTo 0
End Function