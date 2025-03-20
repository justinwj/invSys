Attribute VB_Name = "modGlobals"
'====================
' Modules: modGlobals
'====================
Option Explicit
Public gSelectedCell As Range

Public Sub CommitSelectionAndCloseWrapper()
    frmItemSearch.CommitSelectionAndClose
End Sub

' Add this function to initialize global variables
Public Sub InitializeGlobalVariables()
    ' Make sure the gSelectedCell variable is available
    On Error Resume Next
    Set gSelectedCell = Nothing
    On Error GoTo 0
End Sub

' Direct method to show the form - can be called from anywhere
Public Sub ShowItemSearchForm()
    If ActiveCell Is Nothing Then Exit Sub
    
    ' Check if we're in a tally sheet and ITEMS column
    Dim validCell As Boolean
    validCell = False
    
    On Error Resume Next
    If ActiveSheet.Name = "ShipmentsTally" Then
        Dim shipTbl As ListObject
        Set shipTbl = ActiveSheet.ListObjects("ShipmentsTally")
        If Not shipTbl Is Nothing Then
            If Not Intersect(ActiveCell, shipTbl.ListColumns("ITEMS").DataBodyRange) Is Nothing Then
                validCell = True
            End If
        End If
    ElseIf ActiveSheet.Name = "ReceivedTally" Then
        Dim recvTbl As ListObject
        Set recvTbl = ActiveSheet.ListObjects("ReceivedTally")
        If Not recvTbl Is Nothing Then
            If Not Intersect(ActiveCell, recvTbl.ListColumns("ITEMS").DataBodyRange) Is Nothing Then
                validCell = True
            End If
        End If
    End If
    
    If validCell Then
        Set gSelectedCell = ActiveCell
        frmItemSearch.Show vbModeless
    End If
End Sub

' Test function to verify worksheet events are working
Public Sub TestWorksheetEvents()
    If ActiveSheet.Name <> "ShipmentsTally" And ActiveSheet.Name <> "ReceivedTally" Then
        MsgBox "Please activate a tally sheet first", vbInformation
        Exit Sub
    End If
    
    MsgBox "Event testing:" & vbCrLf & _
           "1. Excel events enabled: " & Application.EnableEvents & vbCrLf & _
           "2. Active sheet: " & ActiveSheet.Name & vbCrLf & _
           "3. After clicking OK, select a cell in the ITEMS column", vbInformation
    
    ' Force a re-initialization of the event handlers
    Application.EnableEvents = True
End Sub

' Add this to test if worksheet events are firing
Public Sub TestEventHandlers()
    MsgBox "Events enabled: " & Application.EnableEvents
    
    ' Test sheet names
    Dim receivedExists As Boolean, shipmentsExists As Boolean
    
    On Error Resume Next
    receivedExists = (ThisWorkbook.Sheets("ReceivedTally").Name <> "")
    shipmentsExists = (ThisWorkbook.Sheets("ShipmentsTally").Name <> "")
    On Error GoTo 0
    
    MsgBox "Sheet check:" & vbCrLf & _
           "ReceivedTally exists: " & receivedExists & vbCrLf & _
           "ShipmentsTally exists: " & shipmentsExists
    
    ' Force enable events
    Application.EnableEvents = True
    
    MsgBox "Events have been enabled. Try clicking in ITEMS column now."
End Sub

Public Sub AddRightClickMenu()
    On Error Resume Next
    
    ' Remove existing menu if it exists
    Application.CommandBars("Cell").Controls("Search Items").Delete
    
    ' Add the menu item
    Dim menuItem As CommandBarButton
    Set menuItem = Application.CommandBars("Cell").Controls.Add(Type:=msoControlButton, temporary:=True)
    With menuItem
        .Caption = "Search Items"
        .OnAction = "modGlobals.ShowItemSearchForm"
        .Style = msoButtonCaption
    End With
    
    On Error GoTo 0
End Sub

Public Sub OpenItemSearchForCurrentCell()
    ' Store the active cell as the selected cell
    Set gSelectedCell = ActiveCell
    
    ' Show the search form
    frmItemSearch.Show vbModeless
End Sub

' Create direct functions to open forms for each sheet
Public Sub OpenSearchInShipmentsTally()
    If ActiveSheet.Name <> "ShipmentsTally" Then
        ThisWorkbook.Sheets("ShipmentsTally").Activate
    End If
    
    OpenItemSearchForCurrentCell
End Sub

Public Sub OpenSearchInReceivedTally()
    If ActiveSheet.Name <> "ReceivedTally" Then
        ThisWorkbook.Sheets("ReceivedTally").Activate
    End If
    
    OpenItemSearchForCurrentCell
End Sub

Public Sub AddExtendedRightClickMenu()
    On Error Resume Next
    
    ' Remove existing menu items if they exist
    Application.CommandBars("Cell").Controls("Search Items").Delete
    
    ' Add menu item
    Dim itemMenu1 As CommandBarButton
    Set itemMenu1 = Application.CommandBars("Cell").Controls.Add(Type:=msoControlButton)
    With itemMenu1
        .Caption = "Search Items (Current Cell)"
        .OnAction = "modGlobals.OpenItemSearchForCurrentCell"
        .BeginGroup = True
    End With
    
    On Error GoTo 0
End Sub

' Ensure a form is visible on screen
Public Sub EnsureFormVisible(frm As Object)
    On Error Resume Next
    
    ' Get screen dimensions
    Dim screenWidth As Long, screenHeight As Long
    screenWidth = Application.Width
    screenHeight = Application.Height
    
    ' Calculate the visible area within Excel
    Dim visibleLeft As Long, visibleTop As Long
    visibleLeft = Application.Left
    visibleTop = Application.Top
    
    ' Set minimum visible margins
    Const MIN_VISIBLE As Long = 50
    
    ' Check if form is too far to the right
    If frm.Left > (visibleLeft + screenWidth - MIN_VISIBLE) Then
        frm.Left = visibleLeft + (screenWidth / 2) - (frm.Width / 2)
    End If
    
    ' Check if form is too far to the left
    If frm.Left < (visibleLeft + MIN_VISIBLE - frm.Width) Then
        frm.Left = visibleLeft + MIN_VISIBLE
    End If
    
    ' Check if form is too far down
    If frm.Top > (visibleTop + screenHeight - MIN_VISIBLE) Then
        frm.Top = visibleTop + (screenHeight / 2) - (frm.Height / 2)
    End If
    
    ' Check if form is too far up
    If frm.Top < (visibleTop + MIN_VISIBLE - frm.Height) Then
        frm.Top = visibleTop + MIN_VISIBLE
    End If
    
    On Error GoTo 0
End Sub







