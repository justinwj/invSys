Attribute VB_Name = "modGlobals"
'====================
' Modules: modGlobals
'====================
Option Explicit
Public gSelectedCell As Range

' Global flag to track if the timer is paused
Public gTimerPaused As Boolean

' Add this at the top of modGlobals
Public Function IsFormOpen(formName As String) As Boolean
    Dim i As Integer
    IsFormOpen = False
    
    ' Check all open UserForms
    For i = 0 To UserForms.Count - 1
        If UserForms(i).Name = formName Then
            IsFormOpen = True
            Exit Function
        End If
    Next i
End Function

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

' Function to pause the timer
Public Sub PauseTimer()
    gTimerPaused = True
    Debug.Print "Timer paused"
End Sub

' Function to resume the timer
Public Sub ResumeTimer()
    gTimerPaused = False
    Debug.Print "Timer resumed"
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

    ' Add this function to lookup UOM by item name
Public Function GetItemUOM(itemName As String) As String
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim itemCol As Range, uomCol As Range
    Dim foundCell As Range
    Dim foundRow As Long
    
    ' Default return value if not found
    GetItemUOM = "each"
    
    ' Check if itemName is empty
    If Len(Trim(itemName)) = 0 Then Exit Function
    
    Set ws = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
    Set tbl = ws.ListObjects("invSys")
    
    Set itemCol = tbl.ListColumns("ITEM").DataBodyRange
    Set uomCol = tbl.ListColumns("UOM").DataBodyRange
    
    ' Find the item in the invSys table
    Set foundCell = itemCol.Find(What:=itemName, _
                                 LookIn:=xlValues, _
                                 LookAt:=xlWhole, _
                                 SearchOrder:=xlByRows, _
                                 MatchCase:=False)
    
    ' If found, return its UOM
    If Not foundCell Is Nothing Then
        foundRow = foundCell.row - itemCol.row + 1
        GetItemUOM = uomCol.Cells(foundRow, 1).value
        
        ' If UOM is empty, return default
        If Trim(GetItemUOM) = "" Then
            GetItemUOM = "each"
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    ' Log the error for debugging
    Debug.Print "Error in GetItemUOM: " & Err.Description
    ' Ensure a default value is returned on error
    GetItemUOM = "each"
End Function

' Function to lookup UOM by ITEM_CODE (preferred) or item name (fallback)
Public Function GetItemUOMByCode(ItemCode As String, itemName As String) As String
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim foundCell As Range
    Dim foundRow As Long
    
    ' Default return value if not found
    GetItemUOMByCode = "each"
    
    Set ws = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
    Set tbl = ws.ListObjects("invSys")
    
    ' First try finding by ITEM_CODE if provided
    If Trim(ItemCode) <> "" Then
        Set foundCell = tbl.ListColumns("ITEM_CODE").DataBodyRange.Find( _
                        What:=ItemCode, _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        MatchCase:=False)
                        
        If Not foundCell Is Nothing Then
            foundRow = foundCell.row - tbl.HeaderRowRange.row
            GetItemUOMByCode = tbl.ListColumns("UOM").DataBodyRange(foundRow).value
            
            ' If UOM is empty, return default
            If Trim(GetItemUOMByCode) = "" Then
                GetItemUOMByCode = "each"
            End If
            
            ' Found by code, return early
            Exit Function
        End If
    End If
    
    ' Fallback: Find by item name if code search failed or no code provided
    If Trim(itemName) <> "" Then
        Set foundCell = tbl.ListColumns("ITEM").DataBodyRange.Find( _
                        What:=itemName, _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        MatchCase:=False)
                        
        If Not foundCell Is Nothing Then
            foundRow = foundCell.row - tbl.HeaderRowRange.row
            GetItemUOMByCode = tbl.ListColumns("UOM").DataBodyRange(foundRow).value
            
            ' If UOM is empty, return default
            If Trim(GetItemUOMByCode) = "" Then
                GetItemUOMByCode = "each"
            End If
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in GetItemUOMByCode: " & Err.Description
    GetItemUOMByCode = "each"
End Function

Public Function GetItemUOMByRowNum(rowNum As String, ItemCode As String, itemName As String) As String
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim foundCell As Range
    Dim foundRow As Long
    
    ' Default return value if not found
    GetItemUOMByRowNum = "each"
    
    Set ws = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
    Set tbl = ws.ListObjects("invSys")
    
    ' Try to find the item by ROW# first (most precise)
    If Trim(rowNum) <> "" Then
        Set foundCell = tbl.ListColumns("ROW").DataBodyRange.Find( _
                        What:=rowNum, _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        MatchCase:=False)
                        
        If Not foundCell Is Nothing Then
            foundRow = foundCell.row - tbl.HeaderRowRange.row
            GetItemUOMByRowNum = tbl.DataBodyRange(foundRow, tbl.ListColumns("UOM").Index).value
            
            ' If UOM is empty, return default
            If Trim(GetItemUOMByRowNum) = "" Then GetItemUOMByRowNum = "each"
            Exit Function
        End If
    End If
    
    ' Try ITEM_CODE next
    If Trim(ItemCode) <> "" Then
        Set foundCell = tbl.ListColumns("ITEM_CODE").DataBodyRange.Find( _
                        What:=ItemCode, _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        MatchCase:=False)
                        
        If Not foundCell Is Nothing Then
            foundRow = foundCell.row - tbl.HeaderRowRange.row
            GetItemUOMByRowNum = tbl.DataBodyRange(foundRow, tbl.ListColumns("UOM").Index).value
            
            ' If UOM is empty, return default
            If Trim(GetItemUOMByRowNum) = "" Then GetItemUOMByRowNum = "each"
            Exit Function
        End If
    End If
    
    ' Last resort: Try item name
    If Trim(itemName) <> "" Then
        Set foundCell = tbl.ListColumns("ITEM").DataBodyRange.Find( _
                        What:=itemName, _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        MatchCase:=False)
                        
        If Not foundCell Is Nothing Then
            foundRow = foundCell.row - tbl.HeaderRowRange.row
            GetItemUOMByRowNum = tbl.DataBodyRange(foundRow, tbl.ListColumns("UOM").Index).value
            
            ' If UOM is empty, return default
            If Trim(GetItemUOMByRowNum) = "" Then GetItemUOMByRowNum = "each"
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in GetItemUOMByRowNum: " & Err.Description
    GetItemUOMByRowNum = "each"
End Function



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

Public Sub OpenItemSearchForCurrentCell()
    ' Pause the timer to prevent conflicts
    PauseTimer
    
    ' Store the active cell as the selected cell
    Set gSelectedCell = ActiveCell
    
    ' Show the form
    frmItemSearch.Show vbModeless
    
    ' Resume the timer after showing the form
    ResumeTimer
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

Public Sub TestFormOpening()
    ' Clear any existing forms
    On Error Resume Next
    Unload frmItemSearch
    On Error GoTo 0
    
    ' Direct test of form initialization
    frmItemSearch.Show vbModeless
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







