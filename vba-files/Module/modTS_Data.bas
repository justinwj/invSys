Attribute VB_Name = "modTS_Data"
' ========================
' Module: modTS_Data
' ========================
Option Explicit

Public Function LoadItemList(Optional ByVal wb As Workbook, _
                             Optional ByVal wsName As String = "INVENTORY MANAGEMENT", _
                             Optional ByVal tblName As String = "invSys") As Variant
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dataRange As Range
    Dim result As Variant
    Dim rowCount As Long, colCount As Long
    Dim i As Long
    
    Debug.Print "LoadItemList: Starting function"
    
    ' If no workbook is specified, use ThisWorkbook
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    
    ' Get worksheet and table references from parameters
    Set ws = wb.Worksheets(wsName)
    Set tbl = ws.ListObjects(tblName)
    
    If tbl Is Nothing Then
        Debug.Print "LoadItemList: Table '" & tblName & "' not found"
        GoTo ErrorHandler
    End If
    
    ' Get row count (exit if empty)
    rowCount = tbl.ListRows.Count
    If rowCount = 0 Then
        Debug.Print "LoadItemList: No rows in table"
        GoTo ErrorHandler
    End If
    
    Debug.Print "LoadItemList: Found " & rowCount & " rows in table"
    
    ' Create result array with space for ROW, ITEM_CODE, ITEM, LOCATION, DESCRIPTION
    ReDim result(1 To rowCount, 0 To 4)
    
    ' Get column references
    Dim itemCodeCol As Integer, rowCol As Integer
    Dim itemCol As Integer, locCol As Integer
    Dim descCol As Integer
    
    On Error Resume Next
    itemCodeCol = tbl.ListColumns("ITEM_CODE").Index
    rowCol = tbl.ListColumns("ROW").Index
    itemCol = tbl.ListColumns("ITEM").Index
    locCol = tbl.ListColumns("LOCATION").Index
    descCol = tbl.ListColumns("DESCRIPTION").Index
    On Error GoTo ErrorHandler
    
    Debug.Print "LoadItemList: Column indexes - ITEM_CODE: " & itemCodeCol & _
                ", ROW: " & rowCol & ", ITEM: " & itemCol & _
                ", LOCATION: " & locCol & ", DESCRIPTION: " & descCol
    
    ' Check if we found the required columns
    If itemCodeCol = 0 Or rowCol = 0 Or itemCol = 0 Then
        Debug.Print "LoadItemList: Required columns missing"
        GoTo ErrorHandler
    End If
    
    ' Fill the result array
    For i = 1 To rowCount
        ' Make sure we're getting the right data in the right order
        result(i, 0) = tbl.DataBodyRange.Cells(i, rowCol).Value      ' ROW (index 0)
        result(i, 1) = tbl.DataBodyRange.Cells(i, itemCodeCol).Value ' ITEM_CODE (index 1)
        result(i, 2) = tbl.DataBodyRange.Cells(i, itemCol).Value     ' ITEM (index 2)
        
        ' For debugging - print what we're loading
        Debug.Print "LoadItemList: Row " & i & " - ROW=" & result(i, 0) & ", ITEM_CODE=" & result(i, 1) & ", ITEM=" & result(i, 2)
        
        ' LOCATION
        If locCol > 0 Then
            result(i, 3) = tbl.DataBodyRange.Cells(i, locCol).Value
        End If
        ' DESCRIPTION
        If descCol > 0 Then
            result(i, 4) = tbl.DataBodyRange.Cells(i, descCol).Value
        End If
    Next i
    
    Debug.Print "LoadItemList: Successfully loaded " & rowCount & " items"
    LoadItemList = result
    Exit Function
    
ErrorHandler:
    Debug.Print "LoadItemList: Error " & Err.Number & " - " & Err.Description
    LoadItemList = Empty
End Function

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

Public Sub GenerateRowNumbers()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rowNumCol As Long
    Dim i As Long
    Dim maxRowNum As Long
    Dim newCol As ListColumn  ' Add this variable for the new column
    
    Set ws = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
    Set tbl = ws.ListObjects("invSys")
    
    ' Find the ROW# column
    On Error Resume Next
    rowNumCol = tbl.ListColumns("ROW").Index
    On Error GoTo ErrorHandler
    
    If rowNumCol = 0 Then
        ' Add the column if it doesn't exist - FIX THIS LINE:
        Set newCol = tbl.ListColumns.Add
        newCol.Name = "ROW"
        rowNumCol = tbl.ListColumns("ROW").Index
    End If
    
    ' Find the highest existing row number
    maxRowNum = 0
    For i = 1 To tbl.ListRows.count
        If IsNumeric(tbl.DataBodyRange(i, rowNumCol).value) Then
            maxRowNum = Application.WorksheetFunction.Max(maxRowNum, tbl.DataBodyRange(i, rowNumCol).value)
        End If
    Next i
    
    ' Assign row numbers to any blank cells
    For i = 1 To tbl.ListRows.count
        If Trim(tbl.DataBodyRange(i, rowNumCol).value & "") = "" Then
            ' Increment and assign new row number
            maxRowNum = maxRowNum + 1
            tbl.DataBodyRange(i, rowNumCol).value = maxRowNum
        End If
    Next i
    
    MsgBox "Row numbers have been generated successfully.", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error generating row numbers: " & Err.Description, vbExclamation
End Sub

' Add these functions to help with form triggering via keyboard
Public Sub InitializeKeyboardHandlers()
    ' Set up key handlers for F2, Enter and Tab keys in Excel
    On Error Resume Next
    
    ' Clear any existing handlers
    Application.OnKey "{F2}"
    Application.OnKey "~"
    Application.OnKey "{TAB}"
    
    ' Set new handlers
    Application.OnKey "{F2}", "modTS_Data.TryShowItemSearchForm"
    Application.OnKey "~", "modTS_Data.TryShowItemSearchForm" ' ~ is the Enter key
    Application.OnKey "{TAB}", "modTS_Data.TryShowItemSearchForm" 
    
    Debug.Print "Keyboard handlers initialized"
    On Error GoTo 0
End Sub

' Checks if we're in an ITEMS column and shows the search form
Public Sub TryShowItemSearchForm()
    ' Check if the active cell is in an ITEMS column
    On Error Resume Next
    
    Debug.Print "TryShowItemSearchForm called for cell " & ActiveCell.Address
    
    ' Only proceed if we're on one of the tally sheets
    If ActiveSheet.Name <> "ShipmentsTally" And ActiveSheet.Name <> "ReceivedTally" Then Exit Sub
    
    ' Get the appropriate table
    Dim tbl As ListObject
    If ActiveSheet.Name = "ShipmentsTally" Then
        Set tbl = ActiveSheet.ListObjects("ShipmentsTally")
    Else
        Set tbl = ActiveSheet.ListObjects("ReceivedTally")
    End If
    
    If tbl Is Nothing Then Exit Sub
    
    ' Get the ITEMS column range
    Dim itemsColIndex As Long
    For itemsColIndex = 1 To tbl.ListColumns.Count
        If UCase(tbl.ListColumns(itemsColIndex).Name) = "ITEMS" Then Exit For
    Next itemsColIndex
    
    ' If we didn't find the ITEMS column
    If itemsColIndex > tbl.ListColumns.Count Then Exit Sub
    
    ' Check if the active cell is in the data area of the ITEMS column
    If ActiveCell.Column = tbl.ListColumns(itemsColIndex).Range.Column Then
        If ActiveCell.Row > tbl.HeaderRowRange.Row Then
            ' This is a cell in the ITEMS column - store it and show the form
            Debug.Print "OnKey handler showing form for " & ActiveSheet.Name & " cell " & ActiveCell.Address
            Set gSelectedCell = ActiveCell
            frmItemSearch.Show vbModeless
        End If
    End If
    
    On Error GoTo 0
End Sub

Public Sub ShowItemSearchForm()
    ' Check if the active cell is in an ITEMS column
    If IsInItemsColumn(ActiveCell) Then
        Set gSelectedCell = ActiveCell
        frmItemSearch.Show vbModeless
    End If
End Sub

Public Sub CheckForItemsColumn()
    ' Check if we're about to enter an ITEMS column cell
    On Error Resume Next
    If IsInItemsColumn(ActiveCell) Then
        Set gSelectedCell = ActiveCell
        ' Only show for empty cells to avoid interfering with normal usage
        If IsEmpty(ActiveCell.Value) Or Trim(ActiveCell.Value) = "" Then
            frmItemSearch.Show vbModeless
        End If
    End If
    On Error GoTo 0
End Sub

Public Function IsInItemsColumn(cell As Range) As Boolean
    ' Default return value
    IsInItemsColumn = False
    
    On Error Resume Next
    
    ' Check for ShipmentsTally
    If cell.Worksheet.Name = "ShipmentsTally" Then
        Dim shipTbl As ListObject
        Set shipTbl = cell.Worksheet.ListObjects("ShipmentsTally")
        If Not shipTbl Is Nothing Then
            If Not Intersect(cell, shipTbl.ListColumns("ITEMS").DataBodyRange) Is Nothing Then
                IsInItemsColumn = True
            End If
        End If
    End If
    
    ' Check for ReceivedTally
    If cell.Worksheet.Name = "ReceivedTally" Then
        Dim recvTbl As ListObject
        Set recvTbl = cell.Worksheet.ListObjects("ReceivedTally")
        If Not recvTbl Is Nothing Then
            If Not Intersect(cell, recvTbl.ListColumns("ITEMS").DataBodyRange) Is Nothing Then
                IsInItemsColumn = True
            End If
        End If
    End If
    
    On Error GoTo 0
End Function

' Call this from Workbook_Open to set up the keyboard handlers
Public Sub SetupAllHandlers()
    ' Clear table filters
    ClearTableFilters
    
    ' Initialize global variables
    modGlobals.InitializeGlobalVariables
    
    ' Add a hotkey for opening the form - F4 key
    On Error Resume Next
    Application.OnKey "{F4}", "modGlobals.OpenItemSearchForCurrentCell"
    On Error GoTo 0
    
    ' Add the big search buttons
    AddBigSearchButton
    
    ' Add right-click menu option - now this will work correctly
    modGlobals.AddExtendedRightClickMenu
End Sub

' Add buttons to tally sheets
Public Sub AddItemSearchButtons()
    On Error Resume Next
    
    ' Add button to ShipmentsTally
    If Not ThisWorkbook.Worksheets("ShipmentsTally") Is Nothing Then
        Dim shipBtn As Shape
        ThisWorkbook.Worksheets("ShipmentsTally").Shapes.Delete "btnItemSearch"
        Set shipBtn = ThisWorkbook.Worksheets("ShipmentsTally").Shapes.AddShape(msoShapeRoundedRectangle, 10, 10, 100, 30)
        With shipBtn
            .Name = "btnItemSearch"
            .TextFrame.Characters.Text = "Search Items"
            .OnAction = "modGlobals.ShowItemSearchForm"
        End With
    End If
    
    ' Add button to ReceivedTally
    If Not ThisWorkbook.Worksheets("ReceivedTally") Is Nothing Then
        Dim recvBtn As Shape
        ThisWorkbook.Worksheets("ReceivedTally").Shapes.Delete "btnItemSearch"
        Set recvBtn = ThisWorkbook.Worksheets("ReceivedTally").Shapes.AddShape(msoShapeRoundedRectangle, 10, 10, 100, 30)
        With recvBtn
            .Name = "btnItemSearch"
            .TextFrame.Characters.Text = "Search Items"
            .OnAction = "modGlobals.ShowItemSearchForm"
        End With
    End If
End Sub

' Add this function since it's being called but not defined
Public Sub ClearTableFilters()
    On Error Resume Next
    
    ' Clear filters on ShipmentsTally
    If Not ThisWorkbook.Worksheets("ShipmentsTally") Is Nothing Then
        If Not ThisWorkbook.Worksheets("ShipmentsTally").ListObjects("ShipmentsTally") Is Nothing Then
            ThisWorkbook.Worksheets("ShipmentsTally").ListObjects("ShipmentsTally").AutoFilter.ShowAllData
        End If
        
        If Not ThisWorkbook.Worksheets("ShipmentsTally").ListObjects("invSysData_Shipping") Is Nothing Then
            ThisWorkbook.Worksheets("ShipmentsTally").ListObjects("invSysData_Shipping").AutoFilter.ShowAllData
        End If
    End If
    
    ' Clear filters on ReceivedTally
    If Not ThisWorkbook.Worksheets("ReceivedTally") Is Nothing Then
        If Not ThisWorkbook.Worksheets("ReceivedTally").ListObjects("ReceivedTally") Is Nothing Then
            ThisWorkbook.Worksheets("ReceivedTally").ListObjects("ReceivedTally").AutoFilter.ShowAllData
        End If
        
        If Not ThisWorkbook.Worksheets("ReceivedTally").ListObjects("invSysData_Receiving") Is Nothing Then
            ThisWorkbook.Worksheets("ReceivedTally").ListObjects("invSysData_Receiving").AutoFilter.ShowAllData
        End If
    End If
    
    On Error GoTo 0
End Sub

' Add this function to monitor selection changes using Application.OnTime
Public Sub MonitorItemsSelection()
    ' Check if we're on a tally sheet
    If ActiveSheet.Name <> "ShipmentsTally" And ActiveSheet.Name <> "ReceivedTally" Then Exit Sub
    
    ' Exit if no cell is selected
    If ActiveCell Is Nothing Then Exit Sub
    
    ' Get sheet references
    Dim tbl As ListObject
    On Error Resume Next
    If ActiveSheet.Name = "ShipmentsTally" Then
        Set tbl = ActiveSheet.ListObjects("ShipmentsTally")
    Else
        Set tbl = ActiveSheet.ListObjects("ReceivedTally")
    End If
    
    If tbl Is Nothing Then Exit Sub
    
    ' Get the ITEMS column
    Dim itemsCol As ListColumn
    Set itemsCol = tbl.ListColumns("ITEMS")
    If itemsCol Is Nothing Then Exit Sub
    
    ' Check if we're in the data area of the ITEMS column
    If ActiveCell.Column = itemsCol.Range.Column And _
       ActiveCell.Row > tbl.HeaderRowRange.Row Then
        ' We're in an ITEMS cell
        Debug.Print "MonitorItemsSelection: detected ITEMS cell at " & ActiveCell.Address
        
        ' Show the form
        Set gSelectedCell = ActiveCell
        frmItemSearch.Show vbModeless
    End If
    
    ' Schedule next check
    Application.OnTime Now + TimeValue("00:00:01"), "modTS_Data.MonitorItemsSelection"
End Sub

Public Sub AddBigSearchButton()
    On Error Resume Next
    
    ' Add to ShipmentsTally
    If Not ThisWorkbook.Sheets("ShipmentsTally") Is Nothing Then
        ThisWorkbook.Sheets("ShipmentsTally").Shapes.Delete "BigSearchBtn"
        
        Dim shipBtn As Shape
        Set shipBtn = ThisWorkbook.Sheets("ShipmentsTally").Shapes.AddShape(msoShapeRoundedRectangle, 10, 10, 180, 40)
        With shipBtn
            .Name = "BigSearchBtn"
            .TextFrame.Characters.Text = "SEARCH ITEMS"
            .Fill.ForeColor.RGB = RGB(0, 112, 192)  ' Blue
            .Line.ForeColor.RGB = RGB(0, 0, 128)    ' Dark blue
            .TextFrame.Characters.Font.Color = RGB(255, 255, 255)  ' White
            .TextFrame.Characters.Font.Bold = True
            .TextFrame.Characters.Font.Size = 12
            .OnAction = "modGlobals.OpenItemSearchForCurrentCell"
        End With
    End If
    
    ' Add to ReceivedTally
    If Not ThisWorkbook.Sheets("ReceivedTally") Is Nothing Then
        ThisWorkbook.Sheets("ReceivedTally").Shapes.Delete "BigSearchBtn"
        
        Dim recvBtn As Shape
        Set recvBtn = ThisWorkbook.Sheets("ReceivedTally").Shapes.AddShape(msoShapeRoundedRectangle, 10, 10, 180, 40)
        With recvBtn
            .Name = "BigSearchBtn"
            .TextFrame.Characters.Text = "SEARCH ITEMS"
            .Fill.ForeColor.RGB = RGB(0, 112, 192)  ' Blue
            .Line.ForeColor.RGB = RGB(0, 0, 128)    ' Dark blue
            .TextFrame.Characters.Font.Color = RGB(255, 255, 255)  ' White
            .TextFrame.Characters.Font.Bold = True
            .TextFrame.Characters.Font.Size = 12
            .OnAction = "modGlobals.OpenItemSearchForCurrentCell"
        End With
    End If
End Sub







