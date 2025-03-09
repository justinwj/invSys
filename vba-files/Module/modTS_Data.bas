Attribute VB_Name = "modTS_Data"
' ========================
' Module: modTS_Data
' ========================
Option Explicit

Public Function LoadItemList() As Variant
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dataRange As Range
    Dim result As Variant
    Dim rowCount As Long, colCount As Long
    Dim i As Long, j As Long
    
    Set ws = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
    Set tbl = ws.ListObjects("invSys")
    
    ' Get the number of items and initialize result array
    rowCount = tbl.ListRows.Count
    If rowCount = 0 Then
        LoadItemList = Empty
        Exit Function
    End If
    
    ' Create result array with space for ITEM_CODE, ROW#, ITEM, and LOCATION
    ReDim result(1 To rowCount, 0 To 3)
    
    ' Get column indexes
    Dim itemCodeCol As Long, rowNumCol As Long, itemCol As Long, locCol As Long
    
    On Error Resume Next
    itemCodeCol = tbl.ListColumns("ITEM_CODE").Index
    rowNumCol = tbl.ListColumns("ROW#").Index
    itemCol = tbl.ListColumns("ITEM").Index
    locCol = tbl.ListColumns("LOCATION").Index
    On Error GoTo ErrorHandler
    
    ' If location column doesn't exist, use empty string
    Dim hasLocationCol As Boolean
    hasLocationCol = (locCol > 0)
    
    ' Check if ROW# column exists
    If rowNumCol = 0 Then
        MsgBox "The 'ROW#' column was not found in the invSys table." & vbCrLf & _
               "Please add this column to uniquely identify inventory rows.", vbExclamation
        LoadItemList = Empty
        Exit Function
    End If
    
    ' Fill the result array
    For i = 1 To rowCount
        ' ITEM_CODE (Column 0)
        result(i, 0) = tbl.DataBodyRange(i, itemCodeCol).Value
        
        ' ROW# (Column 1)
        result(i, 1) = tbl.DataBodyRange(i, rowNumCol).Value
        
        ' ITEM name (Column 2)
        result(i, 2) = tbl.DataBodyRange(i, itemCol).Value
        
        ' LOCATION (Column 3)
        If hasLocationCol Then
            result(i, 3) = tbl.DataBodyRange(i, locCol).Value
        Else
            result(i, 3) = ""
        End If
    Next i
    
    LoadItemList = result
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in LoadItemList: " & Err.Description
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
        GetItemUOM = uomCol.Cells(foundRow, 1).Value
        
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
Public Function GetItemUOMByCode(itemCode As String, itemName As String) As String
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
    If Trim(itemCode) <> "" Then
        Set foundCell = tbl.ListColumns("ITEM_CODE").DataBodyRange.Find( _
                        What:=itemCode, _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        MatchCase:=False)
                        
        If Not foundCell Is Nothing Then
            foundRow = foundCell.Row - tbl.HeaderRowRange.Row
            GetItemUOMByCode = tbl.ListColumns("UOM").DataBodyRange(foundRow).Value
            
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
            foundRow = foundCell.Row - tbl.HeaderRowRange.Row
            GetItemUOMByCode = tbl.ListColumns("UOM").DataBodyRange(foundRow).Value
            
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

Public Function GetItemUOMByRowNum(rowNum As String, itemCode As String, itemName As String) As String
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
        Set foundCell = tbl.ListColumns("ROW#").DataBodyRange.Find( _
                        What:=rowNum, _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        MatchCase:=False)
                        
        If Not foundCell Is Nothing Then
            foundRow = foundCell.Row - tbl.HeaderRowRange.Row
            GetItemUOMByRowNum = tbl.DataBodyRange(foundRow, tbl.ListColumns("UOM").Index).Value
            
            ' If UOM is empty, return default
            If Trim(GetItemUOMByRowNum) = "" Then GetItemUOMByRowNum = "each"
            Exit Function
        End If
    End If
    
    ' Try ITEM_CODE next
    If Trim(itemCode) <> "" Then
        Set foundCell = tbl.ListColumns("ITEM_CODE").DataBodyRange.Find( _
                        What:=itemCode, _
                        LookIn:=xlValues, _
                        LookAt:=xlWhole, _
                        MatchCase:=False)
                        
        If Not foundCell Is Nothing Then
            foundRow = foundCell.Row - tbl.HeaderRowRange.Row
            GetItemUOMByRowNum = tbl.DataBodyRange(foundRow, tbl.ListColumns("UOM").Index).Value
            
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
            foundRow = foundCell.Row - tbl.HeaderRowRange.Row
            GetItemUOMByRowNum = tbl.DataBodyRange(foundRow, tbl.ListColumns("UOM").Index).Value
            
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
    rowNumCol = tbl.ListColumns("ROW#").Index
    On Error GoTo ErrorHandler
    
    If rowNumCol = 0 Then
        ' Add the column if it doesn't exist - FIX THIS LINE:
        Set newCol = tbl.ListColumns.Add
        newCol.Name = "ROW#"
        rowNumCol = tbl.ListColumns("ROW#").Index
    End If
    
    ' Find the highest existing row number
    maxRowNum = 0
    For i = 1 To tbl.ListRows.Count
        If IsNumeric(tbl.DataBodyRange(i, rowNumCol).Value) Then
            maxRowNum = Application.WorksheetFunction.Max(maxRowNum, tbl.DataBodyRange(i, rowNumCol).Value)
        End If
    Next i
    
    ' Assign row numbers to any blank cells
    For i = 1 To tbl.ListRows.Count
        If Trim(tbl.DataBodyRange(i, rowNumCol).Value & "") = "" Then
            ' Increment and assign new row number
            maxRowNum = maxRowNum + 1
            tbl.DataBodyRange(i, rowNumCol).Value = maxRowNum
        End If
    Next i
    
    MsgBox "Row numbers have been generated successfully.", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error generating row numbers: " & Err.Description, vbExclamation
End Sub







