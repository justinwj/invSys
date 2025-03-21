Attribute VB_Name = "modTS_Tally"
' ================================================
' Module: modTS_Tally (TS stands for Tally System)
' ================================================
Option Explicit
' This module is responsible for tallying orders and displaying them in a user form.

' Track if we're already running a tally operation
Private isRunningTally As Boolean

Sub TallyItems(sheetName As String, tableName As String, formToShow As Object)
    ' Debug at beginning
    Debug.Print "Starting TallyItems with: " & sheetName & ", " & tableName & ", " & TypeName(formToShow)
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dict As Object
    Dim i As Long
    Dim key As Variant
    Dim item As Variant, quantity As Double, uom As Variant
    Dim ItemCode As String, rowNum As String
    Dim lb As MSForms.ListBox
    Dim keyParts As Variant
    Dim ctrl As MSForms.Control  ' Add this declaration for the ctrl variable
    
    ' Error checking for the form
    If formToShow Is Nothing Then
        MsgBox "Error: Form reference is null", vbExclamation
        Exit Sub
    End If
    
    ' Find the ListBox in the form
    On Error Resume Next
    Set lb = Nothing
    For Each ctrl In formToShow.Controls
        If TypeName(ctrl) = "ListBox" Then
            Set lb = ctrl
            Debug.Print "Found ListBox with name: " & ctrl.Name
            Exit For
        End If
    Next ctrl
    
    If lb Is Nothing Then
        MsgBox "Error: The form doesn't have a ListBox control", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Continue with the rest of the function...
    Set ws = ThisWorkbook.Sheets(sheetName)
    Set tbl = ws.ListObjects(tableName)
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    
    ' Check if we have ROW and ITEM_CODE columns in the source table
    Dim hasRowCol As Boolean, hasItemCodeCol As Boolean
    Dim rowColIndex As Integer, itemCodeColIndex As Integer
    
    On Error Resume Next
    For i = 1 To tbl.ListColumns.count
        If UCase(tbl.ListColumns(i).Name) = "ROW" Then
            hasRowCol = True
            rowColIndex = i
        ElseIf UCase(tbl.ListColumns(i).Name) = "ITEM_CODE" Then
            hasItemCodeCol = True
            itemCodeColIndex = i
        End If
    Next i
    On Error GoTo 0
    
    ' Process all rows in the table
    For i = 1 To tbl.ListRows.count
        ' Get raw cell values
        item = tbl.ListColumns("ITEMS").DataBodyRange(i, 1).value
        quantity = tbl.ListColumns("QUANTITY").DataBodyRange(i, 1).value
        uom = tbl.ListColumns("UOM").DataBodyRange(i, 1).value
        
        ' Get ROW and ITEM_CODE if available
        rowNum = ""
        ItemCode = ""
        
        If hasRowCol Then rowNum = tbl.DataBodyRange(i, rowColIndex).value
        If hasItemCodeCol Then ItemCode = tbl.DataBodyRange(i, itemCodeColIndex).value
        
        ' Skip rows where the item is empty or quantity is zero/empty
        If Trim(CStr(item)) <> "" And quantity > 0 Then
            ' Create a unique key that includes ROW if available
            Dim uniqueKey As String
            If rowNum <> "" Then
                ' Use ROW for uniqueness (most specific)
                uniqueKey = "ROW_" & rowNum
            ElseIf ItemCode <> "" Then
                ' Use ITEM_CODE as fallback
                uniqueKey = "CODE_" & ItemCode
            Else
                ' Use item name and UOM as last resort
                uniqueKey = "NAME_" & LCase(Trim(CStr(item))) & "|" & LCase(Trim(CStr(uom)))
            End If
            
            ' Tally items using the unique key
            If dict.Exists(uniqueKey) Then
                dict(uniqueKey) = dict(uniqueKey) + quantity
            Else
                dict.Add uniqueKey, quantity
                ' Store reference information
                dict.Add "info_" & uniqueKey, Array(item, ItemCode, rowNum, uom)
            End If
        End If
    Next i
    
    ' Display the tally in the list box
    lb.Clear
    lb.ColumnCount = 5 ' ITEM, QTY, UOM, ITEM_CODE, ROW
    lb.ColumnWidths = "150;50;50;0;0" ' Hide ITEM_CODE and ROW columns
    
    ' Add header row
    lb.AddItem "ITEMS"
    lb.List(0, 1) = "QUANTITY"
    lb.List(0, 2) = "UOM"
    lb.List(0, 3) = "ITEM_CODE"
    lb.List(0, 4) = "ROW"
    
    ' Add data rows
    If dict.count > 0 Then
        For Each key In dict.Keys
            If Left$(key, 5) <> "info_" Then
                Dim itemInfo As Variant
                itemInfo = dict("info_" & key)
                
                lb.AddItem itemInfo(0) ' Item name
                lb.List(lb.ListCount - 1, 1) = dict(key)   ' Quantity
                lb.List(lb.ListCount - 1, 2) = itemInfo(3) ' UOM
                lb.List(lb.ListCount - 1, 3) = itemInfo(1) ' ITEM_CODE
                lb.List(lb.ListCount - 1, 4) = itemInfo(2) ' ROW
            End If
        Next key
        formToShow.Show
    Else
        MsgBox "No valid items found to tally.", vbInformation
    End If
End Sub

' Helper function to normalize text
Private Function NormalizeText(text As String) As String
    ' Trim and convert to lowercase for consistent matching
    Dim result As String
    result = Trim(text)
    NormalizeText = LCase(result)
End Function

Sub TallyShipments()
    ' Create and show form with shipments data
    Dim frm As frmShipmentsTally
    Set frm = New frmShipmentsTally
    
    ' Make sure the form has required controls
    If Not FormHasRequiredControls(frm) Then
        MsgBox "The form is missing required controls.", vbCritical
        Exit Sub
    End If
    
    ' Configure the form
    With frm
        ' Make sure the listbox exists and is configured properly
        .lstBox.Clear
        .lstBox.ColumnCount = 3
        .lstBox.ColumnWidths = "150;50;80" ' Adjust as needed
        .lstBox.AddItem "ITEMS"
        .lstBox.List(0, 1) = "QUANTITY"
        .lstBox.List(0, 2) = "UOM"
    End With
    
    ' Populate the form
    PopulateShipmentsForm frm
    
    ' Show the form if there are items
    If frm.lstBox.ListCount > 1 Then ' More than just the header row
        frm.Show vbModal
    Else
        MsgBox "No shipments to tally", vbInformation
    End If
End Sub

Function FormHasRequiredControls(frm As Object) As Boolean
    On Error Resume Next
    FormHasRequiredControls = Not (frm.lstBox Is Nothing)
    On Error GoTo 0
End Function

Sub PopulateShipmentsForm(frm As frmShipmentsTally)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dict As Object
    Dim i As Long
    Dim j As Long
    Dim key As Variant
    Dim itemInfo As Variant
    
    ' Get worksheet and table references
    Set ws = ThisWorkbook.Sheets("ShipmentsTally")
    Set tbl = ws.ListObjects("ShipmentsTally")
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    
    ' Process and tally items from the table
    For i = 1 To tbl.ListRows.count
        ' Get basic values with error handling
        Dim item As String, quantity As Double, uom As String
        On Error Resume Next
        item = CStr(tbl.DataBodyRange(i, tbl.ListColumns("ITEMS").Index).value)
        
        ' Be extra careful with quantity conversion
        Dim rawQuantity As Variant
        rawQuantity = tbl.DataBodyRange(i, tbl.ListColumns("QUANTITY").Index).value
        If IsNumeric(rawQuantity) Then
            quantity = CDbl(rawQuantity)
        Else
            quantity = 0
        End If
        
        uom = CStr(tbl.DataBodyRange(i, tbl.ListColumns("UOM").Index).value)
        On Error GoTo ErrorHandler
        
        ' Skip empty rows or rows with zero quantity
        If Trim(item) <> "" And quantity > 0 Then
            ' Extract ROW and ITEM_CODE if available
            Dim rowNum As String, ItemCode As String
            rowNum = ""
            ItemCode = ""
            
            On Error Resume Next
            ' Check if ROW and ITEM_CODE are in columns
            For j = 1 To tbl.ListColumns.count
                If UCase(tbl.ListColumns(j).Name) = "ROW" Then
                    rowNum = CStr(tbl.DataBodyRange(i, j).value)
                ElseIf UCase(tbl.ListColumns(j).Name) = "ITEM_CODE" Then
                    ItemCode = CStr(tbl.DataBodyRange(i, j).value)
                End If
            Next j
            
            ' If we don't have a ROW yet, look up the item in inventory
            If rowNum = "" Then
                Dim invWs As Worksheet
                Dim invTbl As ListObject
                Dim lookupRow As Long
                
                Set invWs = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
                Set invTbl = invWs.ListObjects("invSys")
                
                If ItemCode <> "" Then
                    lookupRow = FindRowByValue(invTbl, "ITEM_CODE", ItemCode)
                End If
                
                If lookupRow = 0 Then
                    lookupRow = FindRowByValue(invTbl, "ITEM", item)
                End If
                
                If lookupRow > 0 Then
                    rowNum = CStr(invTbl.DataBodyRange(lookupRow, invTbl.ListColumns("ROW").Index).value)
                End If
            End If
            
            ' Create a unique key - FIXED: For shipments from inventory, ensure items from different rows stay separate
            Dim uniqueKey As String
            If rowNum <> "" Then
                ' Use ROW for uniqueness (most specific)
                uniqueKey = "ROW_" & rowNum
            ElseIf ItemCode <> "" Then
                ' Use ITEM_CODE as fallback
                uniqueKey = "CODE_" & ItemCode
            Else
                ' If no ROW or ITEM_CODE, treat each entry as unique by including row position
                uniqueKey = "NAME_" & LCase(Trim(item)) & "|" & LCase(Trim(uom)) & "|" & i
            End If
            
            ' Tally items using the unique key
            If dict.Exists(uniqueKey) Then
                dict(uniqueKey) = dict(uniqueKey) + quantity
            Else
                dict.Add uniqueKey, quantity
                ' Store reference information
                dict.Add "info_" & uniqueKey, Array(item, ItemCode, rowNum, uom)
            End If
        End If
    Next i
    
    ' Configure form list box
    frm.lstBox.Clear
    frm.lstBox.ColumnCount = 5 ' ITEM, QTY, UOM, ITEM_CODE, ROW
    frm.lstBox.ColumnWidths = "150;50;50;0;0" ' Hide ITEM_CODE and ROW
    
    ' Add header row
    frm.lstBox.AddItem "ITEMS"
    frm.lstBox.List(0, 1) = "QTY"
    frm.lstBox.List(0, 2) = "UOM"
    
    ' Add data rows
    If dict.count > 0 Then
        For Each key In dict.Keys
            If Left$(key, 5) <> "info_" Then
                itemInfo = dict("info_" & key)
                
                frm.lstBox.AddItem itemInfo(0) ' Item name
                frm.lstBox.List(frm.lstBox.ListCount - 1, 1) = dict(key)   ' Quantity
                frm.lstBox.List(frm.lstBox.ListCount - 1, 2) = itemInfo(3) ' UOM
                frm.lstBox.List(frm.lstBox.ListCount - 1, 3) = itemInfo(1) ' ITEM_CODE
                frm.lstBox.List(frm.lstBox.ListCount - 1, 4) = itemInfo(2) ' ROW
            End If
        Next key
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Error in PopulateShipmentsForm: " & Err.Description, vbCritical
    Debug.Print "Error in PopulateShipmentsForm: " & Err.Description
    Resume Next
End Sub

Sub PopulateReceivedForm(frm As frmReceivedTally)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject, dataTbl As ListObject
    Dim dict As Object
    Dim priceDict As Object  ' New dictionary to track prices
    Dim i As Long, j As Long
    Dim key As Variant
    Dim itemInfo As Variant
    
    ' Get worksheet and table references
    Set ws = ThisWorkbook.Sheets("ReceivedTally")
    Set tbl = ws.ListObjects("ReceivedTally")
    Set dataTbl = ws.ListObjects("invSysData_Receiving")
    
    Set dict = CreateObject("Scripting.Dictionary")
    Set priceDict = CreateObject("Scripting.Dictionary")  ' For tracking prices
    dict.CompareMode = vbTextCompare
    
    ' Debug info
    Debug.Print "Processing ReceivedTally table with " & tbl.ListRows.Count & " rows"
    
    ' Process and tally items from the table
    For i = 1 To tbl.ListRows.Count
        On Error Resume Next
        
        ' Get basic values with error handling
        Dim item As String, quantity As Double, uom As String, price As Double
        item = CStr(tbl.DataBodyRange(i, tbl.ListColumns("ITEMS").Index).Value)
        
        ' Handle quantity conversion
        Dim rawQuantity As Variant
        rawQuantity = tbl.DataBodyRange(i, tbl.ListColumns("QUANTITY").Index).Value
        If IsNumeric(rawQuantity) Then
            quantity = CDbl(rawQuantity)
        Else
            quantity = 0
        End If
        
        ' Handle UOM and price
        On Error Resume Next
        uom = CStr(tbl.DataBodyRange(i, tbl.ListColumns("UOM").Index).Value)
        
        ' Get price from data table if available
        Dim priceColIndex As Long
        For j = 1 To tbl.ListColumns.Count
            If UCase(tbl.ListColumns(j).Name) = "PRICE" Then
                priceColIndex = j
                Exit For
            End If
        Next j
        
        If priceColIndex > 0 Then
            If IsNumeric(tbl.DataBodyRange(i, priceColIndex).Value) Then
                price = CDbl(tbl.DataBodyRange(i, priceColIndex).Value)
            End If
        End If
        
        On Error GoTo ErrorHandler
        
        ' Skip empty rows or rows with zero quantity
        If Trim(item) <> "" And quantity > 0 Then
            ' Get ROW and ITEM_CODE
            Dim rowNum As String, ItemCode As String
            rowNum = ""
            ItemCode = ""
            
            ' Check data tables for metadata
            If Not dataTbl Is Nothing Then
                ' Look for matching entry in data table
                For j = 1 To dataTbl.ListRows.Count
                    If CStr(dataTbl.DataBodyRange(j, dataTbl.ListColumns("ITEMS").Index).Value) = item Then
                        ' Get ROW and ITEM_CODE from data table
                        On Error Resume Next
                        For k = 1 To dataTbl.ListColumns.Count
                            If UCase(dataTbl.ListColumns(k).Name) = "ROW" Then
                                rowNum = CStr(dataTbl.DataBodyRange(j, k).Value)
                            ElseIf UCase(dataTbl.ListColumns(k).Name) = "ITEM_CODE" Then
                                ItemCode = CStr(dataTbl.DataBodyRange(j, k).Value)
                            End If
                        Next k
                        On Error GoTo ErrorHandler
                        Exit For
                    End If
                Next j
            End If
            
            ' If still no ROW, look it up in inventory
            If rowNum = "" Then
                Dim invWs As Worksheet
                Dim invTbl As ListObject
                Dim lookupRow As Long
                
                Set invWs = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
                Set invTbl = invWs.ListObjects("invSys")
                
                If ItemCode <> "" Then
                    lookupRow = FindRowByValue(invTbl, "ITEM_CODE", ItemCode)
                End If
                
                If lookupRow = 0 Then
                    lookupRow = FindRowByValue(invTbl, "ITEM", item)
                End If
                
                If lookupRow > 0 Then
                    On Error Resume Next
                    rowNum = CStr(invTbl.DataBodyRange(lookupRow, invTbl.ListColumns("ROW").Index).Value)
                    
                    ' If we have a row number but no UOM, get it from inventory
                    If Trim(uom) = "" Then
                        uom = CStr(invTbl.DataBodyRange(lookupRow, invTbl.ListColumns("UOM").Index).Value)
                    End If
                    On Error GoTo ErrorHandler
                End If
            End If
            
            ' Create a unique key for tallying
            Dim uniqueKey As String
            If rowNum <> "" Then
                uniqueKey = "ROW_" & rowNum
                Debug.Print "Using ROW key for " & item & ": " & uniqueKey
            ElseIf ItemCode <> "" Then
                uniqueKey = "CODE_" & ItemCode
                Debug.Print "Using CODE key for " & item & ": " & uniqueKey
            Else
                uniqueKey = "NAME_" & LCase(Trim(item)) & "|" & LCase(Trim(uom))
                Debug.Print "Using NAME key for " & item & ": " & uniqueKey
            End If
            
            ' Tally items
            If dict.Exists(uniqueKey) Then
                dict(uniqueKey) = dict(uniqueKey) + quantity
                priceDict(uniqueKey) = priceDict(uniqueKey) + price * quantity
            Else
                dict.Add uniqueKey, quantity
                dict.Add "info_" & uniqueKey, Array(item, ItemCode, rowNum, uom)
                priceDict.Add uniqueKey, price * quantity
            End If
        End If
    Next i
    
    ' Configure form list box
    frm.lstBox.Clear
    frm.lstBox.ColumnCount = 6 ' ITEM, QTY, UOM, PRICE, ITEM_CODE(hidden), ROW(hidden)
    frm.lstBox.ColumnWidths = "150;50;50;70;0;0" ' Show price, hide ITEM_CODE and ROW
    
    ' Add header row
    frm.lstBox.AddItem "ITEMS"
    frm.lstBox.List(0, 1) = "QTY"
    frm.lstBox.List(0, 2) = "UOM"
    frm.lstBox.List(0, 3) = "PRICE"
    
    ' Add data rows
    If dict.Count > 0 Then
        For Each key In dict.Keys
            If Left$(key, 5) <> "info_" Then
                itemInfo = dict("info_" & key)
                Dim totalQty As Double, totalPrice As Double, unitPrice As Double
                
                totalQty = dict(key)
                totalPrice = priceDict(key)
                
                ' Calculate unit price (price per unit)
                If totalQty > 0 Then
                    unitPrice = totalPrice / totalQty
                Else
                    unitPrice = 0
                End If
                
                frm.lstBox.AddItem itemInfo(0) ' Item name
                frm.lstBox.List(frm.lstBox.ListCount - 1, 1) = totalQty      ' Quantity
                frm.lstBox.List(frm.lstBox.ListCount - 1, 2) = itemInfo(3)   ' UOM
                frm.lstBox.List(frm.lstBox.ListCount - 1, 3) = Round(unitPrice, 2) ' Price (per unit)
                frm.lstBox.List(frm.lstBox.ListCount - 1, 4) = itemInfo(1)   ' ITEM_CODE
                frm.lstBox.List(frm.lstBox.ListCount - 1, 5) = itemInfo(2)   ' ROW
            End If
        Next key
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Error in PopulateReceivedForm: " & Err.Description & " (Line: " & Erl & ")", vbCritical
    Debug.Print "Error in PopulateReceivedForm: " & Err.Description & " at line " & Erl
    Resume Next
End Sub

Sub TallyReceived()
    On Error GoTo ErrorHandler
    
    ' Debug info
    Debug.Print "TallyReceived: Starting..."
    
    ' Verify the worksheet exists
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("ReceivedTally")
    On Error GoTo ErrorHandler
    If ws Is Nothing Then
        MsgBox "The worksheet 'ReceivedTally' does not exist!", vbExclamation
        Exit Sub
    End If
    
    ' Verify the table exists
    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = ws.ListObjects("ReceivedTally")
    On Error GoTo ErrorHandler
    If tbl Is Nothing Then
        MsgBox "The table 'ReceivedTally' does not exist on worksheet 'ReceivedTally'!", vbExclamation
        Exit Sub
    End If
    
    ' Create and show form with received items data
    Dim frm As New frmReceivedTally
    
    ' Configure the form
    With frm
        ' Make sure the listbox exists and is configured properly
        .lstBox.Clear
        .lstBox.ColumnCount = 5  ' ITEM, QTY, UOM, ITEM_CODE(hidden), ROW(hidden)
        .lstBox.ColumnWidths = "150;50;50;0;0" ' Hide ITEM_CODE and ROW columns
        .lstBox.AddItem "ITEMS"
        .lstBox.List(0, 1) = "QUANTITY"
        .lstBox.List(0, 2) = "UOM"
    End With
    
    ' Populate form directly using our PopulateReceivedForm function
    PopulateReceivedForm frm
    
    ' Show the form if there are items
    If frm.lstBox.ListCount > 1 Then ' More than just the header row
        frm.Show vbModal
    Else
        MsgBox "No received items to tally", vbInformation
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Error in TallyReceived: " & Err.Description & " (Error " & Err.Number & ")", vbCritical
    Debug.Print "Error in TallyReceived: " & Err.Description & " (Error " & Err.Number & ")"
End Sub

' This should be in your ribbon callback or worksheet button
Public Sub LaunchShipmentsTally()
    Application.ScreenUpdating = False
    TallyShipments
    Application.ScreenUpdating = True
End Sub

' This should be in your ribbon callback or worksheet button
Public Sub LaunchReceivedTally()
    Application.ScreenUpdating = False
    TallyReceived
    Application.ScreenUpdating = True
End Sub

' Helper function to find a row by column value
Private Function FindRowByValue(tbl As ListObject, colName As String, value As Variant) As Long
    Dim i As Long
    Dim colIndex As Integer
    
    FindRowByValue = 0 ' Default return value if not found
    
    On Error Resume Next
    colIndex = tbl.ListColumns(colName).Index
    On Error GoTo 0
    
    If colIndex = 0 Then Exit Function
    
    For i = 1 To tbl.ListRows.count
        ' Convert both values to strings for more reliable comparison
        If CStr(tbl.DataBodyRange(i, colIndex).value) = CStr(value) Then
            FindRowByValue = i
            Debug.Print "Found match in " & colName & " column: " & value & " at row " & i
            Exit Function
        End If
    Next i
    
    Debug.Print "No match found in " & colName & " column for value: " & CStr(value)
End Function
