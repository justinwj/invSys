<<<<<<< HEAD
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
    Dim normItem As String, normUom As String
    Dim lb As MSForms.ListBox
    Dim keyParts As Variant
    
    ' Error checking for the form
    If formToShow Is Nothing Then
        MsgBox "Error: Form reference is null", vbExclamation
        Exit Sub
    End If
    
    ' Make sure the form has a lstBox control
    On Error Resume Next
    Set lb = formToShow.lstBox
    If Err.Number <> 0 Or lb Is Nothing Then
        MsgBox "Error: The form " & TypeName(formToShow) & " doesn't have a lstBox control", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    Set ws = ThisWorkbook.Sheets(sheetName)
    Set tbl = ws.ListObjects(tableName)
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    Set lb = formToShow.lstBox
    
    ' Tally the items
    For i = 1 To tbl.ListRows.count
        ' Get raw cell values
        item = tbl.ListColumns("ITEMS").DataBodyRange(i, 1).Value
        quantity = tbl.ListColumns("QUANTITY").DataBodyRange(i, 1).Value
        uom = tbl.ListColumns("UOM").DataBodyRange(i, 1).Value
        
        ' Skip rows where the item is empty or quantity is zero/empty
        If Trim(CStr(item)) <> "" And quantity > 0 Then
            ' Normalize item name and UOM
            normItem = NormalizeText(CStr(item))
            normUom = NormalizeText(CStr(uom))
            
            ' Force default unit if missing
            If normUom = "" Then normUom = "each"
            
            key = normItem & "|" & normUom
            
            If dict.Exists(key) Then
                dict(key) = dict(key) + quantity
            Else
                dict.Add key, quantity
            End If
        End If
    Next i
    
    ' Display the tally in the list box
    lb.Clear
    lb.ColumnCount = 3
    lb.ColumnWidths = "47;60;180"
    
    ' Add header row
    lb.AddItem "ITEMS"
    lb.List(lb.ListCount - 1, 1) = "QUANTITY"
    lb.List(lb.ListCount - 1, 2) = "UOM"
    
    ' Add data rows
    If dict.count > 0 Then
        For Each key In dict.Keys
            keyParts = Split(key, "|")
            lb.AddItem
            lb.List(lb.ListCount - 1, 0) = keyParts(0)
            lb.List(lb.ListCount - 1, 1) = dict(key)
            lb.List(lb.ListCount - 1, 2) = keyParts(1)
        Next key
        formToShow.Show
    Else
        MsgBox "No valid items found to tally.", vbInformation
    End If
End Sub

' Helper function to normalize text
Private Function NormalizeText(text As String) As String
    Dim result As String
    
    result = Application.WorksheetFunction.Trim(text)
    ' Replace multiple spaces with single space
    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop
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
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dict As Object
    Dim i As Long
    Dim j As Long
    Dim key As Variant
    Dim itemInfo As Variant  ' Moved this declaration up here
    
    Set ws = ThisWorkbook.Sheets("ShipmentsTally")
    Set tbl = ws.ListObjects("ShipmentsTally")
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    
    ' Process and tally items from the table
    For i = 1 To tbl.ListRows.Count
        ' Get basic values
        Dim item As String, quantity As Double, uom As String
        item = tbl.DataBodyRange(i, tbl.ListColumns("ITEMS").Index).Value
        quantity = tbl.DataBodyRange(i, tbl.ListColumns("QUANTITY").Index).Value
        uom = tbl.DataBodyRange(i, tbl.ListColumns("UOM").Index).Value
        
        ' Skip empty rows
        If Trim(item) <> "" And quantity > 0 Then
            ' Extract ROW# and ITEM_CODE from comments
            Dim rowNum As String, itemCode As String
            rowNum = ""
            itemCode = ""
            
            On Error Resume Next
            ' First check if ROW# and ITEM_CODE are in hidden columns
            For j = 1 To tbl.ListColumns.Count  ' Here's where j is used
                If UCase(tbl.ListColumns(j).Name) = "ROW#" Then
                    rowNum = tbl.DataBodyRange(i, j).Value
                ElseIf UCase(tbl.ListColumns(j).Name) = "ITEM_CODE" Then
                    itemCode = tbl.DataBodyRange(i, j).Value
                End If
            Next j
            
            ' If not found in columns, try comment
            If rowNum = "" Or itemCode = "" Then
                If Not tbl.DataBodyRange(i, tbl.ListColumns("ITEMS").Index).Comment Is Nothing Then
                    Dim commentText As String
                    commentText = tbl.DataBodyRange(i, tbl.ListColumns("ITEMS").Index).Comment.Text
                    
                    ' Extract ITEM_CODE
                    If InStr(commentText, "ITEM_CODE: ") > 0 Then
                        Dim startPos As Long, endPos As Long
                        startPos = InStr(commentText, "ITEM_CODE: ") + 11
                        endPos = InStr(startPos, commentText, vbCrLf)
                        If endPos > 0 Then
                            itemCode = Mid(commentText, startPos, endPos - startPos)
                        Else
                            itemCode = Mid(commentText, startPos)
                        End If
                    End If
                    
                    ' Extract ROW#
                    If InStr(commentText, "ROW#: ") > 0 Then
                        startPos = InStr(commentText, "ROW#: ") + 6
                        endPos = InStr(startPos, commentText, vbCrLf)
                        If endPos > 0 Then
                            rowNum = Mid(commentText, startPos, endPos - startPos)
                        Else
                            rowNum = Mid(commentText, startPos)
                        End If
                    End If
                End If
            End If
            On Error GoTo 0
            
            ' Create a unique key that includes ROW# if available
            Dim uniqueKey As String
            If rowNum <> "" Then
                ' Use ROW# for uniqueness (most specific)
                uniqueKey = "ROW_" & rowNum
            ElseIf itemCode <> "" Then
                ' Use ITEM_CODE as fallback
                uniqueKey = "CODE_" & itemCode
            Else
                ' Use item name and UOM as last resort
                uniqueKey = "NAME_" & LCase(Trim(item)) & "|" & LCase(Trim(uom))
            End If
            
            ' Tally items using the unique key
            If dict.Exists(uniqueKey) Then
                dict(uniqueKey) = dict(uniqueKey) + quantity
            Else
                dict.Add uniqueKey, quantity
                ' Store reference information
                dict.Add "info_" & uniqueKey, Array(item, itemCode, rowNum, uom)
            End If
        End If
    Next i
    
    ' Configure form list box
    frm.lstBox.Clear
    frm.lstBox.ColumnCount = 5 ' ITEM, QTY, UOM, ITEM_CODE, ROW#
    frm.lstBox.ColumnWidths = "150;50;50;0;0" ' Hide ITEM_CODE and ROW#
    
    ' Add header row
    frm.lstBox.AddItem "ITEMS"
    frm.lstBox.List(0, 1) = "QTY"
    frm.lstBox.List(0, 2) = "UOM"
    
    ' Add data rows
    If dict.Count > 0 Then
        For Each key In dict.Keys
            If Left$(key, 5) <> "info_" Then
                itemInfo = dict("info_" & key)
                
                frm.lstBox.AddItem
                frm.lstBox.List(frm.lstBox.ListCount - 1, 0) = itemInfo(0) ' Item name
                frm.lstBox.List(frm.lstBox.ListCount - 1, 1) = dict(key)   ' Quantity
                frm.lstBox.List(frm.lstBox.ListCount - 1, 2) = itemInfo(3) ' UOM
                frm.lstBox.List(frm.lstBox.ListCount - 1, 3) = itemInfo(1) ' ITEM_CODE
                frm.lstBox.List(frm.lstBox.ListCount - 1, 4) = itemInfo(2) ' ROW#
            End If
        Next key
    End If
End Sub

Sub TallyReceived()
    TallyItems "ReceivedTally", "ReceivedTally", frmReceivedTally
End Sub

' This should be in your ribbon callback or worksheet button
Public Sub LaunchShipmentsTally()
    Application.ScreenUpdating = False
    TallyShipments
    Application.ScreenUpdating = True
End Sub
=======
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
    Dim tbl As ListObject
    Dim dict As Object
    Dim i As Long
    Dim j As Long
    Dim key As Variant
    Dim itemInfo As Variant
    
    ' Get worksheet and table references
    Set ws = ThisWorkbook.Sheets("ReceivedTally")
    Set tbl = ws.ListObjects("ReceivedTally")
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    
    ' Debug info
    Debug.Print "Processing ReceivedTally table with " & tbl.ListRows.count & " rows"
    Debug.Print "Table headers: "
    For i = 1 To tbl.HeaderRowRange.Columns.count
        Debug.Print " - " & tbl.HeaderRowRange.Cells(1, i).value
    Next i
    
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
            Debug.Print "Warning: Non-numeric quantity at row " & i & ": " & rawQuantity
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
            
            ' If we don't have a ROW yet, use LOOKUP function to find in inventory
            If rowNum = "" Then
                ' Look up the item in the inventory table
                Dim invWs As Worksheet
                Dim invTbl As ListObject
                Dim lookupRow As Long
                
                Set invWs = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
                Set invTbl = invWs.ListObjects("invSys")
                
                ' Try to find by ITEM_CODE first if we have one
                If ItemCode <> "" Then
                    lookupRow = FindRowByValue(invTbl, "ITEM_CODE", ItemCode)
                End If
                
                ' If not found by ITEM_CODE, try by item name
                If lookupRow = 0 Then
                    lookupRow = FindRowByValue(invTbl, "ITEM", item)
                End If
                
                ' If found, get the ROW value
                If lookupRow > 0 Then
                    On Error Resume Next
                    rowNum = CStr(invTbl.DataBodyRange(lookupRow, invTbl.ListColumns("ROW").Index).value)
                    On Error GoTo ErrorHandler
                    Debug.Print "Found ROW " & rowNum & " for item " & item
                End If
            End If
            On Error GoTo ErrorHandler
            
            ' Create a unique key that correctly identifies inventory rows
            ' FIXED: Do NOT include table position (i) to allow items from same inventory row to group together
            Dim uniqueKey As String
            If rowNum <> "" Then
                ' Use ROW for uniqueness (most specific) - DO NOT include table row
                uniqueKey = "ROW_" & rowNum
                Debug.Print "Using ROW key for " & item & ": " & uniqueKey
            ElseIf ItemCode <> "" Then
                ' Use ITEM_CODE as fallback - DO NOT include table row
                uniqueKey = "CODE_" & ItemCode
                Debug.Print "Using CODE key for " & item & ": " & uniqueKey
            Else
                ' Use item name and UOM as last resort
                uniqueKey = "NAME_" & LCase(Trim(item)) & "|" & LCase(Trim(uom))
                Debug.Print "Using NAME key for " & item & ": " & uniqueKey
            End If
            
            ' Tally items using the unique key - items from same row get added together
            If dict.Exists(uniqueKey) Then
                dict(uniqueKey) = dict(uniqueKey) + quantity
            Else
                dict.Add uniqueKey, quantity
                ' Store reference information (item, itemCode, rowNum, uom)
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
    MsgBox "Error in PopulateReceivedForm: " & Err.Description & " (Error " & Err.Number & ")", vbCritical
    Debug.Print "Error in PopulateReceivedForm: " & Err.Description & " (Error " & Err.Number & ")"
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
>>>>>>> 0e2f3dc (Refactored a lot, tally system does not work as intended but big button works)
