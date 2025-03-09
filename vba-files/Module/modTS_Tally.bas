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
