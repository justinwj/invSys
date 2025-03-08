Attribute VB_Name = "modTS_OrdersTally"
' ========================
' Module: modTS_OrdersTally
' ========================
Option Explicit
' This module is responsible for tallying orders and displaying them in a user form.
Sub TallyOrders()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dict As Object
    Dim i As Long
    Dim key As Variant
    Dim item As Variant, quantity As Double, uom As Variant
    Dim normItem As String, normUom As String
    Dim lb As MSForms.ListBox
    Dim keyParts As Variant
    
    Set ws = ThisWorkbook.Sheets("OrdersTally")
    Set tbl = ws.ListObjects("OrdersTally")
    Set dict = CreateObject("Scripting.Dictionary")
    ' Make dictionary case-insensitive
    dict.CompareMode = vbTextCompare
    Set lb = frmOrderTally.lstBox
    
    ' Tally the orders.
    For i = 1 To tbl.ListRows.count
        ' Get raw cell values.
        item = tbl.ListColumns("ITEMS").DataBodyRange(i, 1).Value
        quantity = tbl.ListColumns("QUANTITY").DataBodyRange(i, 1).Value
        uom = tbl.ListColumns("UOM").DataBodyRange(i, 1).Value
        
        ' More thorough normalization to handle edge cases
        ' First convert to string, then remove all excess spaces
        normItem = CStr(item)
        normItem = Application.WorksheetFunction.Trim(normItem)
        ' Replace multiple spaces with single space
        Do While InStr(normItem, "  ") > 0
            normItem = Replace(normItem, "  ", " ")
        Loop
        normItem = LCase(normItem)
        
        ' Same thorough normalization for UOM
        normUom = CStr(uom)
        normUom = Application.WorksheetFunction.Trim(normUom)
        Do While InStr(normUom, "  ") > 0
            normUom = Replace(normUom, "  ", " ")
        Loop
        normUom = LCase(normUom)
        
        ' Force default unit if missing.
        If normUom = "" Then normUom = "each"
        
        key = normItem & "|" & normUom
        
        If dict.Exists(key) Then
            dict(key) = dict(key) + quantity
        Else
            dict.Add key, quantity
        End If
    Next i
    
    ' Display the tally in the list box with three columns.
    lb.Clear
    lb.ColumnCount = 3
    lb.ColumnWidths = "150;50;30"   ' adjust widths as needed
    
    ' Manually add header row.
    lb.AddItem "ITEMS"
    lb.List(lb.ListCount - 1, 1) = "QUANTITY"
    lb.List(lb.ListCount - 1, 2) = "UOM"
    
    ' Add data rows.
    For Each key In dict.Keys
        keyParts = Split(key, "|")
        lb.AddItem
        lb.List(lb.ListCount - 1, 0) = keyParts(0)
        lb.List(lb.ListCount - 1, 1) = dict(key)
        lb.List(lb.ListCount - 1, 2) = keyParts(1)
    Next key
    
    frmOrderTally.Show
End Sub
