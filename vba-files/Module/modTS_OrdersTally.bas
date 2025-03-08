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
    Dim key As Variant              ' Changed from String to Variant
    Dim item As String
    Dim quantity As Double
    Dim uom As String
    Dim lb As MSForms.ListBox
    Dim keyParts As Variant
    
    Set ws = ThisWorkbook.Sheets("OrdersTally")
    Set tbl = ws.ListObjects("OrdersTally")
    Set dict = CreateObject("Scripting.Dictionary")
    Set lb = frmOrderTally.lstBox
    
    ' Tally the orders. Use a delimiter (|) to separate item and uom.
    For i = 1 To tbl.ListRows.Count
        item = tbl.ListColumns("ITEMS").DataBodyRange(i, 1).Value
        quantity = tbl.ListColumns("QUANTITY").DataBodyRange(i, 1).Value
        uom = tbl.ListColumns("UOM").DataBodyRange(i, 1).Value
        
        key = item & "|" & uom
        If dict.Exists(key) Then
            dict(key) = dict(key) + quantity
        Else
            dict.Add key, quantity
        End If
    Next i
    
    ' Display the tally in the list box with three columns: ITEMS, QUANTITY and UOM.
    lb.Clear
    lb.ColumnCount = 3
    lb.ColumnWidths = "150;50;30"   ' Adjust the widths as needed
    
    ' Manually add header row (headers now count as a normal row).
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
    
    ' Open the form.
    frmOrderTally.Show
End Sub
