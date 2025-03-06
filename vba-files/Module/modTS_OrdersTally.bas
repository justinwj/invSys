Attribute VB_Name = "modTS_OrdersTally"
' ========================
' Module: modTS_OrdersTally
' ========================

Sub TallyOrders()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dict As Object
    Dim i As Long
    Dim key As Variant
    Dim item As String
    Dim quantity As Double
    Dim uom As String
    Dim lb As MSForms.ListBox
    
    Set ws = ThisWorkbook.Sheets("Order Tally")
    Set tbl = ws.ListObjects("OrdersTally")
    Set dict = CreateObject("Scripting.Dictionary")
    Set lb = frmOrderTally.ListBox1
    
    ' Tally the orders
    For i = 1 To tbl.ListRows.Count
        item = tbl.ListColumns("ITEMS").DataBodyRange(i, 1).Value
        quantity = tbl.ListColumns("QUANTITY").DataBodyRange(i, 1).Value
        uom = tbl.ListColumns("UOM").DataBodyRange(i, 1).Value
        
        If dict.exists(item & uom) Then
            dict(item & uom) = dict(item & uom) + quantity
        Else
            dict.Add item & uom, quantity
        End If
    Next i
    
    ' Display the tally in the list box
    lb.Clear
    For Each key In dict.keys
        lb.AddItem Split(key, uom)(0)
        lb.List(lb.ListCount - 1, 1) = dict(key)
        lb.List(lb.ListCount - 1, 2) = uom
    Next key
End Sub

Sub SendOrders()
    Dim wsOrdersLog As Worksheet
    Dim wsOrderTally As Worksheet
    Dim wsInvSys As Worksheet
    Dim tblOrdersLog As ListObject
    Dim tblOrderTally As ListObject
    Dim tblInvSys As ListObject
    Dim i As Long
    
    Set wsOrdersLog = ThisWorkbook.Sheets("OrdersLog")
    Set wsOrderTally = ThisWorkbook.Sheets("Order Tally")
    Set wsInvSys = ThisWorkbook.Sheets("invSys")
    Set tblOrdersLog = wsOrdersLog.ListObjects("OrdersLog")
    Set tblOrderTally = wsOrderTally.ListObjects("OrdersTally")
    Set tblInvSys = wsInvSys.ListObjects("SHIPMENTS")
    
    ' Send orders to OrdersLog table
    For i = 1 To tblOrderTally.ListRows.Count
        tblOrdersLog.ListRows.Add
        tblOrdersLog.ListRows(tblOrdersLog.ListRows.Count).Range.Value = tblOrderTally.ListRows(i).Range.Value
    Next i
    
    ' Clear OrdersTally table
    tblOrderTally.DataBodyRange.ClearContents
    
    ' Send tally to SHIPMENTS header in invSys table
    For i = 1 To frmOrderTally.ListBox1.ListCount
        tblInvSys.ListRows.Add
        tblInvSys.ListRows(tblInvSys.ListRows.Count).Range(1, 1).Value = frmOrderTally.ListBox1.List(i - 1, 0)
        tblInvSys.ListRows(tblInvSys.ListRows.Count).Range(1, 2).Value = frmOrderTally.ListBox1.List(i - 1, 1)
        tblInvSys.ListRows(tblInvSys.ListRows.Count).Range(1, 3).Value = frmOrderTally.ListBox1.List(i - 1, 2)
    Next i
End Sub