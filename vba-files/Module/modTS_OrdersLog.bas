Attribute VB_Name = "modTS_OrdersLog"
' ========================
' Module: modTS_OrdersLog
' ========================
Option Explicit

Sub LogOrders(orderSummary As Object)
    Dim key As Variant
    Dim newRow As ListRow
    For Each key In orderSummary.Keys
        Set newRow = ThisWorkbook.Sheets("OrdersLog").ListObjects("OrdersLog").ListRows.Add
        newRow.Range(1, 1).Value = GenerateOrderNumber()
        newRow.Range(1, 2).Value = key
        newRow.Range(1, 3).Value = orderSummary(key)
        newRow.Range(1, 4).Value = Now()
    Next key
End Sub

Function GenerateOrderNumber() As String
    GenerateOrderNumber = "ORD" & Format(Now(), "YYMMDDHHMMSS")
End Function
