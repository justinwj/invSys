Attribute VB_Name = "modTS_Log"
' ==============================================
' Module: modTS_Log (TS stands for Tally System)
' ==============================================
Option Explicit

' Shared function for generating unique order numbers
    Function GenerateOrderNumber() As String
        GenerateOrderNumber = "ORD" & Format(Now(), "YYMMDDHHMMSS")
    End Function

Sub LogShipments(shipmentSummary As Object)
    Dim key As Variant
    Dim newRow As ListRow
    For Each key In shipmentSummary.Keys
        Set newRow = ThisWorkbook.Sheets("ShipmentsLog").ListObjects("ShipmentsLog").ListRows.Add
        newRow.Range(1, 1).Value = GenerateOrderNumber()
        newRow.Range(1, 2).Value = key
        newRow.Range(1, 3).Value = shipmentSummary(key)
        newRow.Range(1, 4).Value = Now()
    Next key
End Sub

Sub LogReceived(receivedSummary As Object)
    Dim key As Variant
    Dim newRow As ListRow
    For Each key In receivedSummary.Keys
        Set newRow = ThisWorkbook.Sheets("ReceivedLog").ListObjects("ReceivedLog").ListRows.Add
        newRow.Range(1, 1).Value = GenerateOrderNumber()
        newRow.Range(1, 2).Value = key
        newRow.Range(1, 3).Value = receivedSummary(key)
        newRow.Range(1, 4).Value = Now()
    Next key
End Sub
