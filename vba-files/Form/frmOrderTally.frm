VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOrderTally 
   Caption         =   "Order Tally"
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7005
   OleObjectBlob   =   "frmOrderTally.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOrderTally"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub SendOrders()
   Dim wsOrdersLog As Worksheet
   Dim wsOrderTally As Worksheet
   Dim wsInvSys As Worksheet
   Dim tblOrdersLog As ListObject
   Dim tblOrderTally As ListObject
   Dim tblInvSys As ListObject
   Dim i As Long
   
   Set wsOrdersLog = ThisWorkbook.Sheets("OrdersLog")
   Set wsOrderTally = ThisWorkbook.Sheets("OrdersTally")
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
   For i = 1 To frmOrderTally.lstBox.ListCount
       tblInvSys.ListRows.Add
       tblInvSys.ListRows(tblInvSys.ListRows.Count).Range(1, 1).Value = frmOrderTally.lstBox.List(i - 1, 0)
       tblInvSys.ListRows(tblInvSys.ListRows.Count).Range(1, 2).Value = frmOrderTally.lstBox.List(i - 1, 1)
       tblInvSys.ListRows(tblInvSys.ListRows.Count).Range(1, 3).Value = frmOrderTally.lstBox.List(i - 1, 2)
   Next i
End Sub