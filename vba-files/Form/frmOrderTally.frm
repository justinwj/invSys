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


' Handle btnSend click event
Private Sub btnSend_Click()
    SendOrderData
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    ' The lstBox should already be populated by TallyOrders()
    ' Center the form on screen
    Me.StartUpPosition = 0 'Manual
    Me.Left = Application.Left + (Application.Width - Me.Width) / 2
    Me.Top = Application.Top + (Application.Height - Me.Height) / 2
End Sub

Sub SendOrderData()
   On Error GoTo ErrorHandler
   
   Dim wsOrdersLog As Worksheet
   Dim wsOrderTally As Worksheet
   Dim wsInvSys As Worksheet
   Dim tblOrdersLog As ListObject
   Dim tblOrderTally As ListObject
   Dim tblInvSys As ListObject
   Dim i As Long
   Dim timestamp As String
   Dim orderClickID As String
   Dim hasValidItems As Boolean
   
   ' Get current timestamp for logging
   timestamp = Format(Now, "yyyy-mm-dd hh:mm:ss")
   ' Generate a unique ID for this order group
   orderClickID = "OrderTally-" & Format(Now, "yymmddhhmmss")
   
   ' Set worksheet and table references
   Set wsOrdersLog = ThisWorkbook.Sheets("OrdersLog")
   Set wsOrderTally = ThisWorkbook.Sheets("OrdersTally")
   Set wsInvSys = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
   Set tblOrdersLog = wsOrdersLog.ListObjects("OrdersLog")
   Set tblOrderTally = wsOrderTally.ListObjects("OrdersTally")
   Set tblInvSys = wsInvSys.ListObjects("invSys")
   
   hasValidItems = False
   
   ' Step 1: Log the orders to OrdersLog table
   For i = 1 To tblOrderTally.ListRows.count
       ' Skip rows where the item is empty or quantity is zero/empty
       If Trim(CStr(tblOrderTally.ListColumns("ITEMS").DataBodyRange(i, 1).Value)) <> "" And _
          CDbl(Val(tblOrderTally.ListColumns("QUANTITY").DataBodyRange(i, 1).Value)) > 0 Then
           
           hasValidItems = True
           
           With tblOrdersLog.ListRows.Add
               ' Copy values from OrdersTally to OrdersLog
               .Range(1).Value = tblOrderTally.ListRows(i).Range(1).Value  ' ORDER_NUMBER
               .Range(2).Value = tblOrderTally.ListRows(i).Range(2).Value  ' ITEMS
               .Range(3).Value = tblOrderTally.ListRows(i).Range(3).Value  ' QUANTITY
               .Range(4).Value = tblOrderTally.ListRows(i).Range(4).Value  ' UOM
               .Range(5).Value = timestamp                                ' TIMESTAMP
               .Range(6).Value = orderClickID                             ' ON_CLICK_ID (consistent for batch)
           End With
       End If
   Next i
   
   ' Step 2: Send the inventory tally to invSys table - update SHIPMENTS column
   ' Skip header row (index 0) in the listbox
   For i = 1 To Me.lstBox.ListCount - 1
       Dim itemName As String
       Dim quantity As Double
       Dim uom As String
       Dim foundCell As Range
       Dim foundRow As Long
       
       ' Get values from the listbox
       itemName = Me.lstBox.List(i, 0)  ' ITEMS column
       quantity = Val(Me.lstBox.List(i, 1))  ' QUANTITY column
       uom = Me.lstBox.List(i, 2)  ' UOM column
       
       ' Skip processing if item is empty or quantity is zero
       If Trim(itemName) <> "" And quantity > 0 Then
           ' Find the item in the invSys table
           Set foundCell = tblInvSys.ListColumns("ITEM").DataBodyRange.Find(itemName, LookAt:=xlWhole)
           
           If Not foundCell Is Nothing Then
               ' Found the item, update the SHIPMENTS value
               foundRow = foundCell.row - tblInvSys.HeaderRowRange.row
               
               ' Update the SHIPMENTS column for this item
               ' Add to existing value if there is one
               Dim currentVal As Variant
               currentVal = tblInvSys.ListColumns("SHIPMENTS").DataBodyRange(foundRow).Value
               
               If IsNumeric(currentVal) Then
                   tblInvSys.ListColumns("SHIPMENTS").DataBodyRange(foundRow).Value = currentVal + quantity
               Else
                   tblInvSys.ListColumns("SHIPMENTS").DataBodyRange(foundRow).Value = quantity
               End If
               
               ' Update the LAST EDITED column
               tblInvSys.ListColumns("LAST EDITED").DataBodyRange(foundRow).Value = Now
           Else
               MsgBox "Warning: Item '" & itemName & "' was not found in the inventory system.", vbExclamation
           End If
       End If
   Next i
   
   ' Step 3: Clear the OrdersTally table (properly delete rows)
   If Not tblOrderTally.DataBodyRange Is Nothing Then
       ' Delete all rows from the table
       ' First check if there's more than one row to avoid errors
       If tblOrderTally.ListRows.count > 0 Then
           ' Delete each row - must delete in reverse order to avoid index issues
           For i = tblOrderTally.ListRows.count To 1 Step -1
               tblOrderTally.ListRows(i).Delete
           Next i
       End If
   End If
   
   If hasValidItems Then
       MsgBox "Order data has been processed successfully!", vbInformation
   Else
       MsgBox "No valid order data was found to process.", vbInformation
   End If
   
   Exit Sub
   
ErrorHandler:
   MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & _
          "At line: " & Erl(), vbCritical
End Sub

