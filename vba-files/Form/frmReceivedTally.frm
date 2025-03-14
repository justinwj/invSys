<<<<<<< HEAD
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReceivedTally 
   Caption         =   "Items Received Tally"
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7005
   OleObjectBlob   =   "frmReceivedTally.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmReceivedTally"
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
  
  Dim wsReceivedLog As Worksheet
  Dim wsReceivedTally As Worksheet
  Dim wsInvSys As Worksheet
  Dim tblReceivedLog As ListObject
  Dim tblReceivedTally As ListObject
  Dim tblInvSys As ListObject
  Dim i As Long
  Dim timestamp As String
  Dim orderClickID As String
  Dim hasValidItems As Boolean
  
  ' Get current timestamp for logging
  timestamp = Format(Now, "yyyy-mm-dd hh:mm:ss")
  ' Generate a unique ID for this order group
  orderClickID = "ReceivedTally-" & Format(Now, "yymmddhhmmss")
  
  ' Set worksheet and table references
  Set wsReceivedLog = ThisWorkbook.Sheets("ReceivedLog")
  Set wsReceivedTally = ThisWorkbook.Sheets("ReceivedTally")
  Set wsInvSys = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
  Set tblReceivedLog = wsReceivedLog.ListObjects("ReceivedLog")
  Set tblReceivedTally = wsReceivedTally.ListObjects("ReceivedTally")
  Set tblInvSys = wsInvSys.ListObjects("invSys")
  
  hasValidItems = False
  
  ' Step 1: Log the orders to ReceivedLog table
  For i = 1 To tblReceivedTally.ListRows.count
      ' Skip rows where the item is empty or quantity is zero/empty
      If Trim(CStr(tblReceivedTally.ListColumns("ITEMS").DataBodyRange(i, 1).Value)) <> "" And _
         CDbl(Val(tblReceivedTally.ListColumns("QUANTITY").DataBodyRange(i, 1).Value)) > 0 Then
          
          hasValidItems = True
          
          With tblReceivedLog.ListRows.Add
              ' Leave ORDER_NUMBER empty - will be filled from QuickBooks
              .Range(1).Value = ""  ' ORDER_NUMBER should be empty or from QuickBooks
              
              ' Copy remaining values from ReceivedTally to ReceivedLog
              .Range(2).Value = tblReceivedTally.ListRows(i).Range(2).Value  ' ITEMS
              .Range(3).Value = tblReceivedTally.ListRows(i).Range(3).Value  ' QUANTITY
              .Range(4).Value = tblReceivedTally.ListRows(i).Range(4).Value  ' UOM
              .Range(5).Value = timestamp                                    ' TIMESTAMP
              .Range(6).Value = orderClickID                                 ' ON_CLICK_ID (consistent for batch)
          End With
      End If
  Next i
  
  ' Step 2: Send the inventory tally to invSys table - update RECEIVED column
  ' Skip header row (index 0) in the listbox
  For i = 1 To Me.lstBox.ListCount - 1
      Dim itemName As String
      Dim quantity As Double
      Dim uom As String
      Dim foundCell As Range
      Dim foundRow As Long
      Dim itemCode As String
      
      ' Get values from the listbox
      itemName = Me.lstBox.List(i, 0)  ' ITEMS column
      quantity = Val(Me.lstBox.List(i, 1))  ' QUANTITY column
      uom = Me.lstBox.List(i, 2)  ' UOM column
      
      ' Find ITEM_CODE from cell comment or hidden column
      itemCode = ""
      On Error Resume Next
      itemCode = tblReceivedTally.ListColumns("ITEM_CODE").DataBodyRange(i).Value
      On Error GoTo 0
      
      ' Skip processing if item is empty or quantity is zero
      If Trim(itemName) <> "" And quantity > 0 Then
          ' Find the item in the invSys table - prefer ITEM_CODE if available
          If itemCode <> "" Then
              Set foundCell = tblInvSys.ListColumns("ITEM_CODE").DataBodyRange.Find(itemCode, LookAt:=xlWhole)
          Else
              Set foundCell = tblInvSys.ListColumns("ITEM").DataBodyRange.Find(itemName, LookAt:=xlWhole)
          End If
          
          If Not foundCell Is Nothing Then
              ' Found the item, update the RECEIVED value
              foundRow = foundCell.row - tblInvSys.HeaderRowRange.row
              
              ' Update the RECEIVED column for this item
              ' Add to existing value if there is one
              Dim currentVal As Variant
              currentVal = tblInvSys.ListColumns("RECEIVED").DataBodyRange(foundRow).Value
              
              If IsNumeric(currentVal) Then
                  tblInvSys.ListColumns("RECEIVED").DataBodyRange(foundRow).Value = currentVal + quantity
              Else
                  tblInvSys.ListColumns("RECEIVED").DataBodyRange(foundRow).Value = quantity
              End If
              
              ' Update the LAST EDITED column
              tblInvSys.ListColumns("LAST EDITED").DataBodyRange(foundRow).Value = Now
          Else
              MsgBox "Warning: Item '" & itemName & "' was not found in the inventory system.", vbExclamation
          End If
      End If
  Next i
  
  ' Step 3: Clear the OrdersTally table (properly delete rows)
  If Not tblReceivedTally.DataBodyRange Is Nothing Then
      ' Delete all rows from the table
      ' First check if there's more than one row to avoid errors
      If tblReceivedTally.ListRows.count > 0 Then
          ' Delete each row - must delete in reverse order to avoid index issues
          For i = tblReceivedTally.ListRows.count To 1 Step -1
              tblReceivedTally.ListRows(i).Delete
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

=======
VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReceivedTally 
   Caption         =   "Items Received Tally"
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   7010
   OleObjectBlob   =   "frmReceivedTally.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmReceivedTally"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


' Handle btnSend click event
Private Sub btnSend_Click()
    On Error GoTo ErrorHandler
    
    ' Process the data first
    SendOrderData
    
    ' Then unload the form
    Unload Me
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in btnSend_Click: " & Err.Description
    MsgBox "Error: " & Err.Description, vbCritical
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
    
    Dim i As Long, shipmentsSummary As Object
    Set shipmentsSummary = CreateObject("Scripting.Dictionary")
    
    ' Skip header row (row 0)
    For i = 1 To Me.lstBox.ListCount - 1
        Dim item As String, quantity As Double, uom As String
        Dim ItemCode As String, rowNum As String
        
        item = Me.lstBox.List(i, 0)
        quantity = CDbl(Me.lstBox.List(i, 1))
        uom = Me.lstBox.List(i, 2)
        
        ' Get the hidden columns with ITEM_CODE and ROW
        ItemCode = Me.lstBox.List(i, 3)
        rowNum = Me.lstBox.List(i, 4)
        
        ' Create a unique key with ROW or ITEM_CODE
        Dim uniqueKey As String
        If rowNum <> "" Then
            uniqueKey = "ROW_" & rowNum
        ElseIf ItemCode <> "" Then
            uniqueKey = "CODE_" & ItemCode
        Else
            uniqueKey = "NAME_" & item & "|" & uom
        End If
        
        ' Store in dictionary with all needed information
        shipmentsSummary(uniqueKey) = Array(item, quantity, uom, ItemCode, rowNum)
    Next i
    
    ' Log the received items - pass the full dictionary with ROW and ITEM_CODE data
    modTS_Log.LogReceived shipmentsSummary
    
    ' Update quantities in inventory system
    UpdateInventory shipmentsSummary, "RECEIVED"
    
    ' Close form after processing
    Unload Me
    Exit Sub
    
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub

' Function to update inventory based on ROW or ITEM_CODE
Private Sub UpdateInventory(itemsDict As Object, ColumnName As String)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim key As Variant
    Dim foundRow As Long
    Dim currentQty As Double, newQty As Double
    
    Set ws = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
    Set tbl = ws.ListObjects("invSys")
    
    ' Get column index for the target column (e.g., "RECEIVED", "SHIPMENTS")
    Dim targetColIndex As Integer
    targetColIndex = tbl.ListColumns(ColumnName).Index
    
    ws.Unprotect
    Application.EnableEvents = False
    
    For Each key In itemsDict.Keys
        Dim itemData As Variant
        itemData = itemsDict(key)
        
        ' Extract info from the array
        Dim item As String, quantity As Double
        Dim ItemCode As String, rowNum As String
        item = itemData(0)
        quantity = itemData(1)
        ItemCode = itemData(3) ' itemCode at index 3
        rowNum = itemData(4)   ' rowNum at index 4
        
        foundRow = 0
        
        ' Try to find by ROW number first (most specific)
        If rowNum <> "" Then
            On Error Resume Next
            foundRow = FindRowByValue(tbl, "ROW", rowNum)
            On Error GoTo ErrorHandler
        End If
        
        ' If ROW didn't work, try ITEM_CODE
        If foundRow = 0 And ItemCode <> "" Then
            On Error Resume Next
            foundRow = FindRowByValue(tbl, "ITEM_CODE", ItemCode)
            On Error GoTo ErrorHandler
        End If
        
        ' As last resort, try finding by item name
        If foundRow = 0 Then
            On Error Resume Next
            foundRow = FindRowByValue(tbl, "ITEM", item)
            On Error GoTo ErrorHandler
        End If
        
        ' If we found the row, update it
        If foundRow > 0 Then
            ' Get current quantity
            currentQty = 0
            On Error Resume Next
            currentQty = tbl.DataBodyRange(foundRow, targetColIndex).value
            If IsEmpty(currentQty) Then currentQty = 0
            On Error GoTo ErrorHandler
            
            ' Update with new quantity
            newQty = currentQty + quantity
            tbl.DataBodyRange(foundRow, targetColIndex).value = newQty
            
            ' Log this change
            LogInventoryChange "UPDATE", ItemCode, item, quantity, newQty
        Else
            ' Log that we couldn't find the item
            LogInventoryChange "ERROR", ItemCode, item, quantity, 0
        End If
    Next key
    
    Application.EnableEvents = True
    ws.Protect
    Exit Sub
    
ErrorHandler:
    Application.EnableEvents = True
    ws.Protect
    MsgBox "Error updating inventory: " & Err.Description, vbCritical
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
        If tbl.DataBodyRange(i, colIndex).value = value Then
            FindRowByValue = i
            Exit Function
        End If
    Next i
End Function

' Helper function to log inventory changes
Private Sub LogInventoryChange(Action As String, ItemCode As String, itemName As String, qtyChange As Double, newQty As Double)
    ' This would call your inventory logging system
    On Error Resume Next
    ' You might want to use the modTS_Log module for this
End Sub
>>>>>>> 0e2f3dc (Refactored a lot, tally system does not work as intended but big button works)
