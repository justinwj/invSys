Attribute VB_Name = "modInvMan"
Public Sub AddGoodsReceived_Click()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dataArr As Variant
    Dim receivedCol As Long, totalInvCol As Long, itemCodeCol As Long, itemNameCol As Long
    Dim i As Long, rowCount As Long
    Dim LogEntries As Collection
    Dim insertedCount As Long

    On Error GoTo ErrorHandler

    ' Start Transaction - groups all changes into one Undo step
    Call modUR_Transaction.BeginTransaction

    ' Set reference to invSys table
    Set ws = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
    Set tbl = ws.ListObjects("invSys")

    If tbl Is Nothing Or tbl.ListRows.count = 0 Then
        MsgBox "No data in invSys table.", vbExclamation, "Error"
        GoTo Cleanup
    End If

    ' Get column indexes dynamically
    itemCodeCol = tbl.ListColumns("ITEM_CODE").Index
    itemNameCol = tbl.ListColumns("ITEM").Index
    receivedCol = tbl.ListColumns("RECEIVED").Index
    totalInvCol = tbl.ListColumns("TOTAL INV").Index

    dataArr = tbl.DataBodyRange.Value
    rowCount = UBound(dataArr, 1)

    Set LogEntries = New Collection

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    For i = 1 To rowCount
        If IsNumeric(dataArr(i, receivedCol)) And dataArr(i, receivedCol) > 0 Then
            Dim OldTotalInv As Variant
            OldTotalInv = dataArr(i, totalInvCol)
            dataArr(i, totalInvCol) = dataArr(i, totalInvCol) + dataArr(i, receivedCol)
            
            Call modUR_Transaction.TrackTransactionChange("CellUpdate", _
                    dataArr(i, itemCodeCol), "TOTAL INV", OldTotalInv, dataArr(i, totalInvCol))
            
            ' Build log entry array: USER, ACTION, ITEM_CODE, ITEM_NAME, QUANTITY_CHANGE, NEW_QUANTITY
            LogEntries.Add Array(Environ("USERNAME"), "Added Goods Received", _
                dataArr(i, itemCodeCol), dataArr(i, itemNameCol), dataArr(i, receivedCol), dataArr(i, totalInvCol))
            
            dataArr(i, receivedCol) = 0
        End If
    Next i

    tbl.DataBodyRange.Value = dataArr

    ' Log transaction in InventoryLog if there are any entries; capture count of inserted rows.
    If LogEntries.count > 0 Then
        insertedCount = modInvLogs.LogMultipleInventoryChanges(LogEntries)
        Call modUR_Transaction.SetCurrentTransactionLogCount(insertedCount)
    End If

    Call modUR_Transaction.CommitTransaction

Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Call DisplayMessage("Goods received successfully.")
    Exit Sub

ErrorHandler:
    ' Rollback transaction if an error occurs
    If modUR_Transaction.IsInTransaction() Then
        Call modUR_Transaction.RollbackTransaction
    End If
    Call LogAndHandleError("AddGoodsReceived_Click")
    Resume Cleanup
End Sub
Public Sub DeductUsed_Click()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dataArr As Variant
    Dim usedCol As Long, totalInvCol As Long, itemCodeCol As Long, itemNameCol As Long
    Dim i As Long, rowCount As Long
    Dim LogEntries As Collection
    Dim insertedCount As Long

    On Error GoTo ErrorHandler

    ' Start Transaction - Groups all changes in one Undo step
    Call modUR_Transaction.BeginTransaction

    ' Set reference to invSys table
    Set ws = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
    Set tbl = ws.ListObjects("invSys")

    ' Check if table exists and has data
    If tbl Is Nothing Or tbl.ListRows.count = 0 Then
        MsgBox "No data in invSys table.", vbExclamation, "Error"
        GoTo Cleanup
    End If

    ' Get column indexes dynamically
    itemCodeCol = tbl.ListColumns("ITEM_CODE").Index
    itemNameCol = tbl.ListColumns("ITEM").Index
    usedCol = tbl.ListColumns("USED").Index
    totalInvCol = tbl.ListColumns("TOTAL INV").Index

    ' Store data in an array for fast processing
    dataArr = tbl.DataBodyRange.Value
    rowCount = UBound(dataArr, 1)

    ' Initialize log collection
    Set LogEntries = New Collection

    ' Disable screen updating & calculations for better performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Process each row in the array
    For i = 1 To rowCount
        ' Ensure USED has a valid numeric value and is greater than 0
        If IsNumeric(dataArr(i, usedCol)) And dataArr(i, usedCol) > 0 Then
            ' Store previous value for undo tracking
            Dim OldTotalInv As Variant
            OldTotalInv = dataArr(i, totalInvCol)

            ' Deduct used items from total inventory
            dataArr(i, totalInvCol) = dataArr(i, totalInvCol) - dataArr(i, usedCol)

            ' Track change as part of the transaction
            Call modUR_Transaction.TrackTransactionChange("CellUpdate", _
                dataArr(i, itemCodeCol), "TOTAL INV", OldTotalInv, dataArr(i, totalInvCol))

            ' Store log entry in collection
            LogEntries.Add Array(Environ("USERNAME"), "Deducted Used Items", _
                dataArr(i, itemCodeCol), dataArr(i, itemNameCol), -dataArr(i, usedCol), dataArr(i, totalInvCol))

            ' Reset USED column to zero
            dataArr(i, usedCol) = 0
        End If
    Next i

    ' Write updated inventory values back to the sheet in one operation
    tbl.DataBodyRange.Value = dataArr

    ' Log transaction in InventoryLog if there are any entries; capture count of inserted rows
    If LogEntries.count > 0 Then
        insertedCount = modInvLogs.LogMultipleInventoryChanges(LogEntries)
        Call modUR_Transaction.SetCurrentTransactionLogCount(insertedCount)
    End If

    ' Commit the transaction - Makes Undo treat all changes as one step
    Call modUR_Transaction.CommitTransaction

Cleanup:
    ' Restore screen updating & calculations
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Call DisplayMessage("Used items deducted successfully.")
    Exit Sub

ErrorHandler:
    ' Rollback transaction if an error occurs
    If modUR_Transaction.IsInTransaction() Then
        Call modUR_Transaction.RollbackTransaction
    End If
    Call LogAndHandleError("DeductUsed_Click")
    Resume Cleanup
End Sub
Public Sub AddMadeItems_Click()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dataArr As Variant
    Dim madeCol As Long, totalInvCol As Long, itemCodeCol As Long, itemNameCol As Long
    Dim i As Long, rowCount As Long
    Dim LogEntries As Collection
    Dim insertedCount As Long

    On Error GoTo ErrorHandler

    ' Start Transaction - Groups all changes in one Undo step
    Call modUR_Transaction.BeginTransaction

    ' Set reference to invSys table
    Set ws = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
    Set tbl = ws.ListObjects("invSys")

    ' Check if table exists and has data
    If tbl Is Nothing Or tbl.ListRows.count = 0 Then
        MsgBox "No data in invSys table.", vbExclamation, "Error"
        GoTo Cleanup
    End If

    ' Get column indexes dynamically
    itemCodeCol = tbl.ListColumns("ITEM_CODE").Index
    itemNameCol = tbl.ListColumns("ITEM").Index
    madeCol = tbl.ListColumns("MADE").Index
    totalInvCol = tbl.ListColumns("TOTAL INV").Index

    ' Store data in an array for fast processing
    dataArr = tbl.DataBodyRange.Value
    rowCount = UBound(dataArr, 1)

    ' Initialize log collection
    Set LogEntries = New Collection

    ' Disable screen updating & calculations for better performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Process each row in the array
    For i = 1 To rowCount
        ' Ensure MADE has a valid numeric value and is greater than 0
        If IsNumeric(dataArr(i, madeCol)) And dataArr(i, madeCol) > 0 Then
            ' Store previous value for undo tracking
            Dim OldTotalInv As Variant
            OldTotalInv = dataArr(i, totalInvCol)

            ' Add made quantity to total inventory
            dataArr(i, totalInvCol) = dataArr(i, totalInvCol) + dataArr(i, madeCol)

            ' Track change as part of the transaction
            Call modUR_Transaction.TrackTransactionChange("CellUpdate", _
                dataArr(i, itemCodeCol), "TOTAL INV", OldTotalInv, dataArr(i, totalInvCol))

            ' Store log entry in collection
            LogEntries.Add Array(Environ("USERNAME"), "Made Items Added", _
                dataArr(i, itemCodeCol), dataArr(i, itemNameCol), dataArr(i, madeCol), dataArr(i, totalInvCol))

            ' Reset MADE column to zero
            dataArr(i, madeCol) = 0
        End If
    Next i

    ' Write updated inventory values back to the sheet in one operation
    tbl.DataBodyRange.Value = dataArr

    ' Log transaction in InventoryLog if there are any entries; capture count of inserted rows
    If LogEntries.count > 0 Then
        insertedCount = modInvLogs.LogMultipleInventoryChanges(LogEntries)
        Call modUR_Transaction.SetCurrentTransactionLogCount(insertedCount)
    End If

    ' Commit the transaction - Makes Undo treat all changes as one step
    Call modUR_Transaction.CommitTransaction

Cleanup:
    ' Restore screen updating & calculations
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Call DisplayMessage("Made items added successfully.")
    Exit Sub

ErrorHandler:
    ' Rollback transaction if an error occurs
    If modUR_Transaction.IsInTransaction() Then
        Call modUR_Transaction.RollbackTransaction
    End If
    Call LogAndHandleError("AddMadeItems_Click")
    Resume Cleanup
End Sub

Public Sub DeductShipments_Click()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dataArr As Variant
    Dim shipmentsCol As Long, totalInvCol As Long, itemCodeCol As Long, itemNameCol As Long
    Dim i As Long, rowCount As Long
    Dim LogEntries As Collection
    Dim insertedCount As Long

    On Error GoTo ErrorHandler

    ' Start Transaction - Groups all changes in one Undo step
    Call modUR_Transaction.BeginTransaction

    ' Set reference to invSys table
    Set ws = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
    Set tbl = ws.ListObjects("invSys")

    ' Check if table exists and has data
    If tbl Is Nothing Or tbl.ListRows.count = 0 Then
        MsgBox "No data in invSys table.", vbExclamation, "Error"
        GoTo Cleanup
    End If

    ' Get column indexes dynamically
    itemCodeCol = tbl.ListColumns("ITEM_CODE").Index
    itemNameCol = tbl.ListColumns("ITEM").Index
    shipmentsCol = tbl.ListColumns("SHIPMENTS").Index
    totalInvCol = tbl.ListColumns("TOTAL INV").Index

    ' Store data in an array for fast processing
    dataArr = tbl.DataBodyRange.Value
    rowCount = UBound(dataArr, 1)

    ' Initialize log collection
    Set LogEntries = New Collection

    ' Disable screen updating & calculations for better performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Process each row in the array
    For i = 1 To rowCount
        ' Ensure SHIPMENTS has a valid numeric value and is greater than 0
        If IsNumeric(dataArr(i, shipmentsCol)) And dataArr(i, shipmentsCol) > 0 Then
            ' Store previous value for undo tracking
            Dim OldTotalInv As Variant
            OldTotalInv = dataArr(i, totalInvCol)

            ' Deduct shipments from total inventory
            dataArr(i, totalInvCol) = dataArr(i, totalInvCol) - dataArr(i, shipmentsCol)

            ' Track change as part of the transaction
            Call modUR_Transaction.TrackTransactionChange("CellUpdate", _
                dataArr(i, itemCodeCol), "TOTAL INV", OldTotalInv, dataArr(i, totalInvCol))

            ' Store log entry in collection
            LogEntries.Add Array(Environ("USERNAME"), "Shipments Deducted", _
                dataArr(i, itemCodeCol), dataArr(i, itemNameCol), -dataArr(i, shipmentsCol), dataArr(i, totalInvCol))

            ' Reset SHIPMENTS column to zero
            dataArr(i, shipmentsCol) = 0
        End If
    Next i

    ' Write updated inventory values back to the sheet in one operation
    tbl.DataBodyRange.Value = dataArr

    ' Log transaction in InventoryLog if there are any entries; capture count of inserted rows
    If LogEntries.count > 0 Then
        insertedCount = modInvLogs.LogMultipleInventoryChanges(LogEntries)
        Call modUR_Transaction.SetCurrentTransactionLogCount(insertedCount)
    End If

    ' Commit the transaction - Makes Undo treat all changes as one step
    Call modUR_Transaction.CommitTransaction

Cleanup:
    ' Restore screen updating & calculations
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Call DisplayMessage("Shipments deducted successfully.")
    Exit Sub

ErrorHandler:
    ' Rollback transaction if an error occurs
    If modUR_Transaction.IsInTransaction() Then
        Call modUR_Transaction.RollbackTransaction
    End If
    Call LogAndHandleError("DeductShipments_Click")
    Resume Cleanup
End Sub
Public Sub Adjustments_Click()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dataArr As Variant
    Dim adjustmentsCol As Long, totalInvCol As Long, itemCodeCol As Long, itemNameCol As Long
    Dim i As Long, rowCount As Long
    Dim LogEntries As Collection
    Dim insertedCount As Long

    On Error GoTo ErrorHandler

    ' Start Transaction - Groups all changes in one Undo step
    Call modUR_Transaction.BeginTransaction

    ' Set reference to invSys table
    Set ws = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
    Set tbl = ws.ListObjects("invSys")

    ' Check if table exists and has data
    If tbl Is Nothing Or tbl.ListRows.count = 0 Then
        MsgBox "No data in invSys table.", vbExclamation, "Error"
        GoTo Cleanup
    End If

    ' Get column indexes dynamically
    itemCodeCol = tbl.ListColumns("ITEM_CODE").Index
    itemNameCol = tbl.ListColumns("ITEM").Index
    adjustmentsCol = tbl.ListColumns("ADJUSTMENTS").Index
    totalInvCol = tbl.ListColumns("TOTAL INV").Index

    ' Store data in an array for fast processing
    dataArr = tbl.DataBodyRange.Value
    rowCount = UBound(dataArr, 1)

    ' Initialize log collection
    Set LogEntries = New Collection

    ' Disable screen updating & calculations for better performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Process each row in the array
    For i = 1 To rowCount
        ' Ensure ADJUSTMENTS has a valid numeric value
        If IsNumeric(dataArr(i, adjustmentsCol)) And dataArr(i, adjustmentsCol) <> 0 Then
            ' Store previous value for undo tracking
            Dim OldTotalInv As Variant
            OldTotalInv = dataArr(i, totalInvCol)

            ' Apply adjustment (positive adds, negative subtracts)
            dataArr(i, totalInvCol) = dataArr(i, totalInvCol) + dataArr(i, adjustmentsCol)

            ' Track change as part of the transaction
            Call modUR_Transaction.TrackTransactionChange("CellUpdate", _
                dataArr(i, itemCodeCol), "TOTAL INV", OldTotalInv, dataArr(i, totalInvCol))

            ' Store log entry in collection
            LogEntries.Add Array(Environ("USERNAME"), "Inventory Adjustment", _
                dataArr(i, itemCodeCol), dataArr(i, itemNameCol), dataArr(i, adjustmentsCol), dataArr(i, totalInvCol))

            ' Reset ADJUSTMENTS column to zero
            dataArr(i, adjustmentsCol) = 0
        End If
    Next i

    ' Write updated inventory values back to the sheet in one operation
    tbl.DataBodyRange.Value = dataArr

    ' Log transaction in InventoryLog if there are any entries; capture count of inserted rows
    If LogEntries.count > 0 Then
        insertedCount = modInvLogs.LogMultipleInventoryChanges(LogEntries)
        Call modUR_Transaction.SetCurrentTransactionLogCount(insertedCount)
    End If

    ' Commit the transaction - Makes Undo treat all changes as one step
    Call modUR_Transaction.CommitTransaction

Cleanup:
    ' Restore screen updating & calculations
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Call DisplayMessage("Adjustments applied successfully.")
    Exit Sub

ErrorHandler:
    ' Rollback transaction if an error occurs
    If modUR_Transaction.IsInTransaction() Then
        Call modUR_Transaction.RollbackTransaction
    End If
    Call LogAndHandleError("Adjustments_Click")
    Resume Cleanup
End Sub

Public Sub DisplayMessage(msg As String)
    Dim ws As Worksheet
    Dim shp As Shape
    
    ' Set reference to the sheet
    Set ws = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")

    ' Check if the shape exists before updating
    On Error Resume Next
    Set shp = ws.Shapes("lblMessage")
    On Error GoTo 0

    ' If shape exists, update the text
    If Not shp Is Nothing Then
        shp.TextFrame2.TextRange.Text = msg
    Else
        MsgBox "Error: lblMessage text box not found!", vbCritical, "DisplayMessage Error"
    End If
End Sub




