Attribute VB_Name = "modInvMan"
Public Sub AddGoodsReceived_Click()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rng As Range
    Dim receivedCol As Long, totalInvCol As Long, itemCodeCol As Long, itemNameCol As Long, lastEditedCol As Long
    Dim i As Long, rowCount As Long
    Dim LogEntries As Collection
    Dim insertedCount As Long

    On Error GoTo ErrorHandler

    Call modUR_Transaction.BeginTransaction

    Set ws = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
    Set tbl = ws.ListObjects("invSys")

    If tbl Is Nothing Or tbl.ListRows.Count = 0 Then
        MsgBox "No data in invSys table.", vbExclamation, "Error"
        GoTo Cleanup
    End If

    ' Get column indexes dynamically
    itemCodeCol = tbl.ListColumns("ITEM_CODE").Index
    itemNameCol = tbl.ListColumns("ITEM").Index
    receivedCol = tbl.ListColumns("RECEIVED").Index
    totalInvCol = tbl.ListColumns("TOTAL INV").Index
    lastEditedCol = tbl.ListColumns("LAST EDITED").Index

    rowCount = tbl.ListRows.Count
    Set rng = tbl.DataBodyRange

    Set LogEntries = New Collection

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    For i = 1 To rowCount
        Dim receivedVal As Variant
        receivedVal = rng.Cells(i, receivedCol).Value
        
        If IsNumeric(receivedVal) And receivedVal > 0 Then
            Dim oldTotalInv As Variant
            oldTotalInv = rng.Cells(i, totalInvCol).Value
            
            ' Update TOTAL INV
            rng.Cells(i, totalInvCol).Value = oldTotalInv + receivedVal
            
            ' Update LAST EDITED
            rng.Cells(i, lastEditedCol).Value = Now
            
            ' Track the change
            Call modUR_Transaction.TrackTransactionChange("CellUpdate", _
                rng.Cells(i, itemCodeCol).Value, "TOTAL INV", oldTotalInv, rng.Cells(i, totalInvCol).Value)
            
            ' Log the change
            LogEntries.Add Array(Environ("USERNAME"), "Added Goods Received", _
                rng.Cells(i, itemCodeCol).Value, rng.Cells(i, itemNameCol).Value, receivedVal, rng.Cells(i, totalInvCol).Value)
            
            ' Reset RECEIVED
            rng.Cells(i, receivedCol).Value = 0
        End If
    Next i

    If LogEntries.Count > 0 Then
        insertedCount = modInvLogs.LogMultipleInventoryChanges(LogEntries)
        Call modUR_Transaction.SetCurrentTransactionLogCount(insertedCount)
    End If

    Call modUR_Transaction.CommitTransaction

Cleanup:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Call DisplayMessage("Goods received successfully.")
    Exit Sub

ErrorHandler:
    If modUR_Transaction.IsInTransaction() Then
        Call modUR_Transaction.RollbackTransaction
    End If
    Call LogAndHandleError("AddGoodsReceived_Click")
    Resume Cleanup
End Sub

    Public Sub DeductUsed_Click()
        Dim ws As Worksheet
        Dim tbl As ListObject
        Dim rng As Range
        Dim usedCol As Long, totalInvCol As Long, itemCodeCol As Long, itemNameCol As Long, lastEditedCol As Long
        Dim i As Long, rowCount As Long
        Dim LogEntries As Collection
        Dim insertedCount As Long
    
        On Error GoTo ErrorHandler
    
        Call modUR_Transaction.BeginTransaction
    
        Set ws = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
        Set tbl = ws.ListObjects("invSys")
    
        If tbl Is Nothing Or tbl.ListRows.Count = 0 Then
            MsgBox "No data in invSys table.", vbExclamation, "Error"
            GoTo Cleanup
        End If
    
        ' Get column indexes dynamically
        itemCodeCol = tbl.ListColumns("ITEM_CODE").Index
        itemNameCol = tbl.ListColumns("ITEM").Index
        usedCol = tbl.ListColumns("USED").Index
        totalInvCol = tbl.ListColumns("TOTAL INV").Index
        lastEditedCol = tbl.ListColumns("LAST EDITED").Index
    
        rowCount = tbl.ListRows.Count
        Set rng = tbl.DataBodyRange
    
        Set LogEntries = New Collection
    
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
    
        For i = 1 To rowCount
            Dim usedVal As Variant
            usedVal = rng.Cells(i, usedCol).Value
            
            If IsNumeric(usedVal) And usedVal > 0 Then
                Dim oldTotalInv As Variant
                oldTotalInv = rng.Cells(i, totalInvCol).Value
                
                ' Update TOTAL INV
                rng.Cells(i, totalInvCol).Value = oldTotalInv - usedVal
                
                ' Update LAST EDITED
                rng.Cells(i, lastEditedCol).Value = Now
                
                ' Track the change
                Call modUR_Transaction.TrackTransactionChange("CellUpdate", _
                    rng.Cells(i, itemCodeCol).Value, "TOTAL INV", oldTotalInv, rng.Cells(i, totalInvCol).Value)
                
                ' Log the change
                LogEntries.Add Array(Environ("USERNAME"), "Deducted Used Items", _
                    rng.Cells(i, itemCodeCol).Value, rng.Cells(i, itemNameCol).Value, -usedVal, rng.Cells(i, totalInvCol).Value)
                
                ' Reset USED
                rng.Cells(i, usedCol).Value = 0
            End If
        Next i
    
        If LogEntries.Count > 0 Then
            insertedCount = modInvLogs.LogMultipleInventoryChanges(LogEntries)
            Call modUR_Transaction.SetCurrentTransactionLogCount(insertedCount)
        End If
    
        Call modUR_Transaction.CommitTransaction
    
    Cleanup:
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        Call DisplayMessage("Used items deducted successfully.")
        Exit Sub
    
    ErrorHandler:
        If modUR_Transaction.IsInTransaction() Then
            Call modUR_Transaction.RollbackTransaction
        End If
        Call LogAndHandleError("DeductUsed_Click")
        Resume Cleanup
    End Sub

        Public Sub DeductShipments_Click()
            Dim ws As Worksheet
            Dim tbl As ListObject
            Dim rng As Range
            Dim shipmentsCol As Long, totalInvCol As Long, itemCodeCol As Long, itemNameCol As Long, lastEditedCol As Long
            Dim i As Long, rowCount As Long
            Dim LogEntries As Collection
            Dim insertedCount As Long
        
            On Error GoTo ErrorHandler
        
            Call modUR_Transaction.BeginTransaction
        
            Set ws = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
            Set tbl = ws.ListObjects("invSys")
        
            If tbl Is Nothing Or tbl.ListRows.Count = 0 Then
                MsgBox "No data in invSys table.", vbExclamation, "Error"
                GoTo Cleanup
            End If
        
            ' Get column indexes dynamically
            itemCodeCol = tbl.ListColumns("ITEM_CODE").Index
            itemNameCol = tbl.ListColumns("ITEM").Index
            shipmentsCol = tbl.ListColumns("SHIPMENTS").Index
            totalInvCol = tbl.ListColumns("TOTAL INV").Index
            lastEditedCol = tbl.ListColumns("LAST EDITED").Index
        
            rowCount = tbl.ListRows.Count
            Set rng = tbl.DataBodyRange
        
            Set LogEntries = New Collection
        
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
            Application.EnableEvents = False
        
            For i = 1 To rowCount
                Dim shipmentsVal As Variant
                shipmentsVal = rng.Cells(i, shipmentsCol).Value
                
                If IsNumeric(shipmentsVal) And shipmentsVal > 0 Then
                    Dim oldTotalInv As Variant
                    oldTotalInv = rng.Cells(i, totalInvCol).Value
                    
                    ' Update TOTAL INV
                    rng.Cells(i, totalInvCol).Value = oldTotalInv - shipmentsVal
                    
                    ' Update LAST EDITED
                    rng.Cells(i, lastEditedCol).Value = Now
                    
                    ' Track the change
                    Call modUR_Transaction.TrackTransactionChange("CellUpdate", _
                        rng.Cells(i, itemCodeCol).Value, "TOTAL INV", oldTotalInv, rng.Cells(i, totalInvCol).Value)
                    
                    ' Log the change
                    LogEntries.Add Array(Environ("USERNAME"), "Shipments Deducted", _
                        rng.Cells(i, itemCodeCol).Value, rng.Cells(i, itemNameCol).Value, -shipmentsVal, rng.Cells(i, totalInvCol).Value)
                    
                    ' Reset SHIPMENTS
                    rng.Cells(i, shipmentsCol).Value = 0
                End If
            Next i
        
            If LogEntries.Count > 0 Then
                insertedCount = modInvLogs.LogMultipleInventoryChanges(LogEntries)
                Call modUR_Transaction.SetCurrentTransactionLogCount(insertedCount)
            End If
        
            Call modUR_Transaction.CommitTransaction
        
        Cleanup:
            Application.EnableEvents = True
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            Call DisplayMessage("Shipments deducted successfully.")
            Exit Sub
        
        ErrorHandler:
            If modUR_Transaction.IsInTransaction() Then
                Call modUR_Transaction.RollbackTransaction
            End If
            Call LogAndHandleError("DeductShipments_Click")
            Resume Cleanup
        End Sub

            Public Sub Adjustments_Click()
                Dim ws As Worksheet
                Dim tbl As ListObject
                Dim rng As Range
                Dim adjustmentsCol As Long, totalInvCol As Long, itemCodeCol As Long, itemNameCol As Long, lastEditedCol As Long
                Dim i As Long, rowCount As Long
                Dim LogEntries As Collection
                Dim insertedCount As Long
            
                On Error GoTo ErrorHandler
            
                Call modUR_Transaction.BeginTransaction
            
                Set ws = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
                Set tbl = ws.ListObjects("invSys")
            
                If tbl Is Nothing Or tbl.ListRows.Count = 0 Then
                    MsgBox "No data in invSys table.", vbExclamation, "Error"
                    GoTo Cleanup
                End If
            
                ' Get column indexes dynamically
                itemCodeCol = tbl.ListColumns("ITEM_CODE").Index
                itemNameCol = tbl.ListColumns("ITEM").Index
                adjustmentsCol = tbl.ListColumns("ADJUSTMENTS").Index
                totalInvCol = tbl.ListColumns("TOTAL INV").Index
                lastEditedCol = tbl.ListColumns("LAST EDITED").Index
            
                rowCount = tbl.ListRows.Count
                Set rng = tbl.DataBodyRange
            
                Set LogEntries = New Collection
            
                Application.ScreenUpdating = False
                Application.Calculation = xlCalculationManual
                Application.EnableEvents = False
            
                For i = 1 To rowCount
                    Dim adjustmentVal As Variant
                    adjustmentVal = rng.Cells(i, adjustmentsCol).Value
                    
                    If IsNumeric(adjustmentVal) And adjustmentVal <> 0 Then
                        Dim oldTotalInv As Variant
                        oldTotalInv = rng.Cells(i, totalInvCol).Value
                        
                        ' Update TOTAL INV (positive adds, negative subtracts)
                        rng.Cells(i, totalInvCol).Value = oldTotalInv + adjustmentVal
                        
                        ' Update LAST EDITED
                        rng.Cells(i, lastEditedCol).Value = Now
                        
                        ' Track the change
                        Call modUR_Transaction.TrackTransactionChange("CellUpdate", _
                            rng.Cells(i, itemCodeCol).Value, "TOTAL INV", oldTotalInv, rng.Cells(i, totalInvCol).Value)
                        
                        ' Log the change
                        LogEntries.Add Array(Environ("USERNAME"), "Inventory Adjustment", _
                            rng.Cells(i, itemCodeCol).Value, rng.Cells(i, itemNameCol).Value, adjustmentVal, rng.Cells(i, totalInvCol).Value)
                        
                        ' Reset ADJUSTMENTS
                        rng.Cells(i, adjustmentsCol).Value = 0
                    End If
                Next i
            
                If LogEntries.Count > 0 Then
                    insertedCount = modInvLogs.LogMultipleInventoryChanges(LogEntries)
                    Call modUR_Transaction.SetCurrentTransactionLogCount(insertedCount)
                End If
            
                Call modUR_Transaction.CommitTransaction
            
            Cleanup:
                Application.EnableEvents = True
                Application.Calculation = xlCalculationAutomatic
                Application.ScreenUpdating = True
                Call DisplayMessage("Adjustments applied successfully.")
                Exit Sub
            
            ErrorHandler:
                If modUR_Transaction.IsInTransaction() Then
                    Call modUR_Transaction.RollbackTransaction
                End If
                Call LogAndHandleError("Adjustments_Click")
                Resume Cleanup
            End Sub

    Public Sub AddMadeItems_Click()
        Dim ws As Worksheet
        Dim tbl As ListObject
        Dim rng As Range
        Dim madeCol As Long, totalInvCol As Long, itemCodeCol As Long, itemNameCol As Long, lastEditedCol As Long
        Dim i As Long, rowCount As Long
        Dim LogEntries As Collection
        Dim insertedCount As Long
    
        On Error GoTo ErrorHandler
    
        Call modUR_Transaction.BeginTransaction
    
        Set ws = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
        Set tbl = ws.ListObjects("invSys")
    
        If tbl Is Nothing Or tbl.ListRows.Count = 0 Then
            MsgBox "No data in invSys table.", vbExclamation, "Error"
            GoTo Cleanup
        End If
    
        ' Get column indexes dynamically
        itemCodeCol = tbl.ListColumns("ITEM_CODE").Index
        itemNameCol = tbl.ListColumns("ITEM").Index
        madeCol = tbl.ListColumns("MADE").Index
        totalInvCol = tbl.ListColumns("TOTAL INV").Index
        lastEditedCol = tbl.ListColumns("LAST EDITED").Index
    
        rowCount = tbl.ListRows.Count
        Set rng = tbl.DataBodyRange
    
        Set LogEntries = New Collection
    
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
    
        For i = 1 To rowCount
            Dim madeVal As Variant
            madeVal = rng.Cells(i, madeCol).Value
            
            If IsNumeric(madeVal) And madeVal > 0 Then
                Dim oldTotalInv As Variant
                oldTotalInv = rng.Cells(i, totalInvCol).Value
                
                ' Update TOTAL INV
                rng.Cells(i, totalInvCol).Value = oldTotalInv + madeVal
                
                ' Update LAST EDITED
                rng.Cells(i, lastEditedCol).Value = Now
                
                ' Track the change
                Call modUR_Transaction.TrackTransactionChange("CellUpdate", _
                    rng.Cells(i, itemCodeCol).Value, "TOTAL INV", oldTotalInv, rng.Cells(i, totalInvCol).Value)
                
                ' Log the change
                LogEntries.Add Array(Environ("USERNAME"), "Made Items Added", _
                    rng.Cells(i, itemCodeCol).Value, rng.Cells(i, itemNameCol).Value, madeVal, rng.Cells(i, totalInvCol).Value)
                
                ' Reset MADE
                rng.Cells(i, madeCol).Value = 0
            End If
        Next i
    
        If LogEntries.Count > 0 Then
            insertedCount = modInvLogs.LogMultipleInventoryChanges(LogEntries)
            Call modUR_Transaction.SetCurrentTransactionLogCount(insertedCount)
        End If
    
        Call modUR_Transaction.CommitTransaction
    
    Cleanup:
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        Call DisplayMessage("Made items added successfully.")
        Exit Sub
    
    ErrorHandler:
        If modUR_Transaction.IsInTransaction() Then
            Call modUR_Transaction.RollbackTransaction
        End If
        Call LogAndHandleError("AddMadeItems_Click")
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




