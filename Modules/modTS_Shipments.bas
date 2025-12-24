Attribute VB_Name = "modTS_Shipments"
Option Explicit

' =============================================================
' Module: modTS_Shipments
' Purpose: All logic for the ShippingTally system (box builder,
'          holding subsystem, confirm/build/ship macros, logging).
' Notes:
'   - Buttons are generated dynamically (similar to modTS_Received).
'   - ShippingBOM sheet stores one ListObject per BOM (Box Name).
'   - BOM entries store ROW/QUANTITY/UOM only; item metadata is
'     resolved from invSys (InventoryManagement!invSys).
'   - Hold subsystem keeps packages on NotShipped until released.
'   - Additional confirm/build/ship routines will be implemented in
'     subsequent iterations (placeholders provided below).
' =============================================================

' ===== constants =====
Private Const SHEET_SHIPMENTS As String = "ShipmentsTally"
Private Const SHEET_INV As String = "InventoryManagement"
Private Const SHEET_BOM As String = "ShippingBOM"

Private Const TABLE_SHIPMENTS As String = "ShipmentsTally"
Private Const TABLE_NOTSHIPPED As String = "NotShipped"
Private Const TABLE_AGG_BOM As String = "AggregateBoxBOM"
Private Const TABLE_AGG_PACK As String = "AggregatePackages"
Private Const TABLE_BOX_BUILDER As String = "BoxBuilder"
Private Const TABLE_BOX_BOM As String = "BoxBOM"
Private Const TABLE_CHECK_INV As String = "Check_invSys"

Private Const BTN_SHOW_BUILDER As String = "BTN_SHOW_BUILDER"
Private Const BTN_HIDE_BUILDER As String = "BTN_HIDE_BUILDER"
Private Const BTN_SAVE_BOX As String = "BTN_SAVE_BOX"
Private Const BTN_UNSHIP As String = "BTN_UNSHIP"
Private Const BTN_SEND_HOLD As String = "BTN_SEND_HOLD"
Private Const BTN_RETURN_HOLD As String = "BTN_RETURN_HOLD"
Private Const BTN_CONFIRM_INV As String = "BTN_CONFIRM_INV"
Private Const BTN_BOXES_MADE As String = "BTN_BOXES_MADE"
Private Const BTN_TO_TOTALINV As String = "BTN_TO_TOTALINV"
Private Const BTN_TO_SHIPMENTS As String = "BTN_TO_SHIPMENTS"
Private Const BTN_SHIPMENTS_SENT As String = "BTN_SHIPMENTS_SENT"

Private Const SHIPPING_BOM_BLOCK_ROWS As Long = 52
Private Const SHIPPING_BOM_DATA_ROWS As Long = 50
Private Const SHIPPING_BOM_COLS As Long = 3 ' ROW, QUANTITY, UOM

' ===== public entry points =====
Public Sub InitializeShipmentsUI()
    EnsureShipmentsButtons
End Sub

Public Sub BtnShowBuilder()
    ToggleBuilderTables True
End Sub

Public Sub BtnHideBuilder()
    ToggleBuilderTables False
End Sub

Public Sub BtnSaveBox()
    On Error GoTo ErrHandler
    Dim ws As Worksheet: Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Sub

    Dim loMeta As ListObject: Set loMeta = GetListObject(ws, TABLE_BOX_BUILDER)
    Dim loBom As ListObject: Set loBom = GetListObject(ws, TABLE_BOX_BOM)
    If loMeta Is Nothing Or loBom Is Nothing Then
        MsgBox "Box Builder tables not found on ShipmentsTally sheet.", vbExclamation
        Exit Sub
    End If

    Dim boxName As String
    boxName = Trim$(NzStr(ValueFromTable(loMeta, "Box Name")))
    If boxName = "" Then
        MsgBox "Enter a Box Name before saving.", vbExclamation
        Exit Sub
    End If

    Dim components As Collection
    Set components = ReadBoxBomComponents(loBom)
    If components.count = 0 Then
        MsgBox "Add at least one component to the BoxBOM table.", vbExclamation
        Exit Sub
    End If
    If components.count > SHIPPING_BOM_DATA_ROWS Then
        MsgBox "BOM exceeds the 50-row limit. Remove extra rows and try again.", vbExclamation
        Exit Sub
    End If

    Dim wsBOM As Worksheet: Set wsBOM = SheetExists(SHEET_BOM)
    If wsBOM Is Nothing Then
        MsgBox "ShippingBOM sheet not found.", vbCritical
        Exit Sub
    End If

    Dim bomTable As ListObject, blockRange As Range
    Set bomTable = EnsureBomTable(wsBOM, boxName, blockRange)
    If bomTable Is Nothing Then Exit Sub

    WriteBomData bomTable, blockRange, components

    MsgBox "Saved BOM '" & boxName & "' (" & components.count & " items).", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "BTN_SAVE_BOX failed: " & Err.Description, vbCritical
End Sub

Public Sub BtnUnship()
    Dim ws As Worksheet: Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Sub
    Dim lo As ListObject: Set lo = GetListObject(ws, TABLE_NOTSHIPPED)
    If lo Is Nothing Then
        MsgBox "NotShipped table not found.", vbExclamation
        Exit Sub
    End If
    Dim isHidden As Boolean
    isHidden = lo.Range.EntireColumn.Hidden
    lo.Range.EntireColumn.Hidden = Not isHidden
End Sub

Public Sub BtnSendHold()
    MoveSelectionToHold True
End Sub

Public Sub BtnReturnHold()
    MoveSelectionToHold False
End Sub

Public Sub BtnConfirmInventory()
    ' Placeholder for full confirm workflow
    MsgBox "BTN_CONFIRM_INV logic pending implementation.", vbInformation
End Sub

Public Sub BtnBoxesMade()
    ' Placeholder for BOM build workflow
    MsgBox "BTN_BOXES_MADE logic pending implementation.", vbInformation
End Sub

Public Sub BtnToTotalInv()
    MsgBox "BTN_TO_TOTALINV logic pending implementation.", vbInformation
End Sub

Public Sub BtnToShipments()
    MsgBox "BTN_TO_SHIPMENTS logic pending implementation.", vbInformation
End Sub

Public Sub BtnShipmentsSent()
    MsgBox "BTN_SHIPMENTS_SENT logic pending implementation.", vbInformation
End Sub

' ===== button scaffolding =====
Private Sub EnsureShipmentsButtons()
    Dim ws As Worksheet: Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Sub

    Dim leftA As Double: leftA = ws.Columns("A").Left + 4
    Dim nextTop As Double: nextTop = ws.Rows(2).Top

    EnsureButtonCustom ws, BTN_SHOW_BUILDER, "Show builder", "modTS_Shipments.BtnShowBuilder", leftA, nextTop
    nextTop = nextTop + 22
    EnsureButtonCustom ws, BTN_HIDE_BUILDER, "Hide builder", "modTS_Shipments.BtnHideBuilder", leftA, nextTop
    nextTop = nextTop + 22
    EnsureButtonCustom ws, BTN_SAVE_BOX, "Save box", "modTS_Shipments.BtnSaveBox", leftA, nextTop
    nextTop = nextTop + 28
    EnsureButtonCustom ws, BTN_CONFIRM_INV, "Confirm inventory", "modTS_Shipments.BtnConfirmInventory", leftA, nextTop
    nextTop = nextTop + 22
    EnsureButtonCustom ws, BTN_BOXES_MADE, "Boxes made", "modTS_Shipments.BtnBoxesMade", leftA, nextTop
    nextTop = nextTop + 22
    EnsureButtonCustom ws, BTN_TO_TOTALINV, "To TotalInv", "modTS_Shipments.BtnToTotalInv", leftA, nextTop
    nextTop = nextTop + 22
    EnsureButtonCustom ws, BTN_TO_SHIPMENTS, "To Shipments", "modTS_Shipments.BtnToShipments", leftA, nextTop
    nextTop = nextTop + 22
    EnsureButtonCustom ws, BTN_SHIPMENTS_SENT, "Shipments sent", "modTS_Shipments.BtnShipmentsSent", leftA, nextTop

    Dim loHold As ListObject: Set loHold = GetListObject(ws, TABLE_NOTSHIPPED)
    If Not loHold Is Nothing Then
        Dim topBand As Double
        topBand = loHold.HeaderRowRange.Top - 24
        Dim leftBand As Double
        leftBand = loHold.HeaderRowRange.Left
        EnsureButtonCustom ws, BTN_UNSHIP, "Toggle NotShipped", "modTS_Shipments.BtnUnship", leftBand, topBand
        EnsureButtonCustom ws, BTN_SEND_HOLD, "Send to hold", "modTS_Shipments.BtnSendHold", leftBand + 120, topBand
        EnsureButtonCustom ws, BTN_RETURN_HOLD, "Return from hold", "modTS_Shipments.BtnReturnHold", leftBand + 240, topBand
    End If
End Sub

Private Sub EnsureButtonCustom(ws As Worksheet, shapeName As String, caption As String, onActionMacro As String, leftPos As Double, topPos As Double)
    Const BTN_WIDTH As Double = 118
    Const BTN_HEIGHT As Double = 20
    Dim shp As Shape
    On Error Resume Next
    Set shp = ws.Shapes(shapeName)
    On Error GoTo 0
    If shp Is Nothing Then
        Set shp = ws.Shapes.AddFormControl(xlButtonControl, leftPos, topPos, BTN_WIDTH, BTN_HEIGHT)
        shp.Name = shapeName
        shp.TextFrame.Characters.Text = caption
        shp.OnAction = onActionMacro
    Else
        shp.Left = leftPos
        shp.Top = topPos
        shp.Width = BTN_WIDTH
        shp.Height = BTN_HEIGHT
        shp.TextFrame.Characters.Text = caption
        shp.OnAction = onActionMacro
    End If
End Sub

' ===== builder helpers =====
Private Sub ToggleBuilderTables(ByVal makeVisible As Boolean)
    Dim ws As Worksheet: Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Sub
    Dim lo1 As ListObject: Set lo1 = GetListObject(ws, TABLE_BOX_BUILDER)
    Dim lo2 As ListObject: Set lo2 = GetListObject(ws, TABLE_BOX_BOM)
    If lo1 Is Nothing Or lo2 Is Nothing Then Exit Sub
    lo1.Range.EntireRow.Hidden = Not makeVisible
    lo2.Range.EntireRow.Hidden = Not makeVisible
End Sub

Private Function ReadBoxBomComponents(loBom As ListObject) As Collection
    Dim result As New Collection
    If loBom Is Nothing Then
        Set ReadBoxBomComponents = result
        Exit Function
    End If

    Dim cRow As Long: cRow = ColumnIndex(loBom, "ROW")
    Dim cQty As Long: cQty = ColumnIndex(loBom, "QUANTITY")
    Dim cUOM As Long: cUOM = ColumnIndex(loBom, "UOM")
    If cRow = 0 Or cQty = 0 Or cUOM = 0 Then
        MsgBox "BoxBOM table must contain ROW, QUANTITY, and UOM columns.", vbExclamation
        Exit Function
    End If

    If loBom.DataBodyRange Is Nothing Then
        Set ReadBoxBomComponents = result
        Exit Function
    End If

    Dim arr As Variant: arr = loBom.DataBodyRange.Value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        Dim partRow As Long: partRow = NzLng(arr(r, cRow))
        Dim qty As Double: qty = NzDbl(arr(r, cQty))
        Dim partUom As String: partUom = Trim$(NzStr(arr(r, cUOM)))
        If partRow > 0 And qty > 0 And partUom <> "" Then
            Dim info(1 To 3) As Variant
            info(1) = partRow
            info(2) = qty
            info(3) = partUom
            result.Add info
        End If
    Next
    Set ReadBoxBomComponents = result
End Function

Private Function EnsureBomTable(ws As Worksheet, ByVal boxName As String, ByRef blockRange As Range) As ListObject
    Dim cleanName As String: cleanName = SafeTableName(boxName)

    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(cleanName)
    On Error GoTo 0
    If Not lo Is Nothing Then
        Set blockRange = BlockRangeFromHeader(ws, lo.HeaderRowRange.Row)
        If blockRange Is Nothing Then
            Set blockRange = lo.Range
        End If
        lo.Resize blockRange
        lo.HeaderRowRange.Cells(1, 1).Value = "ROW"
        lo.HeaderRowRange.Cells(1, 2).Value = "QUANTITY"
        lo.HeaderRowRange.Cells(1, 3).Value = "UOM"
        Set EnsureBomTable = lo
        Exit Function
    End If

    Dim startRow As Long: startRow = NextAvailableBomRow(ws)
    If startRow = 0 Then
        MsgBox "ShippingBOM sheet has no space for additional BOMs.", vbCritical
        Exit Function
    End If
    Set blockRange = ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow + SHIPPING_BOM_DATA_ROWS, SHIPPING_BOM_COLS))
    blockRange.Clear
    blockRange.Rows(1).Cells(1, 1).Value = "ROW"
    blockRange.Rows(1).Cells(1, 2).Value = "QUANTITY"
    blockRange.Rows(1).Cells(1, 3).Value = "UOM"
    Set lo = ws.ListObjects.Add(xlSrcRange, blockRange, , xlYes)
    lo.Name = cleanName
    Set EnsureBomTable = lo
End Function

Private Sub WriteBomData(lo As ListObject, blockRange As Range, comps As Collection)
    If lo Is Nothing Then Exit Sub
    If blockRange Is Nothing Then Set blockRange = lo.Range
    lo.Resize blockRange
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.ClearContents

    If comps.count = 0 Then Exit Sub
    Dim arr() As Variant
    ReDim arr(1 To SHIPPING_BOM_DATA_ROWS, 1 To SHIPPING_BOM_COLS)

    Dim i As Long
    For i = 1 To comps.count
        Dim info As Variant
        info = comps(i)
        arr(i, 1) = info(1)
        arr(i, 2) = info(2)
        arr(i, 3) = info(3)
    Next

    Dim dataRange As Range
    Set dataRange = lo.DataBodyRange.Resize(SHIPPING_BOM_DATA_ROWS, SHIPPING_BOM_COLS)
    dataRange.Value = arr
End Sub

Private Function NextAvailableBomRow(ws As Worksheet) As Long
    Dim totalRows As Long: totalRows = ws.Rows.Count
    Dim startRow As Long
    startRow = 1
    Do
        If startRow + SHIPPING_BOM_BLOCK_ROWS - 1 > totalRows Then
            NextAvailableBomRow = 0
            Exit Function
        End If
        If IsBlockFree(ws, startRow) Then
            NextAvailableBomRow = startRow
            Exit Function
        End If
        startRow = startRow + SHIPPING_BOM_BLOCK_ROWS
    Loop
End Function

Private Function IsBlockFree(ws As Worksheet, startRow As Long) As Boolean
    Dim rg As Range
    Set rg = BlockRangeFromHeader(ws, startRow)
    If rg Is Nothing Then
        IsBlockFree = False
        Exit Function
    End If
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        If Not Intersect(lo.Range, rg) Is Nothing Then
            IsBlockFree = False
            Exit Function
        End If
    Next
    IsBlockFree = True
End Function

Private Function BlockRangeFromHeader(ws As Worksheet, startRow As Long) As Range
    On Error Resume Next
    Set BlockRangeFromHeader = ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow + SHIPPING_BOM_DATA_ROWS, SHIPPING_BOM_COLS))
    On Error GoTo 0
End Function

Private Function SafeTableName(ByVal sourceName As String) As String
    Dim cleaned As String
    cleaned = Trim$(sourceName)
    If cleaned = "" Then cleaned = "BOM_" & Format(Now, "yyyymmdd_hhnnss")
    Dim i As Long, ch As String, kept As String
    For i = 1 To Len(cleaned)
        ch = Mid$(cleaned, i, 1)
        If ch Like "[A-Za-z0-9_]" Then
            kept = kept & ch
        Else
            kept = kept & "_"
        End If
    Next
    If kept = "" Then kept = "BOM_" & Format(Now, "yyyymmdd_hhnnss")
    If Not kept Like "[A-Za-z_]*" Then kept = "BOM_" & kept
    SafeTableName = kept
End Function

Private Function ValueFromTable(lo As ListObject, headerName As String) As Variant
    Dim colIdx As Long: colIdx = ColumnIndex(lo, headerName)
    If colIdx = 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    ValueFromTable = lo.DataBodyRange.Cells(1, colIdx).Value
End Function

' ===== hold helpers =====
Private Sub MoveSelectionToHold(ByVal moveToHold As Boolean)
    Dim ws As Worksheet: Set ws = SheetExists(SHEET_SHIPMENTS)
    If ws Is Nothing Then Exit Sub
    Dim loShip As ListObject: Set loShip = GetListObject(ws, TABLE_SHIPMENTS)
    Dim loHold As ListObject: Set loHold = GetListObject(ws, TABLE_NOTSHIPPED)
    If loShip Is Nothing Or loHold Is Nothing Then Exit Sub
    If loShip.DataBodyRange Is Nothing Then Exit Sub

    Dim targetTable As ListObject
    Dim sourceTable As ListObject
    If moveToHold Then
        Set sourceTable = loShip
        Set targetTable = loHold
    Else
        Set sourceTable = loHold
        Set targetTable = loShip
    End If

    Dim rngSel As Range
    On Error Resume Next
    Set rngSel = Application.Intersect(Application.Selection, sourceTable.DataBodyRange)
    On Error GoTo 0
    If rngSel Is Nothing Then
        MsgBox "Select rows inside the " & sourceTable.Name & " table first.", vbInformation
        Exit Sub
    End If

    Dim processed As Object: Set processed = CreateObject("Scripting.Dictionary")
    Dim cell As Range
    For Each cell In rngSel.Areas
        Dim r As Range
        For Each r In cell.Rows
            Dim rowIndex As Long
            rowIndex = r.Row - sourceTable.DataBodyRange.Row + 1
            If rowIndex >= 1 And rowIndex <= sourceTable.ListRows.Count Then
                If Not processed.Exists(rowIndex) Then
                    processed(rowIndex) = True
                    HandleHoldRow sourceTable, targetTable, rowIndex, moveToHold
                End If
            End If
        Next r
    Next cell
End Sub

Private Sub HandleHoldRow(sourceTable As ListObject, targetTable As ListObject, rowIndex As Long, moveToHold As Boolean)
    Dim cRef As Long: cRef = ColumnIndex(sourceTable, "REF_NUMBER")
    Dim cItems As Long: cItems = ColumnIndex(sourceTable, "ITEMS")
    Dim cQty As Long: cQty = ColumnIndex(sourceTable, "QUANTITY")
    If cQty = 0 Then
        MsgBox sourceTable.Name & " table needs a QUANTITY column.", vbCritical
        Exit Sub
    End If

    Dim refVal As String: refVal = NzStr(sourceTable.DataBodyRange.Cells(rowIndex, cRef).Value)
    Dim itemVal As String: itemVal = NzStr(sourceTable.DataBodyRange.Cells(rowIndex, cItems).Value)
    Dim qtyVal As Double: qtyVal = NzDbl(sourceTable.DataBodyRange.Cells(rowIndex, cQty).Value)
    If qtyVal <= 0 Then Exit Sub

    Dim prompt As String
    If moveToHold Then
        prompt = "Enter quantity to hold for '" & itemVal & "' (available " & qtyVal & "):"
    Else
        prompt = "Enter quantity to return to shipments for '" & itemVal & "' (available " & qtyVal & "):"
    End If
    Dim qtyInput As Variant
    qtyInput = Application.InputBox(prompt, "Hold quantity", qtyVal, Type:=1)
    If qtyInput = False Then Exit Sub
    Dim qtyMove As Double: qtyMove = CDbl(qtyInput)
    If qtyMove <= 0 Then Exit Sub
    If qtyMove > qtyVal Then qtyMove = qtyVal

    AppendHoldRow targetTable, refVal, itemVal, qtyMove

    Dim newQty As Double
    If moveToHold Then
        newQty = qtyVal - qtyMove
    Else
        newQty = qtyVal - qtyMove
    End If
    If newQty <= 0 Then
        sourceTable.ListRows(rowIndex).Range.ClearContents
    Else
        sourceTable.DataBodyRange.Cells(rowIndex, cQty).Value = newQty
    End If
End Sub

Private Sub AppendHoldRow(targetTable As ListObject, refVal As String, itemVal As String, qtyMove As Double)
    Dim cRef As Long: cRef = ColumnIndex(targetTable, "REF_NUMBER")
    Dim cItems As Long: cItems = ColumnIndex(targetTable, "ITEMS")
    Dim cQty As Long: cQty = ColumnIndex(targetTable, "QUANTITY")
    Dim lr As ListRow: Set lr = targetTable.ListRows.Add
    If cRef > 0 Then lr.Range.Cells(1, cRef).Value = refVal
    If cItems > 0 Then lr.Range.Cells(1, cItems).Value = itemVal
    If cQty > 0 Then lr.Range.Cells(1, cQty).Value = qtyMove
End Sub

' ===== helpers reused from modTS_Received =====
Private Function SheetExists(nameOrCode As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, nameOrCode, vbTextCompare) = 0 _
           Or StrComp(ws.CodeName, nameOrCode, vbTextCompare) = 0 Then
            Set SheetExists = ws
            Exit Function
        End If
    Next ws
End Function

Private Function GetListObject(ws As Worksheet, tableName As String) As ListObject
    On Error Resume Next
    Set GetListObject = ws.ListObjects(tableName)
    On Error GoTo 0
End Function

Private Function ColumnIndex(lo As ListObject, colName As String) As Long
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If StrComp(lc.Name, colName, vbTextCompare) = 0 Then
            ColumnIndex = lc.Index
            Exit Function
        End If
    Next lc
    ColumnIndex = 0
End Function

Public Function NzStr(v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        NzStr = ""
    Else
        NzStr = CStr(v)
    End If
End Function

Public Function NzDbl(v As Variant) As Double
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Or v = "" Then
        NzDbl = 0#
    Else
        NzDbl = CDbl(v)
    End If
End Function

Public Function NzLng(v As Variant) As Long
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Or v = "" Then
        NzLng = 0
    Else
        NzLng = CLng(v)
    End If
End Function
