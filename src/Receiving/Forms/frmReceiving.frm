VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReceiving
   Caption         =   "Receiving"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmReceiving"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@RuntimeStubUserFormCode
Option Explicit

Private WithEvents mTxtRef As MSForms.TextBox
Private WithEvents mTxtReceiptId As MSForms.TextBox
Private WithEvents mTxtSearch As MSForms.TextBox
Private WithEvents mTxtQty As MSForms.TextBox
Private WithEvents mBtnRefresh As MSForms.CommandButton
Private WithEvents mBtnAdd As MSForms.CommandButton
Private WithEvents mBtnConfirm As MSForms.CommandButton
Private WithEvents mBtnUndo As MSForms.CommandButton
Private WithEvents mBtnRedo As MSForms.CommandButton
Private WithEvents mBtnClear As MSForms.CommandButton
Private WithEvents mBtnClose As MSForms.CommandButton
Private WithEvents mLstInventory As MSForms.ListBox
Private WithEvents mLstStaged As MSForms.ListBox
Private WithEvents mLstAggregate As MSForms.ListBox

Private mLblRef As MSForms.Label
Private mLblReceiptId As MSForms.Label
Private mLblSearch As MSForms.Label
Private mLblQty As MSForms.Label
Private mLblInventoryTitle As MSForms.Label
Private mLblInventoryHeader As MSForms.Label
Private mLblStagedTitle As MSForms.Label
Private mLblStagedHeader As MSForms.Label
Private mLblAggregateTitle As MSForms.Label
Private mLblAggregateHeader As MSForms.Label
Private mTxtStatus As MSForms.TextBox
Private mOperatorWorkbook As Workbook
Private mInventoryRows As Variant
Private mBuilt As Boolean
Private mLoading As Boolean
Private mResizeInitialized As Boolean
Private mResizing As Boolean

Private Const RECEIVING_BASE_WIDTH As Double = 1020
Private Const RECEIVING_BASE_HEIGHT As Double = 650

Private Sub UserForm_Initialize()
    BuildLayout
End Sub

Private Sub UserForm_Activate()
    If Not mResizeInitialized Then
        On Error Resume Next
        Application.Run "modUserFormResizeWin.EnableResizableUserForm", Me, True, True
        On Error GoTo 0
        mResizeInitialized = True
    End If
    ResizeReceivingLayout
End Sub

Private Sub UserForm_Resize()
    ResizeReceivingLayout
End Sub

Private Sub UserForm_Terminate()
    Set mOperatorWorkbook = Nothing
End Sub

Public Sub SetOperatorWorkbook(ByVal wb As Workbook)
    If Not wb Is Nothing Then
        If Not wb.IsAddin Then Set mOperatorWorkbook = wb
    End If
End Sub

Public Sub InitializeFromReceiving()
    On Error GoTo ErrHandler

    Dim wb As Workbook
    If Not mBuilt Then BuildLayout
    Set wb = ResolveOperatorWorkbook()
    If wb Is Nothing Then
        ShowStatus "Open a Receiving operator workbook before using the Receiving form."
        Exit Sub
    End If

    mLoading = True
    ActivateOperatorWorkbook
    modTS_Received.InitializeReceivingUiForWorkbook wb
    If Trim$(CStr(mTxtReceiptId.Value)) = "" Then mTxtReceiptId.Value = DefaultReceiptId()
    mTxtRef.Value = ""
    RefreshAllViews
    mLoading = False
    ShowStatus "Receiving form loaded for " & wb.Name & "."
    Exit Sub

ErrHandler:
    mLoading = False
    ShowStatus "Receiving form load failed: " & Err.Description
End Sub

Public Function TestSearchInventoryCount(ByVal values As Variant, ByVal filterText As String) As Long
    If Not mBuilt Then BuildLayout
    mInventoryRows = values
    FillInventoryList FilterInventoryRows(filterText)
    TestSearchInventoryCount = mLstInventory.ListCount
End Function

Public Function TestReceiptIdSeparatedFromReference() As Long
    If Not mBuilt Then BuildLayout
    mTxtReceiptId.Value = DefaultReceiptId()
    mTxtRef.Value = ""
    If Left$(CStr(mTxtReceiptId.Value), 4) = "RCV-" _
       And Trim$(CStr(mTxtRef.Value)) = "" Then
        TestReceiptIdSeparatedFromReference = 1
    End If
End Function

Public Function TestInitializeForWorkbook(ByVal operatorWb As Workbook) As String
    If operatorWb Is Nothing Then Exit Function
    If Not mBuilt Then BuildLayout
    SetOperatorWorkbook operatorWb
    TestInitializeForWorkbook = _
        "OK|BoundWorkbook=" & mOperatorWorkbook.Name & _
        "|Caption=" & Me.Caption
End Function

Public Function TestRunConfirmWritesActionForWorkbook(ByVal operatorWb As Workbook, _
                                                       Optional ByVal activatedWb As Workbook = Nothing) As String
    If Not mBuilt Then BuildLayout
    SetOperatorWorkbook operatorWb
    If Not activatedWb Is Nothing Then activatedWb.Activate
    mBtnConfirm_Click
    TestRunConfirmWritesActionForWorkbook = _
        "Succeeded=" & CStr(modTS_Received.LastConfirmWritesSucceeded()) & _
        "; Status=" & modTS_Received.LastConfirmWritesStatus() & _
        "; BoundWorkbook=" & mOperatorWorkbook.Name
End Function

Private Sub BuildLayout()
    If mBuilt Then Exit Sub

    Me.Caption = "Receiving"
    Me.Width = RECEIVING_BASE_WIDTH
    Me.Height = RECEIVING_BASE_HEIGHT
    Me.ScrollBars = fmScrollBarsBoth
    Me.KeepScrollBarsVisible = fmScrollBarsNone
    Me.ScrollWidth = RECEIVING_BASE_WIDTH - 20
    Me.ScrollHeight = RECEIVING_BASE_HEIGHT - 35

    Set mLblReceiptId = AddLabel("lblReceiptId", "Receipt ID", 18, 18, 70, 18, True)
    Set mTxtReceiptId = AddTextBox("txtReceiptId", 90, 16, 150, 22)
    mTxtReceiptId.Locked = True
    mTxtReceiptId.BackColor = &HEFEFEF
    Set mLblRef = AddLabel("lblRef", "PO/BOL Ref", 258, 18, 78, 18, True)
    Set mTxtRef = AddTextBox("txtRef", 338, 16, 150, 22)
    Set mLblSearch = AddLabel("lblSearch", "Search inventory", 506, 18, 110, 18, True)
    Set mTxtSearch = AddTextBox("txtSearch", 618, 16, 190, 22)
    Set mLblQty = AddLabel("lblQty", "Qty", 826, 18, 34, 18, True)
    Set mTxtQty = AddTextBox("txtQty", 862, 16, 80, 22)
    mTxtQty.Value = "1"

    Set mBtnRefresh = AddButton("btnRefresh", "Refresh", 778, 48, 96, 28)
    Set mBtnAdd = AddButton("btnAdd", "Add", 884, 48, 98, 28)
    Set mBtnConfirm = AddButton("btnConfirm", "Confirm Writes", 18, 580, 110, 30)
    Set mBtnUndo = AddButton("btnUndo", "Undo", 136, 580, 72, 30)
    Set mBtnRedo = AddButton("btnRedo", "Redo", 216, 580, 72, 30)
    Set mBtnClear = AddButton("btnClear", "Clear", 296, 580, 72, 30)
    Set mBtnClose = AddButton("btnClose", "Close", 892, 580, 90, 30)

    Set mLblInventoryTitle = AddLabel("lblInventoryTitle", "Inventory", 18, 56, 130, 18, True)
    Set mLblInventoryHeader = AddLabel("lblInventoryHeader", "ROW     Code          Item                         UOM   Inv       Location      Description / Vendor", 18, 78, 930, 16, False)
    Set mLstInventory = AddListBox("lstInventory", 18, 96, 964, 206, 9, "42 pt;82 pt;190 pt;42 pt;62 pt;80 pt;230 pt;120 pt;0 pt")

    Set mLblStagedTitle = AddLabel("lblStagedTitle", "Received Tally", 18, 318, 150, 18, True)
    Set mLblStagedHeader = AddLabel("lblStagedHeader", "Ref number             Item                                      Qty        ROW", 18, 340, 520, 16, False)
    Set mLstStaged = AddListBox("lstStaged", 18, 358, 520, 190, 4, "130 pt;250 pt;70 pt;50 pt")

    Set mLblAggregateTitle = AddLabel("lblAggregateTitle", "Aggregate Received", 560, 318, 160, 18, True)
    Set mLblAggregateHeader = AddLabel("lblAggregateHeader", "Code          Vendor        Vendor code   Description              Item                    UOM   Qty    Location   ROW   Ref", 560, 340, 420, 16, False)
    Set mLstAggregate = AddListBox("lstAggregate", 560, 358, 422, 190, 10, "84 pt;88 pt;80 pt;150 pt;160 pt;42 pt;58 pt;75 pt;45 pt;90 pt")

    Set mTxtStatus = AddTextBox("txtStatus", 18, 620, 964, 34)
    With mTxtStatus
        .Locked = True
        .MultiLine = True
        .EnterKeyBehavior = False
        .ScrollBars = fmScrollBarsVertical
        .BackColor = &HFFFFFF
    End With

    mBuilt = True
    ResizeReceivingLayout
End Sub

Private Function AddLabel(ByVal name As String, ByVal captionText As String, ByVal leftPos As Double, _
                          ByVal topPos As Double, ByVal widthVal As Double, ByVal heightVal As Double, _
                          ByVal boldText As Boolean) As MSForms.Label
    Dim lbl As MSForms.Label
    Set lbl = Me.Controls.Add("Forms.Label.1", name, True)
    With lbl
        .Caption = captionText
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
        .Font.Bold = boldText
    End With
    Set AddLabel = lbl
End Function

Private Function AddTextBox(ByVal name As String, ByVal leftPos As Double, ByVal topPos As Double, _
                            ByVal widthVal As Double, ByVal heightVal As Double) As MSForms.TextBox
    Dim txt As MSForms.TextBox
    Set txt = Me.Controls.Add("Forms.TextBox.1", name, True)
    With txt
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
    End With
    Set AddTextBox = txt
End Function

Private Function AddButton(ByVal name As String, ByVal captionText As String, ByVal leftPos As Double, _
                           ByVal topPos As Double, ByVal widthVal As Double, ByVal heightVal As Double) As MSForms.CommandButton
    Dim btn As MSForms.CommandButton
    Set btn = Me.Controls.Add("Forms.CommandButton.1", name, True)
    With btn
        .Caption = captionText
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
    End With
    Set AddButton = btn
End Function

Private Function AddListBox(ByVal name As String, ByVal leftPos As Double, ByVal topPos As Double, _
                            ByVal widthVal As Double, ByVal heightVal As Double, ByVal colCount As Long, _
                            ByVal widths As String) As MSForms.ListBox
    Dim lst As MSForms.ListBox
    Set lst = Me.Controls.Add("Forms.ListBox.1", name, True)
    With lst
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
        .ColumnCount = colCount
        .ColumnWidths = widths
        .IntegralHeight = False
    End With
    Set AddListBox = lst
End Function

Private Sub ResizeReceivingLayout()
    If mResizing Or Not mBuilt Then Exit Sub
    mResizing = True
    On Error GoTo Done

    Dim layoutW As Double
    Dim layoutH As Double
    Dim margin As Double
    Dim topArea As Double
    Dim bottomTop As Double
    Dim bottomH As Double
    Dim leftW As Double
    Dim rightW As Double
    Dim buttonTop As Double
    Dim statusTop As Double

    margin = 18
    layoutW = MaxDoubleReceiving(RECEIVING_BASE_WIDTH - 20, Me.InsideWidth)
    layoutH = MaxDoubleReceiving(RECEIVING_BASE_HEIGHT - 35, Me.InsideHeight)
    Me.ScrollWidth = layoutW
    Me.ScrollHeight = layoutH

    mTxtSearch.Width = MaxDoubleReceiving(160, layoutW - mTxtSearch.Left - 206)
    mLblQty.Left = mTxtSearch.Left + mTxtSearch.Width + 16
    mTxtQty.Left = mLblQty.Left + 36
    mBtnRefresh.Left = layoutW - 224
    mBtnAdd.Left = layoutW - 116

    topArea = MaxDoubleReceiving(180, (layoutH - 145) * 0.46)
    mLblInventoryHeader.Width = layoutW - (margin * 2)
    mLstInventory.Width = layoutW - (margin * 2)
    mLstInventory.Height = topArea

    bottomTop = mLstInventory.Top + mLstInventory.Height + 56
    bottomH = MaxDoubleReceiving(120, layoutH - bottomTop - 98)
    leftW = MaxDoubleReceiving(410, (layoutW - (margin * 3)) * 0.54)
    rightW = layoutW - leftW - (margin * 3)

    mLblStagedTitle.Top = bottomTop - 40
    mLblStagedHeader.Top = bottomTop - 18
    mLblStagedHeader.Width = leftW
    mLstStaged.Top = bottomTop
    mLstStaged.Width = leftW
    mLstStaged.Height = bottomH

    mLblAggregateTitle.Left = margin + leftW + margin
    mLblAggregateTitle.Top = bottomTop - 40
    mLblAggregateHeader.Left = mLblAggregateTitle.Left
    mLblAggregateHeader.Top = bottomTop - 18
    mLblAggregateHeader.Width = rightW
    mLstAggregate.Left = mLblAggregateTitle.Left
    mLstAggregate.Top = bottomTop
    mLstAggregate.Width = rightW
    mLstAggregate.Height = bottomH

    buttonTop = layoutH - 72
    statusTop = layoutH - 38
    mBtnConfirm.Top = buttonTop
    mBtnUndo.Top = buttonTop
    mBtnRedo.Top = buttonTop
    mBtnClear.Top = buttonTop
    mBtnClose.Left = layoutW - mBtnClose.Width - margin
    mBtnClose.Top = buttonTop

    mTxtStatus.Top = statusTop
    mTxtStatus.Width = layoutW - (margin * 2)

Done:
    mResizing = False
End Sub

Private Function ResolveOperatorWorkbook() As Workbook
    If Not mOperatorWorkbook Is Nothing Then
        If Not mOperatorWorkbook.IsAddin Then
            Set ResolveOperatorWorkbook = mOperatorWorkbook
            Exit Function
        End If
    End If
    If Not Application.ActiveWorkbook Is Nothing Then
        If Not Application.ActiveWorkbook.IsAddin Then
            Set ResolveOperatorWorkbook = Application.ActiveWorkbook
        End If
    End If
End Function

Private Sub ActivateOperatorWorkbook()
    On Error Resume Next
    If Not mOperatorWorkbook Is Nothing Then mOperatorWorkbook.Activate
    On Error GoTo 0
End Sub

Private Sub RefreshAllViews()
    LoadInventoryCache
    RefreshInventory
    RefreshStaging
    modTS_Received.EnforceReceivingSupportSheetsHidden ResolveOperatorWorkbook()
End Sub

Private Sub LoadInventoryCache()
    On Error GoTo ErrHandler
    ActivateOperatorWorkbook
    mInventoryRows = modTS_Received.LoadReceivingFormInventory("")
    Exit Sub
ErrHandler:
    Erase mInventoryRows
    ShowStatus "Inventory cache load failed: " & Err.Description
End Sub

Private Sub RefreshInventory()
    On Error GoTo ErrHandler
    FillInventoryList FilterInventoryRows(CStr(mTxtSearch.Value))
    Exit Sub
ErrHandler:
    ShowStatus "Inventory refresh failed: " & Err.Description
End Sub

Private Sub RefreshStaging()
    On Error GoTo ErrHandler
    ActivateOperatorWorkbook
    FillListBox mLstStaged, modTS_Received.LoadReceivingFormTable("ReceivedTally"), 4
    FillListBox mLstAggregate, modTS_Received.LoadReceivingFormTable("AggregateReceived"), 10
    Exit Sub
ErrHandler:
    ShowStatus "Receiving staging refresh failed: " & Err.Description
End Sub

Private Sub FillListBox(ByVal target As MSForms.ListBox, ByVal values As Variant, ByVal maxCols As Long)
    Dim r As Long
    Dim c As Long
    Dim rowIndex As Long
    Dim rows As Long
    Dim cols As Long

    target.Clear
    If IsEmpty(values) Then Exit Sub
    If Not IsArray(values) Then Exit Sub

    rows = UBound(values, 1)
    cols = UBound(values, 2)
    If rows < 1 Or cols < 1 Then Exit Sub
    If cols > maxCols Then cols = maxCols
    target.ColumnCount = maxCols

    For r = 1 To rows
        target.AddItem NzText(values(r, 1))
        rowIndex = target.ListCount - 1
        For c = 2 To cols
            target.List(rowIndex, c - 1) = NzText(values(r, c))
        Next c
    Next r
End Sub

Private Function FilterInventoryRows(ByVal filterText As String) As Variant
    Dim tokens() As String
    Dim sourceRows As Variant
    Dim result() As Variant
    Dim trimmed() As Variant
    Dim r As Long
    Dim c As Long
    Dim outRow As Long
    Dim haystack As String
    Dim token As Variant
    Dim matched As Boolean

    If IsEmpty(mInventoryRows) Or Not IsArray(mInventoryRows) Then Exit Function
    sourceRows = mInventoryRows
    filterText = NormalizeSearchText(filterText)
    tokens = Split(filterText, " ")

    ReDim result(1 To UBound(sourceRows, 1), 1 To UBound(sourceRows, 2))
    For r = 1 To UBound(sourceRows, 1)
        haystack = ""
        For c = 1 To UBound(sourceRows, 2)
            haystack = haystack & " " & NormalizeSearchText(NzText(sourceRows(r, c)))
        Next c

        matched = True
        If filterText <> "" Then
            For Each token In tokens
                If Trim$(CStr(token)) <> "" Then
                    If InStr(1, haystack, CStr(token), vbTextCompare) = 0 Then
                        matched = False
                        Exit For
                    End If
                End If
            Next token
        End If

        If matched Then
            outRow = outRow + 1
            For c = 1 To UBound(sourceRows, 2)
                result(outRow, c) = sourceRows(r, c)
            Next c
        End If
    Next r

    If outRow = 0 Then Exit Function
    ReDim trimmed(1 To outRow, 1 To UBound(sourceRows, 2))
    For r = 1 To outRow
        For c = 1 To UBound(sourceRows, 2)
            trimmed(r, c) = result(r, c)
        Next c
    Next r
    FilterInventoryRows = trimmed
End Function

Private Sub FillInventoryList(ByVal values As Variant)
    FillListBox mLstInventory, values, 9
    If IsEmpty(values) Or Not IsArray(values) Then
        If Trim$(CStr(mTxtSearch.Value)) <> "" Then ShowStatus "No inventory rows match: " & CStr(mTxtSearch.Value)
    Else
        ShowStatus "Inventory rows shown: " & CStr(UBound(values, 1))
    End If
End Sub

Private Sub mTxtSearch_Change()
    If mLoading Then Exit Sub
    RefreshInventory
End Sub

Private Sub mBtnRefresh_Click()
    RefreshClicked
End Sub

Private Sub mBtnAdd_Click()
    AddSelectedInventory
End Sub

Private Sub mBtnConfirm_Click()
    On Error GoTo ErrHandler
    ActivateOperatorWorkbook
    modTS_Received.ConfirmWrites
    If modTS_Received.LastConfirmWritesSucceeded() Then
        modTS_Received.ClearReceivingFormStaging
        mTxtRef.Value = ""
        mTxtReceiptId.Value = DefaultReceiptId()
        mLstStaged.Clear
        mLstAggregate.Clear
        RefreshAllViews
        mLstStaged.Clear
        mLstAggregate.Clear
        ShowStatus "Confirm Writes finished; staged receiving rows cleared."
    Else
        RefreshAllViews
        ShowStatus "Confirm Writes did not complete; staged receiving rows were kept. " & modTS_Received.LastConfirmWritesStatus()
    End If
    Exit Sub
ErrHandler:
    ShowStatus "Confirm Writes failed: " & Err.Description
End Sub

Private Sub mBtnUndo_Click()
    On Error GoTo ErrHandler
    ActivateOperatorWorkbook
    modTS_Received.MacroUndo
    RefreshAllViews
    ShowStatus "Undo finished."
    Exit Sub
ErrHandler:
    ShowStatus "Undo failed: " & Err.Description
End Sub

Private Sub mBtnRedo_Click()
    On Error GoTo ErrHandler
    ActivateOperatorWorkbook
    modTS_Received.MacroRedo
    RefreshAllViews
    ShowStatus "Redo finished."
    Exit Sub
ErrHandler:
    ShowStatus "Redo failed: " & Err.Description
End Sub

Private Sub mBtnClear_Click()
    On Error GoTo ErrHandler
    ActivateOperatorWorkbook
    modTS_Received.ClearReceivingFormStaging
    RefreshStaging
    ShowStatus "Receiving form staging cleared."
    Exit Sub
ErrHandler:
    ShowStatus "Clear failed: " & Err.Description
End Sub

Private Sub mBtnClose_Click()
    Unload Me
End Sub

Private Sub RefreshClicked()
    On Error GoTo ErrHandler
    ActivateOperatorWorkbook
    modTS_Received.RefreshReceivingUiForWorkbook ResolveOperatorWorkbook(), "LOCAL"
    RefreshAllViews
    ShowStatus "Receiving inventory and staging refreshed."
    Exit Sub
ErrHandler:
    ShowStatus "Refresh failed: " & Err.Description
End Sub

Private Sub AddSelectedInventory()
    On Error GoTo ErrHandler

    Dim idx As Long
    Dim rowVal As Long
    Dim qtyVal As Double
    Dim refVal As String
    Dim report As String
    Dim itemCode As String

    idx = mLstInventory.ListIndex
    If idx < 0 Then
        ShowStatus "Select an inventory item first."
        Exit Sub
    End If

    refVal = Trim$(CStr(mTxtRef.Value))
    If refVal = "" Then
        ShowStatus "Ref number is required."
        Exit Sub
    End If

    rowVal = CLng(Val(NzText(mLstInventory.List(idx, 8))))
    If rowVal <= 0 Then rowVal = CLng(Val(NzText(mLstInventory.List(idx, 0))))
    itemCode = NzText(mLstInventory.List(idx, 1))
    qtyVal = CDbl(Val(CStr(mTxtQty.Value)))
    If qtyVal <= 0 Then
        ShowStatus "Quantity must be greater than zero."
        Exit Sub
    End If

    ActivateOperatorWorkbook
    If modTS_Received.StageReceivingFormItemForWorkbook(ResolveOperatorWorkbook(), refVal, rowVal, itemCode, qtyVal, report) Then
        RefreshStaging
        ShowStatus report
    Else
        ShowStatus report
    End If
    Exit Sub

ErrHandler:
    ShowStatus "Add failed: " & Err.Description
End Sub

Private Sub ShowStatus(ByVal messageText As String)
    If mTxtStatus Is Nothing Then Exit Sub
    mTxtStatus.Text = messageText
End Sub

Private Function DefaultReceiptId() As String
    DefaultReceiptId = "RCV-" & Format$(Now, "yyyymmdd-hhnnss")
End Function

Private Function NormalizeSearchText(ByVal valueIn As String) As String
    valueIn = LCase$(Trim$(valueIn))
    Do While InStr(1, valueIn, "  ", vbBinaryCompare) > 0
        valueIn = Replace$(valueIn, "  ", " ")
    Loop
    NormalizeSearchText = valueIn
End Function

Private Function NzText(ByVal valueIn As Variant) As String
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Then
        NzText = ""
    Else
        NzText = Trim$(CStr(valueIn))
    End If
End Function

Private Function MaxDoubleReceiving(ByVal a As Double, ByVal b As Double) As Double
    If a >= b Then
        MaxDoubleReceiving = a
    Else
        MaxDoubleReceiving = b
    End If
End Function
