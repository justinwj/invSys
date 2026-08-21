VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInventoryViewer
   Caption         =   "Inventory Viewer"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10320
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmInventoryViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@RuntimeStubUserFormCode
Option Explicit

Private Const SETTINGS_APP As String = "invSys"
Private Const SETTINGS_SECTION_OPERATIONS As String = "Operations"
Private Const SETTINGS_EVENT_RANGE As String = "InventoryViewerEventRange"

Private WithEvents mTxtSearch As MSForms.TextBox
Private WithEvents mBtnRefresh As MSForms.CommandButton
Private WithEvents mBtnClose As MSForms.CommandButton
Private WithEvents mTabs As MSForms.TabStrip
Private mCboEventRange As MSForms.ComboBox
Private mLstInventory As MSForms.ListBox
Private mLblTitle As MSForms.Label
Private mLblHeaders As MSForms.Label
Private mLblStatus As MSForms.Label
Private mLblEventRange As MSForms.Label
Private mLblEventRangeHelp As MSForms.Label
Private mLayout As cOperationsAnchorManager
Private mWarehouseId As String
Private mRows As Variant
Private mBuilt As Boolean
Private mResizeInitialized As Boolean
Private mGeneration As Long
Private mColumnCount As Long
Private mLoadStatus As String

Private Sub UserForm_Initialize()
    BuildLayout
End Sub

Private Sub UserForm_Activate()
    If Not mResizeInitialized Then
        modUserFormResizeWin.EnableResizableUserForm Me, True, True
        mResizeInitialized = True
    End If
    If Not mLayout Is Nothing Then mLayout.ApplyAnchoredLayout
End Sub

Private Sub UserForm_Layout()
    If Not mLayout Is Nothing Then mLayout.ApplyAnchoredLayout
End Sub

Private Sub UserForm_Terminate()
    modInventoryViewer.UnregisterInventoryViewer Me
    Set mLayout = Nothing
End Sub

Public Sub SetWarehouse(ByVal warehouseId As String)
    mWarehouseId = Trim$(warehouseId)
    Me.Caption = "Inventory Viewer - " & mWarehouseId
End Sub

Public Sub SetGeneration(ByVal generation As Long)
    mGeneration = generation
End Sub

Public Sub RefreshInventory()
    If Not mBuilt Then BuildLayout
    mColumnCount = 6
    LoadViewerPayload modInventoryViewerData.LoadCurrentInventoryViewerData(), "inventory level(s)"
End Sub

Public Sub RefreshEvents()
    Dim publishedPayload As String
    If Not mBuilt Then BuildLayout
    mColumnCount = 10
    publishedPayload = modInventoryViewerData.LoadCurrentInventoryEventViewerData()
    LoadViewerPayload modInventoryViewer.LoadInventoryViewerEvents(publishedPayload), "event(s)"
End Sub

Private Sub LoadViewerPayload(ByVal payload As String, ByVal rowLabel As String)
    Dim lines As Variant
    Dim header As Variant
    Dim dataRows() As Variant
    Dim fields As Variant
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim dataIndex As Long

    lines = Split(payload, vbCrLf)
    header = Split(CStr(lines(0)), vbTab)
    If UBound(header) < 1 Or StrComp(CStr(header(0)), "OK", vbTextCompare) <> 0 Then
        mRows = Empty
        mLstInventory.Clear
        If UBound(header) >= 1 Then
            mLoadStatus = ViewerUnescape(CStr(header(1)))
        Else
            mLoadStatus = "Inventory snapshot could not be loaded."
        End If
        mLblStatus.Caption = mLoadStatus
        Exit Sub
    End If

    If UBound(lines) >= 1 Then
        ReDim dataRows(1 To UBound(lines), 1 To mColumnCount)
        For rowIndex = 1 To UBound(lines)
            If Trim$(CStr(lines(rowIndex))) <> "" Then
                fields = Split(CStr(lines(rowIndex)), vbTab)
                If UBound(fields) >= mColumnCount - 1 Then
                    dataIndex = dataIndex + 1
                    For columnIndex = 1 To mColumnCount
                        dataRows(dataIndex, columnIndex) = ViewerUnescape(CStr(fields(columnIndex - 1)))
                    Next columnIndex
                End If
            End If
        Next rowIndex
    End If
    If dataIndex = 0 Then
        mRows = Empty
    ElseIf dataIndex = UBound(dataRows, 1) Then
        mRows = dataRows
    Else
        mRows = TrimViewerRows(dataRows, dataIndex, mColumnCount)
    End If
    mLoadStatus = CStr(dataIndex) & " " & rowLabel & ". Published data read at " & CStr(header(2)) & "."
    mLblStatus.Caption = mLoadStatus
    RenderRows Trim$(CStr(mTxtSearch.Value))
End Sub

Public Function TestReport() As String
    TestReport = "OK|Warehouse=" & mWarehouseId & _
        "|VisibleRows=" & CStr(mLstInventory.ListCount) & _
        "|Generation=" & CStr(mGeneration) & _
        "|Modeless=True|Status=" & mLblStatus.Caption
End Function

Public Function TestApplySearch(ByVal filterText As String) As String
    mTxtSearch.Value = filterText
    RenderRows filterText
    TestApplySearch = "OK|Filter=" & filterText & _
        "|VisibleRows=" & CStr(mLstInventory.ListCount) & _
        "|Generation=" & CStr(mGeneration)
End Function

Public Function TestEventsReport(Optional ByVal rangeText As String = "") As String
    Dim removeRows As Long
    Dim readableDates As Long
    Dim rowIndex As Long
    Dim firstDate As String
    Dim firstReference As String
    If Not mBuilt Then BuildLayout
    mTabs.Value = 1
    If Trim$(rangeText) <> "" Then mCboEventRange.Value = rangeText
    mBtnRefresh_Click
    For rowIndex = 0 To mLstInventory.ListCount - 1
        If StrComp(Trim$(CStr(mLstInventory.List(rowIndex, 1))), "Remove", vbTextCompare) = 0 Then
            removeRows = removeRows + 1
        End If
        If Trim$(CStr(mLstInventory.List(rowIndex, 0))) <> "" And _
           Not IsNumeric(Trim$(CStr(mLstInventory.List(rowIndex, 0)))) Then
            readableDates = readableDates + 1
        End If
    Next rowIndex
    If mLstInventory.ListCount > 0 Then
        firstDate = CStr(mLstInventory.List(0, 0))
        firstReference = CStr(mLstInventory.List(0, 2))
    End If
    TestEventsReport = "OK|Title=" & mLblTitle.Caption & _
        "|TabCount=" & CStr(mTabs.Tabs.Count) & _
        "|TabCaptions=" & mTabs.Tabs(0).Caption & "," & mTabs.Tabs(1).Caption & _
        "|SelectedTab=" & mTabs.Tabs(mTabs.Value).Caption & _
        "|VisibleRows=" & CStr(mLstInventory.ListCount) & _
        "|ReadableDates=" & CStr(readableDates) & _
        "|FirstDate=" & firstDate & _
        "|FirstReference=" & firstReference & _
        "|RemoveRows=" & CStr(removeRows) & _
        "|EventRange=" & CStr(mCboEventRange.Value) & _
        "|RangeControlVisible=" & CStr(mCboEventRange.Visible) & _
        "|Columns=" & CStr(mLstInventory.ColumnCount) & _
        "|ReadOnly=True|Generation=" & CStr(mGeneration)
End Function

Private Sub BuildLayout()
    Dim rememberedRange As String
    Dim numericRange As Double

    If mBuilt Then Exit Sub
    Me.Width = 860
    Me.Height = 535

    Set mTabs = Me.Controls.Add("Forms.TabStrip.1", "tabsInventoryViewer", True)
    With mTabs
        .Move 12, 8, 820, 24
        .Tabs(0).Caption = "Inventory"
        .Tabs(1).Caption = "Events"
        .Value = 0
    End With
    Set mLblTitle = AddLabel("lblTitle", "Current inventory levels", 12, 40, 360, 22, True)
    Set mBtnRefresh = AddButton("btnRefresh", "Refresh", 740, 38, 92, 28)
    AddLabel "lblSearch", "Search", 12, 78, 76, 18, True
    Set mTxtSearch = AddTextBox("txtSearch", 92, 74, 740, 24)
    Set mLblEventRange = AddLabel("lblEventRange", "Event range", 12, 110, 96, 18, True)
    Set mCboEventRange = Me.Controls.Add("Forms.ComboBox.1", "cboEventRange", True)
    On Error Resume Next
    rememberedRange = Trim$(GetSetting(SETTINGS_APP, SETTINGS_SECTION_OPERATIONS, SETTINGS_EVENT_RANGE, "All"))
    On Error GoTo 0
    Select Case UCase$(rememberedRange)
        Case "ALL"
            rememberedRange = "All"
        Case "DAY"
            rememberedRange = "Day"
        Case "WEEK"
            rememberedRange = "Week"
        Case "MONTH"
            rememberedRange = "Month"
        Case Else
            If IsNumeric(rememberedRange) Then
                On Error Resume Next
                Err.Clear
                numericRange = CDbl(rememberedRange)
                If Err.Number <> 0 Then numericRange = 0
                On Error GoTo 0
                If numericRange > 0 And numericRange = Fix(numericRange) And numericRange <= 36500 Then
                    rememberedRange = CStr(CLng(numericRange))
                Else
                    rememberedRange = "All"
                End If
            Else
                rememberedRange = "All"
            End If
    End Select
    With mCboEventRange
        .Move 112, 104, 150, 24
        .Style = fmStyleDropDownCombo
        .MatchRequired = False
        .AddItem "All"
        .AddItem "Day"
        .AddItem "Week"
        .AddItem "Month"
        .Value = rememberedRange
    End With
    Set mLblEventRangeHelp = AddLabel("lblEventRangeHelp", _
        "Choose Day, Week, Month, or type a whole number of days; select Refresh to apply.", _
        276, 110, 556, 18, False)
    Set mLblHeaders = AddLabel("lblHeaders", _
        "Item Code                         Item                                  UOM       Quantity       Location                  Condition", _
        12, 140, 820, 18, True)
    Set mLstInventory = AddListBox("lstInventory", 12, 162, 820, 292)
    With mLstInventory
        .ColumnCount = 6
        .ColumnWidths = "135 pt;190 pt;52 pt;72 pt;120 pt;74 pt"
        .IntegralHeight = False
    End With
    Set mLblStatus = AddLabel("lblStatus", "Select Refresh to load the current published snapshot.", 12, 470, 680, 32, False)
    Set mBtnClose = AddButton("btnClose", "Close", 740, 466, 92, 30)

    Set mLayout = modOperationsLayout.OperationsAnchorManager()
    mLayout.ConfigureForForm Me, 720, 430
    mLayout.RegisterControl mTabs, OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_TOP Or OPERATIONS_ANCHOR_RIGHT
    mLayout.RegisterControl mLblTitle, OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_TOP
    mLayout.RegisterControl mBtnRefresh, OPERATIONS_ANCHOR_TOP Or OPERATIONS_ANCHOR_RIGHT
    mLayout.RegisterControl mTxtSearch, OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_TOP Or OPERATIONS_ANCHOR_RIGHT
    mLayout.RegisterControl mLblEventRange, OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_TOP
    mLayout.RegisterControl mCboEventRange, OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_TOP
    mLayout.RegisterControl mLblEventRangeHelp, OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_TOP Or OPERATIONS_ANCHOR_RIGHT
    mLayout.RegisterControl mLblHeaders, OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_TOP Or OPERATIONS_ANCHOR_RIGHT
    mLayout.RegisterControl mLstInventory, OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_TOP Or OPERATIONS_ANCHOR_RIGHT Or OPERATIONS_ANCHOR_BOTTOM
    mLayout.RegisterControl mLblStatus, OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_RIGHT Or OPERATIONS_ANCHOR_BOTTOM
    mLayout.RegisterControl mBtnClose, OPERATIONS_ANCHOR_RIGHT Or OPERATIONS_ANCHOR_BOTTOM
    mBuilt = True
    ApplyViewerTab
End Sub

Private Sub mTabs_Change()
    ApplyViewerTab
End Sub

Private Sub ApplyViewerTab()
    If mTabs Is Nothing Then Exit Sub
    mTxtSearch.Value = vbNullString
    If mTabs.Value = 1 Then
        mLblEventRange.Visible = True
        mCboEventRange.Visible = True
        mLblEventRangeHelp.Visible = True
        mLblTitle.Caption = "Inventory and shipping events"
        mLblHeaders.Caption = "Date                 Event             Reference        Item                    Qty      UOM    Location       Condition    User          Details"
        mLstInventory.ColumnCount = 10
        mLstInventory.ColumnWidths = "105 pt;82 pt;92 pt;130 pt;52 pt;46 pt;82 pt;72 pt;72 pt;190 pt"
        RefreshEvents
    Else
        mLblEventRange.Visible = False
        mCboEventRange.Visible = False
        mLblEventRangeHelp.Visible = False
        mLblTitle.Caption = "Current inventory levels"
        mLblHeaders.Caption = "Item Code                         Item                                  UOM       Quantity       Location                  Condition"
        mLstInventory.ColumnCount = 6
        mLstInventory.ColumnWidths = "135 pt;190 pt;52 pt;72 pt;120 pt;74 pt"
        RefreshInventory
    End If
End Sub

Private Sub mTxtSearch_Change()
    RenderRows Trim$(CStr(mTxtSearch.Value))
End Sub

Private Sub mBtnRefresh_Click()
    If mTabs.Value = 1 Then
        RefreshEvents
    Else
        RefreshInventory
    End If
End Sub

Private Sub mBtnClose_Click()
    Unload Me
End Sub

Private Sub RenderRows(ByVal filterText As String)
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim matches As Boolean
    Dim rangeText As String
    Dim eventDays As Long
    Dim eventCutoff As Date
    Dim eventDateValue As Date
    Dim hasEventDateFilter As Boolean
    Dim numericRange As Double
    Dim storedRange As String

    mLstInventory.Clear
    mLblStatus.Caption = mLoadStatus
    If mTabs.Value = 1 Then
        rangeText = UCase$(Trim$(CStr(mCboEventRange.Value)))
        Select Case rangeText
            Case "", "ALL"
                storedRange = "All"
            Case "DAY"
                eventDays = 1
                storedRange = "Day"
            Case "WEEK"
                eventDays = 7
                storedRange = "Week"
            Case "MONTH"
                eventDays = 30
                storedRange = "Month"
            Case Else
                If IsNumeric(rangeText) Then
                    On Error Resume Next
                    Err.Clear
                    numericRange = CDbl(rangeText)
                    If Err.Number <> 0 Then numericRange = 0
                    On Error GoTo 0
                    If numericRange > 0 And numericRange = Fix(numericRange) And numericRange <= 36500 Then
                        eventDays = CLng(numericRange)
                        storedRange = CStr(eventDays)
                    End If
                End If
                If eventDays = 0 Then
                    mLblStatus.Caption = "Enter All, Day, Week, Month, or a whole number from 1 to 36500, then select Refresh."
                    Exit Sub
                End If
        End Select
        mCboEventRange.Value = storedRange
        On Error Resume Next
        SaveSetting SETTINGS_APP, SETTINGS_SECTION_OPERATIONS, SETTINGS_EVENT_RANGE, storedRange
        On Error GoTo 0
        If eventDays > 0 Then
            hasEventDateFilter = True
            eventCutoff = DateAdd("d", -eventDays, Now)
        End If
    End If
    If IsEmpty(mRows) Then Exit Sub
    filterText = LCase$(Trim$(filterText))
    For rowIndex = LBound(mRows, 1) To UBound(mRows, 1)
        matches = (filterText = "")
        If Not matches Then
            For columnIndex = 1 To mColumnCount
                If InStr(1, LCase$(CStr(mRows(rowIndex, columnIndex))), filterText, vbTextCompare) > 0 Then
                    matches = True
                    Exit For
                End If
            Next columnIndex
        End If
        If matches And hasEventDateFilter Then
            If IsDate(CStr(mRows(rowIndex, 1))) Then
                eventDateValue = CDate(CStr(mRows(rowIndex, 1)))
                matches = (eventDateValue >= eventCutoff And eventDateValue <= Now)
            Else
                matches = False
            End If
        End If
        If matches Then
            mLstInventory.AddItem CStr(mRows(rowIndex, 1))
            For columnIndex = 2 To mColumnCount
                mLstInventory.List(mLstInventory.ListCount - 1, columnIndex - 1) = CStr(mRows(rowIndex, columnIndex))
            Next columnIndex
        End If
    Next rowIndex
    If hasEventDateFilter Then
        mLblStatus.Caption = mLoadStatus & " Showing " & CStr(mLstInventory.ListCount) & _
            " event(s) in the rolling " & CStr(eventDays) & "-day window."
    End If
End Sub

Private Function TrimViewerRows(ByVal sourceRows As Variant, ByVal rowCount As Long, ByVal columnCount As Long) As Variant
    Dim resultRows() As Variant
    Dim rowIndex As Long
    Dim columnIndex As Long

    ReDim resultRows(1 To rowCount, 1 To columnCount)
    For rowIndex = 1 To rowCount
        For columnIndex = 1 To columnCount
            resultRows(rowIndex, columnIndex) = sourceRows(rowIndex, columnIndex)
        Next columnIndex
    Next rowIndex
    TrimViewerRows = resultRows
End Function

Private Function ViewerUnescape(ByVal valueIn As String) As String
    valueIn = Replace(valueIn, "\n", vbLf)
    valueIn = Replace(valueIn, "\r", vbCr)
    valueIn = Replace(valueIn, "\t", vbTab)
    valueIn = Replace(valueIn, "\\", "\")
    ViewerUnescape = valueIn
End Function

Private Function AddLabel(ByVal controlName As String, _
                          ByVal captionText As String, _
                          ByVal leftValue As Double, _
                          ByVal topValue As Double, _
                          ByVal widthValue As Double, _
                          ByVal heightValue As Double, _
                          ByVal boldValue As Boolean) As MSForms.Label
    Set AddLabel = Me.Controls.Add("Forms.Label.1", controlName, True)
    With AddLabel
        .Caption = captionText
        .Move leftValue, topValue, widthValue, heightValue
        .Font.Bold = boldValue
    End With
End Function

Private Function AddTextBox(ByVal controlName As String, _
                            ByVal leftValue As Double, _
                            ByVal topValue As Double, _
                            ByVal widthValue As Double, _
                            ByVal heightValue As Double) As MSForms.TextBox
    Set AddTextBox = Me.Controls.Add("Forms.TextBox.1", controlName, True)
    AddTextBox.Move leftValue, topValue, widthValue, heightValue
End Function

Private Function AddListBox(ByVal controlName As String, _
                            ByVal leftValue As Double, _
                            ByVal topValue As Double, _
                            ByVal widthValue As Double, _
                            ByVal heightValue As Double) As MSForms.ListBox
    Set AddListBox = Me.Controls.Add("Forms.ListBox.1", controlName, True)
    AddListBox.Move leftValue, topValue, widthValue, heightValue
End Function

Private Function AddButton(ByVal controlName As String, _
                           ByVal captionText As String, _
                           ByVal leftValue As Double, _
                           ByVal topValue As Double, _
                           ByVal widthValue As Double, _
                           ByVal heightValue As Double) As MSForms.CommandButton
    Set AddButton = Me.Controls.Add("Forms.CommandButton.1", controlName, True)
    With AddButton
        .Caption = captionText
        .Move leftValue, topValue, widthValue, heightValue
    End With
End Function
