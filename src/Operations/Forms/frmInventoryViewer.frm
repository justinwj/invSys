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

Private WithEvents mTxtSearch As MSForms.TextBox
Private WithEvents mBtnRefresh As MSForms.CommandButton
Private WithEvents mBtnClose As MSForms.CommandButton
Private WithEvents mTabs As MSForms.TabStrip
Private mLstInventory As MSForms.ListBox
Private mLblTitle As MSForms.Label
Private mLblHeaders As MSForms.Label
Private mLblStatus As MSForms.Label
Private mLayout As cOperationsAnchorManager
Private mWarehouseId As String
Private mRows As Variant
Private mBuilt As Boolean
Private mResizeInitialized As Boolean
Private mGeneration As Long
Private mColumnCount As Long

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
            mLblStatus.Caption = ViewerUnescape(CStr(header(1)))
        Else
            mLblStatus.Caption = "Inventory snapshot could not be loaded."
        End If
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
    RenderRows Trim$(CStr(mTxtSearch.Value))
    mLblStatus.Caption = CStr(dataIndex) & " " & rowLabel & ". Published data read at " & CStr(header(2)) & "."
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

Public Function TestEventsReport() As String
    If Not mBuilt Then BuildLayout
    mTabs.Value = 1
    ApplyViewerTab
    TestEventsReport = "OK|Title=" & mLblTitle.Caption & _
        "|VisibleRows=" & CStr(mLstInventory.ListCount) & _
        "|Columns=" & CStr(mLstInventory.ColumnCount) & _
        "|ReadOnly=True|Generation=" & CStr(mGeneration)
End Function

Private Sub BuildLayout()
    If mBuilt Then Exit Sub
    Me.Width = 860
    Me.Height = 535

    Set mTabs = Me.Controls.Add("Forms.TabStrip.1", "tabsInventoryViewer", True)
    With mTabs
        .Move 12, 8, 820, 24
        .Tabs.Add "tabInventory", "Inventory"
        .Tabs.Add "tabEvents", "Events"
        .Value = 0
    End With
    Set mLblTitle = AddLabel("lblTitle", "Current inventory levels", 12, 40, 360, 22, True)
    Set mBtnRefresh = AddButton("btnRefresh", "Refresh", 740, 38, 92, 28)
    AddLabel "lblSearch", "Search", 12, 78, 76, 18, True
    Set mTxtSearch = AddTextBox("txtSearch", 92, 74, 740, 24)
    Set mLblHeaders = AddLabel("lblHeaders", _
        "Item Code                         Item                                  UOM       Quantity       Location                  Condition", _
        12, 110, 820, 18, True)
    Set mLstInventory = AddListBox("lstInventory", 12, 132, 820, 322)
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
        mLblTitle.Caption = "Inventory and shipping events"
        mLblHeaders.Caption = "Date                 Event             Reference        Item                    Qty      UOM    Location       Condition    User          Details"
        mLstInventory.ColumnCount = 10
        mLstInventory.ColumnWidths = "105 pt;82 pt;92 pt;130 pt;52 pt;46 pt;82 pt;72 pt;72 pt;190 pt"
        RefreshEvents
    Else
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

    mLstInventory.Clear
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
        If matches Then
            mLstInventory.AddItem CStr(mRows(rowIndex, 1))
            For columnIndex = 2 To mColumnCount
                mLstInventory.List(mLstInventory.ListCount - 1, columnIndex - 1) = CStr(mRows(rowIndex, columnIndex))
            Next columnIndex
        End If
    Next rowIndex
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
