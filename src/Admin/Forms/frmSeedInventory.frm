VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSeedInventory
   Caption         =   "invSys Admin - Seed Inventory"
   ClientHeight    =   2350
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   6200
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSeedInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@RuntimeStubUserFormCode
Option Explicit

Private mResizeInitialized As Boolean

Private WithEvents mCmbWarehouse As MSForms.ComboBox
Attribute mCmbWarehouse.VB_VarHelpID = -1
Private WithEvents mCboDemoDataSet As MSForms.ComboBox
Attribute mCboDemoDataSet.VB_VarHelpID = -1
Private WithEvents mBtnSeed As MSForms.CommandButton
Attribute mBtnSeed.VB_VarHelpID = -1
Private WithEvents mBtnDelete As MSForms.CommandButton
Attribute mBtnDelete.VB_VarHelpID = -1
Private WithEvents mBtnUpload As MSForms.CommandButton
Attribute mBtnUpload.VB_VarHelpID = -1
Private WithEvents mBtnRepairInboxes As MSForms.CommandButton
Attribute mBtnRepairInboxes.VB_VarHelpID = -1
Private WithEvents mBtnCancel As MSForms.CommandButton
Attribute mBtnCancel.VB_VarHelpID = -1
Private mLblTitle As MSForms.Label
Private mLblWarehouse As MSForms.Label
Private mLblStation As MSForms.Label
Private mLblUser As MSForms.Label
Private mLblRoot As MSForms.Label
Private mLblRootValue As MSForms.Label
Private mLblDemoDataSet As MSForms.Label
Private mLblStatus As MSForms.Label
Private mTxtStation As MSForms.TextBox
Private mTxtUser As MSForms.TextBox

Private mAccepted As Boolean
Private mSelectedWarehouseId As String
Private mSelectedStationId As String
Private mSelectedRuntimeRoot As String
Private mSelectedUserId As String
Private mSelectedAction As String
Private mSelectedUploadPath As String

Public Property Get Accepted() As Boolean
    Accepted = mAccepted
End Property

Public Property Get SelectedWarehouseId() As String
    SelectedWarehouseId = mSelectedWarehouseId
End Property

Public Property Get SelectedStationId() As String
    SelectedStationId = mSelectedStationId
End Property

Public Property Get SelectedRuntimeRoot() As String
    SelectedRuntimeRoot = mSelectedRuntimeRoot
End Property

Public Property Get SelectedUserId() As String
    SelectedUserId = mSelectedUserId
End Property

Public Property Get SelectedAction() As String
    SelectedAction = mSelectedAction
End Property

Public Property Get SelectedUploadPath() As String
    SelectedUploadPath = mSelectedUploadPath
End Property

Public Sub Configure(ByVal warehouseOptions As Collection, _
                     ByVal defaultWarehouseId As String, _
                     ByVal defaultStationId As String, _
                     ByVal defaultUserId As String)
    Dim item As Variant
    Dim rowIndex As Long
    Dim matchIndex As Long

    EnsureControls
    mAccepted = False
    mSelectedAction = ""
    mSelectedUploadPath = ""
    ConfigureDemoDataSetChoices
    mCmbWarehouse.Clear
    matchIndex = -1

    If Not warehouseOptions Is Nothing Then
        For Each item In warehouseOptions
            mCmbWarehouse.AddItem CStr(item(0))
            rowIndex = mCmbWarehouse.ListCount - 1
            mCmbWarehouse.List(rowIndex, 1) = CStr(item(1))
            mCmbWarehouse.List(rowIndex, 2) = CStr(item(2))
            mCmbWarehouse.List(rowIndex, 3) = CStr(item(3))
            mCmbWarehouse.List(rowIndex, 4) = CStr(item(4))
            If matchIndex < 0 _
               And StrComp(CStr(item(1)), defaultWarehouseId, vbTextCompare) = 0 _
               And (defaultStationId = "" Or StrComp(CStr(item(2)), defaultStationId, vbTextCompare) = 0) Then
                matchIndex = rowIndex
            End If
        Next item
    End If

    If mCmbWarehouse.ListCount > 0 Then
        If matchIndex < 0 Then matchIndex = 0
        mCmbWarehouse.ListIndex = matchIndex
    End If

    mTxtStation.Value = modStationIdentity.CurrentComputerStationId()
    mTxtUser.Value = defaultUserId
    ApplyWarehouseSelection
End Sub

Private Sub UserForm_Initialize()
    EnsureControls
End Sub

Private Sub UserForm_Activate()
    If mResizeInitialized Then Exit Sub
    On Error Resume Next
    modUserFormResizeWin.EnableResizableUserForm Me, True, True
    On Error GoTo 0
    mResizeInitialized = True
End Sub

Private Sub EnsureControls()
    If Not mCmbWarehouse Is Nothing Then Exit Sub

    Me.Caption = "invSys Admin - Demo Inventory"
    Me.Width = 620
    Me.Height = 326

    Set mLblTitle = AddLabel("lblTitle", 12, 12, 576, 22, "Choose a demo inventory action and warehouse.")
    Set mLblWarehouse = AddLabel("lblWarehouse", 12, 48, 92, 18, "Warehouse")
    Set mCmbWarehouse = AddCombo("cmbWarehouse", 108, 44, 468, 24)
    mCmbWarehouse.ColumnCount = 5
    mCmbWarehouse.ColumnWidths = "340 pt;0 pt;0 pt;0 pt;0 pt"
    mCmbWarehouse.MatchRequired = True
    mCmbWarehouse.Style = fmStyleDropDownList

    Set mLblStation = AddLabel("lblStation", 12, 82, 92, 18, "Station")
    Set mTxtStation = AddTextBox("txtStation", 108, 78, 90, 22)
    mTxtStation.Locked = True
    mTxtStation.BackColor = &HEFEFEF
    Set mLblUser = AddLabel("lblUser", 240, 82, 84, 18, "Admin user")
    Set mTxtUser = AddTextBox("txtUser", 324, 78, 252, 22)

    Set mLblRoot = AddLabel("lblRoot", 12, 116, 92, 18, "Runtime root")
    Set mLblRootValue = AddLabel("lblRootValue", 108, 116, 468, 36, "")
    mLblRootValue.WordWrap = True

    Set mLblDemoDataSet = AddLabel("lblDemoDataSet", 12, 158, 92, 18, "Data set")
    Set mCboDemoDataSet = AddCombo("cboDemoDataSet", 108, 154, 468, 24)
    mCboDemoDataSet.Style = fmStyleDropDownList
    ConfigureDemoDataSetChoices

    Set mLblStatus = AddLabel("lblStatus", 108, 186, 468, 24, "")
    mLblStatus.ForeColor = 255

    Set mBtnSeed = AddButton("btnSeedDemoInventory", 18, 218, 172, 32, "Seed Demo Inventory")
    Set mBtnDelete = AddButton("btnDeleteDemoInventory", 204, 218, 172, 32, "Delete Demo Inventory")
    Set mBtnUpload = AddButton("btnUploadDemoInventory", 390, 218, 186, 32, "Upload Demo Inventory")
    Set mBtnRepairInboxes = AddButton("btnRepairInboxes", 342, 262, 116, 28, "Repair Inboxes")
    Set mBtnCancel = AddButton("btnCancel", 470, 262, 106, 28, "Cancel")
End Sub

Private Sub ConfigureDemoDataSetChoices()
    If mCboDemoDataSet Is Nothing Then Exit Sub
    mCboDemoDataSet.Clear
    mCboDemoDataSet.AddItem "R1 Workflow Kit (built-in)"
    mCboDemoDataSet.AddItem "Uploaded CSV (choose with Upload Demo Inventory)"
    mCboDemoDataSet.ListIndex = 0
End Sub

Private Function AddLabel(ByVal controlName As String, _
                          ByVal leftPos As Single, _
                          ByVal topPos As Single, _
                          ByVal widthVal As Single, _
                          ByVal heightVal As Single, _
                          ByVal captionText As String) As MSForms.Label
    Set AddLabel = Me.Controls.Add("Forms.Label.1", controlName, True)
    AddLabel.Left = leftPos
    AddLabel.Top = topPos
    AddLabel.Width = widthVal
    AddLabel.Height = heightVal
    AddLabel.Caption = captionText
End Function

Private Function AddCombo(ByVal controlName As String, _
                          ByVal leftPos As Single, _
                          ByVal topPos As Single, _
                          ByVal widthVal As Single, _
                          ByVal heightVal As Single) As MSForms.ComboBox
    Set AddCombo = Me.Controls.Add("Forms.ComboBox.1", controlName, True)
    AddCombo.Left = leftPos
    AddCombo.Top = topPos
    AddCombo.Width = widthVal
    AddCombo.Height = heightVal
End Function

Private Function AddTextBox(ByVal controlName As String, _
                            ByVal leftPos As Single, _
                            ByVal topPos As Single, _
                            ByVal widthVal As Single, _
                            ByVal heightVal As Single) As MSForms.TextBox
    Set AddTextBox = Me.Controls.Add("Forms.TextBox.1", controlName, True)
    AddTextBox.Left = leftPos
    AddTextBox.Top = topPos
    AddTextBox.Width = widthVal
    AddTextBox.Height = heightVal
End Function

Private Function AddButton(ByVal controlName As String, _
                           ByVal leftPos As Single, _
                           ByVal topPos As Single, _
                           ByVal widthVal As Single, _
                           ByVal heightVal As Single, _
                           ByVal captionText As String) As MSForms.CommandButton
    Set AddButton = Me.Controls.Add("Forms.CommandButton.1", controlName, True)
    AddButton.Left = leftPos
    AddButton.Top = topPos
    AddButton.Width = widthVal
    AddButton.Height = heightVal
    AddButton.Caption = captionText
End Function

Private Sub mCmbWarehouse_Change()
    ApplyWarehouseSelection
End Sub

Private Sub AcceptDemoInventoryAction(ByVal actionName As String)
    If mCmbWarehouse.ListIndex < 0 Then
        mLblStatus.Caption = "Choose a warehouse."
        Exit Sub
    End If
    If Trim$(CStr(mTxtStation.Value)) = "" Then
        mLblStatus.Caption = "Station is required."
        Exit Sub
    End If
    If Trim$(CStr(mTxtUser.Value)) = "" Then
        mLblStatus.Caption = "Admin user is required."
        Exit Sub
    End If

    mSelectedWarehouseId = CStr(mCmbWarehouse.List(mCmbWarehouse.ListIndex, 1))
    mSelectedStationId = Trim$(CStr(mTxtStation.Value))
    mSelectedRuntimeRoot = CStr(mCmbWarehouse.List(mCmbWarehouse.ListIndex, 3))
    mSelectedUserId = Trim$(CStr(mTxtUser.Value))
    mSelectedAction = UCase$(Trim$(actionName))
    mAccepted = True
    Me.Hide
End Sub

Private Sub mBtnSeed_Click()
    If mCboDemoDataSet.ListIndex = 1 And Trim$(mSelectedUploadPath) = "" Then
        mLblStatus.ForeColor = 255
        mLblStatus.Caption = "Use Upload Demo Inventory to choose a CSV data set first."
        Exit Sub
    End If
    If mCboDemoDataSet.ListIndex = 0 Then mSelectedUploadPath = ""
    AcceptDemoInventoryAction modAdminInventorySeed.DEMO_ACTION_SEED
End Sub

Private Sub mBtnDelete_Click()
    AcceptDemoInventoryAction modAdminInventorySeed.DEMO_ACTION_DELETE
End Sub

Private Sub mBtnUpload_Click()
    Dim selectedPath As Variant
    Dim displayName As String

    selectedPath = Application.GetOpenFilename( _
        FileFilter:="CSV files (*.csv),*.csv", _
        Title:="Select Demo Inventory Data Set")
    If VarType(selectedPath) = vbBoolean Then Exit Sub
    mSelectedUploadPath = Trim$(CStr(selectedPath))
    displayName = Dir$(mSelectedUploadPath)
    If displayName = "" Then displayName = mSelectedUploadPath
    mCboDemoDataSet.List(1) = "Uploaded CSV: " & displayName
    mCboDemoDataSet.ListIndex = 1
    mLblStatus.ForeColor = &H8000
    mLblStatus.Caption = "Selected data set. Click Seed Demo Inventory to apply it."
End Sub

Private Sub mBtnCancel_Click()
    mAccepted = False
    Me.Hide
End Sub

Public Function TestDemoInventoryActionContract() As String
    EnsureControls
    If mBtnSeed.Caption = "Seed Demo Inventory" _
       And mBtnDelete.Caption = "Delete Demo Inventory" _
       And mBtnUpload.Caption = "Upload Demo Inventory" _
       And mCboDemoDataSet.ListCount = 2 _
       And mCboDemoDataSet.List(0) = "R1 Workflow Kit (built-in)" Then
        TestDemoInventoryActionContract = "OK|Seed=True|Delete=True|Upload=True|Dataset=True"
    Else
        TestDemoInventoryActionContract = "FAIL|Demo inventory actions are incomplete."
    End If
End Function

Private Sub mBtnRepairInboxes_Click()
    Dim warehouseId As String
    Dim stationId As String
    Dim runtimeRoot As String
    Dim report As String

    If mCmbWarehouse.ListIndex < 0 Then
        mLblStatus.ForeColor = 255
        mLblStatus.Caption = "Choose a warehouse."
        Exit Sub
    End If

    warehouseId = CStr(mCmbWarehouse.List(mCmbWarehouse.ListIndex, 1))
    stationId = Trim$(CStr(mTxtStation.Value))
    runtimeRoot = CStr(mCmbWarehouse.List(mCmbWarehouse.ListIndex, 3))
    If stationId = "" Then stationId = modStationIdentity.CurrentComputerStationId()

    If RepairStationInboxes(warehouseId, stationId, runtimeRoot, report) Then
        mCmbWarehouse.List(mCmbWarehouse.ListIndex, 0) = warehouseId & " | " & stationId & " | " & runtimeRoot
        mCmbWarehouse.List(mCmbWarehouse.ListIndex, 2) = stationId
        mCmbWarehouse.List(mCmbWarehouse.ListIndex, 4) = "Ready"
        mLblStatus.ForeColor = 32768
        mLblStatus.Caption = "Inboxes repaired."
    Else
        mLblStatus.ForeColor = 255
        mLblStatus.Caption = report
    End If
End Sub

Private Sub ApplyWarehouseSelection()
    If mCmbWarehouse.ListIndex < 0 Then
        mLblRootValue.Caption = ""
        Exit Sub
    End If

    mTxtStation.Value = modStationIdentity.CurrentComputerStationId()
    mLblRootValue.Caption = CStr(mCmbWarehouse.List(mCmbWarehouse.ListIndex, 3))
    If Trim$(CStr(mCmbWarehouse.List(mCmbWarehouse.ListIndex, 4))) = "" _
       Or StrComp(CStr(mCmbWarehouse.List(mCmbWarehouse.ListIndex, 4)), "Ready", vbTextCompare) = 0 Then
        mLblStatus.ForeColor = 32768
        mLblStatus.Caption = ""
    Else
        mLblStatus.ForeColor = 255
        mLblStatus.Caption = CStr(mCmbWarehouse.List(mCmbWarehouse.ListIndex, 4))
    End If
End Sub

Private Function RepairStationInboxes(ByVal warehouseId As String, _
                                      ByVal stationId As String, _
                                      ByVal runtimeRoot As String, _
                                      ByRef report As String) As Boolean
    Dim inboxPath As String
    Dim stepReport As String

    warehouseId = Trim$(warehouseId)
    stationId = Trim$(stationId)
    runtimeRoot = Trim$(runtimeRoot)
    If warehouseId = "" Then
        report = "WarehouseId is required."
        Exit Function
    End If
    If stationId = "" Then stationId = modStationIdentity.CurrentComputerStationId()
    If runtimeRoot <> "" Then modRuntimeWorkbooks.SetCoreDataRootOverride runtimeRoot

    If Not modConfig.EnsureStationInbox(warehouseId, stationId, "RECEIVE", "", inboxPath, stepReport) Then
        report = "Receiving inbox repair failed: " & stepReport
        Exit Function
    End If

    inboxPath = ""
    stepReport = ""
    If Not modConfig.EnsureStationInbox(warehouseId, stationId, "SHIP", "", inboxPath, stepReport) Then
        report = "Shipping inbox repair failed: " & stepReport
        Exit Function
    End If

    inboxPath = ""
    stepReport = ""
    If Not modConfig.EnsureStationInbox(warehouseId, stationId, "PRODUCTION", "", inboxPath, stepReport) Then
        report = "Production inbox repair failed: " & stepReport
        Exit Function
    End If

    report = "OK"
    RepairStationInboxes = True
End Function
