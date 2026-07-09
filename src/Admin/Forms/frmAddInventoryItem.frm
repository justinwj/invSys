VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddInventoryItem
   Caption         =   "invSys Admin - Add Inventory Item"
   ClientHeight    =   6100
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   6900
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAddInventoryItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@RuntimeStubUserFormCode
Option Explicit

Private WithEvents mBtnAddField As MSForms.CommandButton
Attribute mBtnAddField.VB_VarHelpID = -1
Private WithEvents mBtnRemoveField As MSForms.CommandButton
Attribute mBtnRemoveField.VB_VarHelpID = -1
Private WithEvents mBtnOK As MSForms.CommandButton
Attribute mBtnOK.VB_VarHelpID = -1
Private WithEvents mBtnCancel As MSForms.CommandButton
Attribute mBtnCancel.VB_VarHelpID = -1
Private WithEvents mBtnAddMode As MSForms.CommandButton
Attribute mBtnAddMode.VB_VarHelpID = -1
Private WithEvents mBtnEditMode As MSForms.CommandButton
Attribute mBtnEditMode.VB_VarHelpID = -1
Private WithEvents mCmbEditItem As MSForms.ComboBox
Attribute mCmbEditItem.VB_VarHelpID = -1
Private WithEvents mCmbUom As MSForms.ComboBox
Attribute mCmbUom.VB_VarHelpID = -1
Private WithEvents mTxtQty As MSForms.ComboBox
Attribute mTxtQty.VB_VarHelpID = -1
Private WithEvents mTxtImagePath As MSForms.TextBox
Attribute mTxtImagePath.VB_VarHelpID = -1

Private mLblTitle As MSForms.Label
Private mLblContext As MSForms.Label
Private mLblEditItem As MSForms.Label
Private mLblItemName As MSForms.Label
Private mLblUom As MSForms.Label
Private mLblQty As MSForms.Label
Private mLblLocation As MSForms.Label
Private mLblDescription As MSForms.Label
Private mLblCategory As MSForms.Label
Private mLblVendorName As MSForms.Label
Private mLblVendorCode As MSForms.Label
Private mLblExternalCode As MSForms.Label
Private mLblImagePath As MSForms.Label
Private mLblCustomName As MSForms.Label
Private mLblCustomValue As MSForms.Label
Private mLblGenerated As MSForms.Label
Private mLblStatus As MSForms.Label
Private mTxtItemName As MSForms.TextBox
Private mTxtLocation As MSForms.TextBox
Private mTxtDescription As MSForms.TextBox
Private mTxtCategory As MSForms.TextBox
Private mTxtVendorName As MSForms.TextBox
Private mTxtVendorCode As MSForms.TextBox
Private mTxtExternalCode As MSForms.TextBox
Private mTxtCustomName As MSForms.TextBox
Private mTxtCustomValue As MSForms.TextBox
Private mLstCustomFields As MSForms.ListBox

Private mAccepted As Boolean
Private mGeneratedSku As String
Private mGeneratedRow As Long
Private mWarehouseId As String
Private mStationId As String
Private mUserId As String
Private mAnchors As Object
Private mResizeInitialized As Boolean
Private mCatalogItems As Object
Private mEditMode As Boolean
Private mSelectedEditSku As String
Private mLoading As Boolean
Private mPreviousUom As String
Private mInitStep As String
Private mAllowUomPrompt As Boolean
Private mImagePlaceholderActive As Boolean

Private Const ANCHOR_LEFT As Long = 1
Private Const ANCHOR_TOP As Long = 2
Private Const ANCHOR_RIGHT As Long = 4
Private Const ANCHOR_BOTTOM As Long = 8
Private Const ADD_UOM_OPTION As String = "+ Add UOM..."
Private Const IMAGE_PATH_PLACEHOLDER As String = "Paste picture file path(s) or URL(s); separate multiple pictures with ;"
Private Const QTY_OPTION_UTILITY As String = "Utility"
Private Const QTY_OPTION_SERVICE As String = "Service"
Private Const QTY_OPTION_NOT_COUNTED As String = "Not counted"

Public Property Get Accepted() As Boolean
    Accepted = mAccepted
End Property

Public Property Get GeneratedSku() As String
    If mEditMode And mSelectedEditSku <> "" Then
        GeneratedSku = mSelectedEditSku
    Else
        GeneratedSku = mGeneratedSku
    End If
End Property

Public Property Get GeneratedRow() As Long
    GeneratedRow = mGeneratedRow
End Property

Public Property Get EditMode() As Boolean
    EditMode = mEditMode
End Property

Public Property Get ItemName() As String
    ItemName = Trim$(CStr(mTxtItemName.Value))
End Property

Public Property Get Uom() As String
    Uom = Trim$(CStr(mCmbUom.Value))
End Property

Public Property Get StartingQty() As Double
    If NonCountedItem Then Exit Property
    StartingQty = CDbl(Val(CStr(mTxtQty.Value)))
End Property

Public Property Get NonCountedItem() As Boolean
    NonCountedItem = QuantityModeIsNonCounted(QuantityModeText())
End Property

Public Property Get LocationValue() As String
    LocationValue = Trim$(CStr(mTxtLocation.Value))
End Property

Public Property Get DescriptionValue() As String
    DescriptionValue = Trim$(CStr(mTxtDescription.Value))
End Property

Public Property Get Category() As String
    Category = Trim$(CStr(mTxtCategory.Value))
End Property

Public Property Get VendorName() As String
    VendorName = Trim$(CStr(mTxtVendorName.Value))
End Property

Public Property Get VendorCode() As String
    VendorCode = Trim$(CStr(mTxtVendorCode.Value))
End Property

Public Property Get ExternalCode() As String
    ExternalCode = Trim$(CStr(mTxtExternalCode.Value))
End Property

Public Property Get ImagePath() As String
    If mImagePlaceholderActive Then Exit Property
    ImagePath = Trim$(CStr(mTxtImagePath.Value))
End Property

Public Property Get CustomFields() As Object
    Dim result As Object
    Dim i As Long

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    If NonCountedItem Then
        result("TRACK_QTY") = "FALSE"
        result("ITEM_KIND") = NonCountedItemKind()
    Else
        result("TRACK_QTY") = "TRUE"
        result("ITEM_KIND") = "INVENTORY"
    End If
    If Not mLstCustomFields Is Nothing Then
        For i = 0 To mLstCustomFields.ListCount - 1
            If Trim$(CStr(mLstCustomFields.List(i, 0))) <> "" Then
                result(Trim$(CStr(mLstCustomFields.List(i, 0)))) = Trim$(CStr(mLstCustomFields.List(i, 1)))
            End If
        Next i
    End If
    Set CustomFields = result
End Property

Public Sub TestSetQuantityMode(ByVal quantityMode As String)
    If mTxtQty Is Nothing Then EnsureControls
    mTxtQty.Value = quantityMode
    ApplyQuantityModeState
End Sub

Public Sub Configure(ByVal warehouseId As Variant, _
                     ByVal stationId As Variant, _
                     ByVal userId As Variant, _
                     ByVal generatedSku As Variant, _
                     ByVal generatedRow As Variant, _
                     ByVal defaultLocation As Variant)
    On Error GoTo FailConfigure

    mLoading = True
    mAllowUomPrompt = False
    mInitStep = "EnsureControls"
    EnsureControls
    mInitStep = "Set scalar context"
    mWarehouseId = SafeFormText(warehouseId)
    mStationId = SafeFormText(stationId)
    mUserId = SafeFormText(userId)
    mGeneratedSku = SafeFormText(generatedSku)
    mGeneratedRow = CLng(Val(SafeFormText(generatedRow)))
    Set mCatalogItems = New Collection
    mAccepted = False
    mEditMode = False
    mSelectedEditSku = ""

    mInitStep = "Load UOM options"
    LoadUomOptions
    mInitStep = "Load edit item options"
    LoadEditItemOptions
    mInitStep = "Reset field values"
    mTxtItemName.Value = ""
    mCmbUom.Value = "EA"
    mPreviousUom = "EA"
    mTxtQty.Value = "1"
    mTxtLocation.Value = SafeFormText(defaultLocation)
    mTxtDescription.Value = ""
    mTxtCategory.Value = ""
    mTxtVendorName.Value = ""
    mTxtVendorCode.Value = ""
    mTxtExternalCode.Value = ""
    ApplyQuantityModeState
    ShowImagePathPlaceholder
    mTxtCustomName.Value = ""
    mTxtCustomValue.Value = ""
    mLstCustomFields.Clear
    mInitStep = "Refresh generated label"
    RefreshGeneratedLabel
    mInitStep = "Apply mode layout"
    ApplyModeLayout
    mLblStatus.Caption = ""
    mLoading = False
    mInitStep = ""
    Exit Sub

FailConfigure:
    mLoading = False
    Err.Raise Err.Number, "frmAddInventoryItem.Configure", _
              "Add Inventory Item form failed during " & mInitStep & ": " & Err.Description
End Sub

Private Function SafeFormText(ByVal valueIn As Variant) As String
    If IsError(valueIn) Then Exit Function
    If IsNull(valueIn) Then Exit Function
    If IsEmpty(valueIn) Then Exit Function
    SafeFormText = Trim$(CStr(valueIn))
End Function

Public Sub AddCatalogItem(ByVal sku As String, _
                          ByVal rowValue As String, _
                          ByVal itemName As String, _
                          ByVal uomValue As String, _
                          ByVal locationValue As String, _
                          ByVal descriptionValue As String, _
                          ByVal vendorName As String, _
                          ByVal vendorCode As String, _
                          ByVal categoryValue As String, _
                          ByVal externalCodeValue As String, _
                          ByVal imagePathValue As String, _
                          Optional ByVal trackQtyValue As String = "", _
                          Optional ByVal itemKindValue As String = "")
    Dim item As Object
    Dim selectedUom As String

    EnsureControls
    If mCatalogItems Is Nothing Then Set mCatalogItems = New Collection
    sku = Trim$(sku)
    If sku = "" Then Exit Sub
    selectedUom = Trim$(CStr(mCmbUom.Value))

    Set item = CreateObject("Scripting.Dictionary")
    item.CompareMode = vbTextCompare
    item("SKU") = sku
    item("ROW") = Trim$(rowValue)
    item("ITEM") = Trim$(itemName)
    item("UOM") = UCase$(Trim$(uomValue))
    item("LOCATION") = Trim$(locationValue)
    item("DESCRIPTION") = Trim$(descriptionValue)
    item("VENDOR(s)") = Trim$(vendorName)
    item("VENDOR_CODE") = Trim$(vendorCode)
    item("CATEGORY") = Trim$(categoryValue)
    item("EXTERNAL_CODE") = Trim$(externalCodeValue)
    item("IMAGE_PATH") = Trim$(imagePathValue)
    item("TRACK_QTY") = UCase$(Trim$(trackQtyValue))
    item("ITEM_KIND") = UCase$(Trim$(itemKindValue))
    mCatalogItems.Add item

    mLoading = True
    LoadUomOptions
    If selectedUom <> "" Then
        mCmbUom.Value = selectedUom
    ElseIf Not mEditMode Then
        mCmbUom.Value = "EA"
    End If
    LoadEditItemOptions
    mLoading = False
End Sub

Private Sub UserForm_Initialize()
    mLoading = True
End Sub

Private Sub UserForm_Activate()
    mAllowUomPrompt = True
    If mAnchors Is Nothing Then InitializeAddInventoryAnchors
    If Not mResizeInitialized Then
        modUserFormResizeWin.EnableResizableUserForm Me
        mResizeInitialized = True
    End If
    If Not mAnchors Is Nothing Then mAnchors.ResizeControls
End Sub

Private Sub UserForm_Layout()
    If mAnchors Is Nothing Then Exit Sub
    mAnchors.ResizeControls
End Sub

Private Sub UserForm_Terminate()
    Set mAnchors = Nothing
End Sub

Private Sub EnsureControls()
    If Not mBtnOK Is Nothing Then Exit Sub

    Me.Caption = "invSys Admin - Add Inventory Item"
    Me.Width = 575
    Me.Height = 635

    Set mBtnAddMode = AddButton("btnAddMode", 14, 10, 118, 24, "Add Item")
    Set mBtnEditMode = AddButton("btnEditMode", 138, 10, 118, 24, "Edit Item")

    Set mLblTitle = AddLabel("lblTitle", 14, 44, 530, 20, "Add inventory item")
    mLblTitle.Font.Bold = True
    Set mLblContext = AddLabel("lblContext", 14, 66, 530, 18, "Fill the required fields. Internal item code is generated by invSys.")
    Set mLblGenerated = AddLabel("lblGenerated", 14, 90, 530, 32, "")
    mLblGenerated.WordWrap = True
    Set mLblEditItem = AddLabel("lblEditItem", 14, 128, 126, 18, "Inventory item")
    Set mCmbEditItem = AddCombo("cmbEditItem", 146, 124, 392, 22)
    mCmbEditItem.ColumnCount = 2
    mCmbEditItem.ColumnWidths = "360 pt;0 pt"
    mCmbEditItem.Style = fmStyleDropDownList

    Set mLblItemName = AddLabel("lblItemName", 14, 158, 126, 18, "Item name *")
    Set mTxtItemName = AddTextBox("txtItemName", 146, 154, 392, 22)
    Set mLblUom = AddLabel("lblUom", 14, 190, 126, 18, "UOM *")
    Set mCmbUom = AddCombo("cmbUom", 146, 186, 120, 22)
    LoadUomOptions
    Set mLblQty = AddLabel("lblQty", 288, 190, 92, 18, "Starting qty *")
    Set mTxtQty = AddCombo("txtQty", 386, 186, 152, 22)
    LoadQuantityOptions

    Set mLblLocation = AddLabel("lblLocation", 14, 222, 126, 18, "Default location")
    Set mTxtLocation = AddTextBox("txtLocation", 146, 218, 120, 22)
    Set mLblCategory = AddLabel("lblCategory", 288, 222, 92, 18, "Category")
    Set mTxtCategory = AddTextBox("txtCategory", 386, 218, 152, 22)

    Set mLblDescription = AddLabel("lblDescription", 14, 254, 126, 18, "Description")
    Set mTxtDescription = AddTextBox("txtDescription", 146, 250, 392, 22)
    Set mLblVendorName = AddLabel("lblVendorName", 14, 286, 126, 18, "Vendor(s)")
    Set mTxtVendorName = AddTextBox("txtVendorName", 146, 282, 392, 22)
    Set mLblVendorCode = AddLabel("lblVendorCode", 14, 318, 126, 18, "Vendor code")
    Set mTxtVendorCode = AddTextBox("txtVendorCode", 146, 314, 120, 22)
    Set mLblExternalCode = AddLabel("lblExternalCode", 288, 318, 92, 18, "External code")
    Set mTxtExternalCode = AddTextBox("txtExternalCode", 386, 314, 152, 22)
    Set mLblImagePath = AddLabel("lblImagePath", 14, 350, 126, 18, "Picture path/URL")
    Set mTxtImagePath = AddTextBox("txtImagePath", 146, 346, 392, 22)
    ShowImagePathPlaceholder

    Set mLblCustomName = AddLabel("lblCustomName", 14, 388, 126, 18, "Additional field")
    Set mTxtCustomName = AddTextBox("txtCustomName", 146, 384, 144, 22)
    Set mLblCustomValue = AddLabel("lblCustomValue", 298, 388, 42, 18, "Value")
    Set mTxtCustomValue = AddTextBox("txtCustomValue", 342, 384, 132, 22)
    Set mBtnAddField = AddButton("btnAddField", 484, 383, 54, 24, "Add")

    Set mLstCustomFields = AddListBox("lstCustomFields", 146, 414, 328, 96)
    mLstCustomFields.ColumnCount = 2
    mLstCustomFields.ColumnWidths = "130 pt;190 pt"
    Set mBtnRemoveField = AddButton("btnRemoveField", 484, 414, 54, 24, "Remove")

    Set mLblStatus = AddLabel("lblStatus", 146, 518, 328, 28, "")
    mLblStatus.ForeColor = 255
    Set mBtnOK = AddButton("btnOK", 374, 562, 78, 28, "Add Item")
    Set mBtnCancel = AddButton("btnCancel", 460, 562, 78, 28, "Cancel")

    ApplyModeLayout
End Sub

Private Sub InitializeAddInventoryAnchors()
    Set mAnchors = modDynamicForms.CreateFormAnchorManager()
    mAnchors.Initialize Me, 575, 635

    mAnchors.Add mBtnAddMode, ANCHOR_LEFT Or ANCHOR_TOP
    mAnchors.Add mBtnEditMode, ANCHOR_LEFT Or ANCHOR_TOP
    mAnchors.Add mLblContext, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mLblGenerated, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mCmbEditItem, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mTxtItemName, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mTxtDescription, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mTxtVendorName, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mTxtImagePath, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mTxtCustomValue, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mBtnAddField, ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mLstCustomFields, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add mBtnRemoveField, ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mLblStatus, ANCHOR_LEFT Or ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add mBtnOK, ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add mBtnCancel, ANCHOR_RIGHT Or ANCHOR_BOTTOM
End Sub

Private Sub LoadUomOptions()
    Dim seen As Object
    Dim item As Variant
    Dim uomValue As String

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    mCmbUom.Clear
    AddUomOption seen, "EA"
    AddUomOption seen, "CS"
    AddUomOption seen, "LB"
    AddUomOption seen, "OZ"
    AddUomOption seen, "GAL"
    AddUomOption seen, "L"
    AddUomOption seen, "ML"
    AddUomOption seen, "KG"
    If Not mCatalogItems Is Nothing Then
        For Each item In mCatalogItems
            uomValue = CatalogField(item, "UOM")
            If uomValue <> "" Then AddUomOption seen, uomValue
        Next item
    End If
    mCmbUom.AddItem ADD_UOM_OPTION
End Sub

Private Sub AddUomOption(ByVal seen As Object, ByVal uomValue As String)
    uomValue = UCase$(Trim$(uomValue))
    If uomValue = "" Then Exit Sub
    If seen.Exists(uomValue) Then Exit Sub
    seen(uomValue) = True
    mCmbUom.AddItem uomValue
End Sub

Private Sub LoadEditItemOptions()
    Dim item As Variant
    Dim displayText As String
    Dim rowIndex As Long

    If mCmbEditItem Is Nothing Then Exit Sub
    mCmbEditItem.Clear
    If mCatalogItems Is Nothing Then Exit Sub
    For Each item In mCatalogItems
        If CatalogField(item, "SKU") <> "" Then
            displayText = CatalogField(item, "ITEM")
            If displayText = "" Then displayText = CatalogField(item, "SKU")
            If CatalogField(item, "UOM") <> "" Then displayText = displayText & " [" & CatalogField(item, "UOM") & "]"
            mCmbEditItem.AddItem displayText
            rowIndex = mCmbEditItem.ListCount - 1
            mCmbEditItem.List(rowIndex, 1) = CatalogField(item, "SKU")
        End If
    Next item
End Sub

Private Sub LoadQuantityOptions()
    If mTxtQty Is Nothing Then Exit Sub
    mTxtQty.Clear
    mTxtQty.AddItem "1"
    mTxtQty.AddItem QTY_OPTION_UTILITY
    mTxtQty.AddItem QTY_OPTION_SERVICE
    mTxtQty.AddItem QTY_OPTION_NOT_COUNTED
End Sub

Private Sub RefreshGeneratedLabel()
    If mEditMode Then
        If mSelectedEditSku <> "" Then
            mLblGenerated.Caption = "Internal code: " & mSelectedEditSku & "    ROW: " & CStr(mGeneratedRow) & _
                                    "    Warehouse: " & mWarehouseId & " / " & mStationId
        Else
            mLblGenerated.Caption = "Choose an inventory item to edit. Warehouse: " & mWarehouseId & " / " & mStationId
        End If
    Else
        mLblGenerated.Caption = "Internal code: " & mGeneratedSku & "    ROW: " & CStr(mGeneratedRow) & _
                                "    Warehouse: " & mWarehouseId & " / " & mStationId
    End If
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

Private Function AddListBox(ByVal controlName As String, _
                            ByVal leftPos As Single, _
                            ByVal topPos As Single, _
                            ByVal widthVal As Single, _
                            ByVal heightVal As Single) As MSForms.ListBox
    Set AddListBox = Me.Controls.Add("Forms.ListBox.1", controlName, True)
    AddListBox.Left = leftPos
    AddListBox.Top = topPos
    AddListBox.Width = widthVal
    AddListBox.Height = heightVal
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

Private Sub mBtnAddField_Click()
    Dim fieldName As String
    Dim fieldValue As String
    Dim rowIndex As Long

    fieldName = NormalizeCustomFieldName(Trim$(CStr(mTxtCustomName.Value)))
    fieldValue = Trim$(CStr(mTxtCustomValue.Value))
    If fieldName = "" Then
        mLblStatus.Caption = "Additional field name is required."
        Exit Sub
    End If
    If IsReservedCustomField(fieldName) Then
        mLblStatus.Caption = "That field is already handled by the form."
        Exit Sub
    End If
    If fieldValue = "" Then
        mLblStatus.Caption = "Additional field value is required."
        Exit Sub
    End If

    rowIndex = FindCustomFieldRow(fieldName)
    If rowIndex < 0 Then
        mLstCustomFields.AddItem fieldName
        rowIndex = mLstCustomFields.ListCount - 1
    End If
    mLstCustomFields.List(rowIndex, 1) = fieldValue
    mTxtCustomName.Value = ""
    mTxtCustomValue.Value = ""
    mLblStatus.Caption = ""
End Sub

Private Sub mBtnRemoveField_Click()
    If mLstCustomFields.ListIndex >= 0 Then mLstCustomFields.RemoveItem mLstCustomFields.ListIndex
End Sub

Private Sub mBtnOK_Click()
    If Not ValidateForm Then Exit Sub
    mAccepted = True
    Me.Hide
End Sub

Private Sub mBtnAddMode_Click()
    If mLoading Then Exit Sub
    mEditMode = False
    ApplyModeLayout
End Sub

Private Sub mBtnEditMode_Click()
    If mLoading Then Exit Sub
    mEditMode = True
    ApplyModeLayout
End Sub

Private Sub mCmbEditItem_Change()
    If mLoading Then Exit Sub
    If Not mEditMode Then Exit Sub
    LoadSelectedEditItem
End Sub

Private Sub mCmbUom_Change()
    Dim newUom As String

    If mLoading Then Exit Sub
    If Trim$(CStr(mCmbUom.Value)) <> ADD_UOM_OPTION Then
        If Trim$(CStr(mCmbUom.Value)) <> "" Then mPreviousUom = Trim$(CStr(mCmbUom.Value))
        Exit Sub
    End If
    If Not mAllowUomPrompt Then
        If mPreviousUom <> "" Then
            mCmbUom.Value = mPreviousUom
        Else
            mCmbUom.Value = "EA"
        End If
        Exit Sub
    End If

    newUom = UCase$(Trim$(InputBox("New UOM:", "invSys Admin - Add UOM")))
    If newUom = "" Then
        If mPreviousUom <> "" Then
            mCmbUom.Value = mPreviousUom
        Else
            mCmbUom.Value = "EA"
        End If
        Exit Sub
    End If
    InsertUomBeforeAddOption newUom
    mCmbUom.Value = newUom
    mPreviousUom = newUom
End Sub

Private Sub mTxtQty_Change()
    If mLoading Then Exit Sub
    ApplyQuantityModeState
End Sub

Private Sub mTxtImagePath_Enter()
    If mImagePlaceholderActive Then
        mImagePlaceholderActive = False
        mTxtImagePath.Value = ""
        mTxtImagePath.ForeColor = vbWindowText
    End If
End Sub

Private Sub mTxtImagePath_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Trim$(CStr(mTxtImagePath.Value)) = "" Then ShowImagePathPlaceholder
End Sub

Private Sub ShowImagePathPlaceholder()
    If mTxtImagePath Is Nothing Then Exit Sub
    mImagePlaceholderActive = True
    mTxtImagePath.Value = IMAGE_PATH_PLACEHOLDER
    mTxtImagePath.ForeColor = RGB(128, 128, 128)
End Sub

Private Sub SetImagePathValue(ByVal imagePathValue As String)
    imagePathValue = Trim$(imagePathValue)
    If imagePathValue = "" Then
        ShowImagePathPlaceholder
    Else
        mImagePlaceholderActive = False
        mTxtImagePath.ForeColor = vbWindowText
        mTxtImagePath.Value = imagePathValue
    End If
End Sub

Private Function QuantityModeText() As String
    If mTxtQty Is Nothing Then Exit Function
    QuantityModeText = Trim$(CStr(mTxtQty.Value))
End Function

Private Function QuantityModeIsNonCounted(ByVal valueText As String) As Boolean
    Select Case UCase$(Trim$(valueText))
        Case "UTILITY", "SERVICE", "NOT COUNTED", "NON-COUNTED", "NONCOUNTED", "UNTRACKED"
            QuantityModeIsNonCounted = True
    End Select
End Function

Private Function NonCountedItemKind() As String
    Select Case UCase$(QuantityModeText())
        Case "UTILITY"
            NonCountedItemKind = "UTILITY"
        Case "SERVICE"
            NonCountedItemKind = "SERVICE"
        Case Else
            NonCountedItemKind = "NON_COUNTED"
    End Select
End Function

Private Sub ApplyQuantityModeState()
    If mTxtQty Is Nothing Or mLblQty Is Nothing Then Exit Sub
    mTxtQty.Enabled = Not mEditMode
    If NonCountedItem Then
        mLblQty.Caption = "Qty mode"
        If Not mTxtCategory Is Nothing Then
            If Trim$(CStr(mTxtCategory.Value)) = "" Then mTxtCategory.Value = NonCountedItemKind()
        End If
        If UCase$(QuantityModeText()) = "UTILITY" Then
            If Not mTxtVendorName Is Nothing Then
                If Trim$(CStr(mTxtVendorName.Value)) = "" Then mTxtVendorName.Value = "Utility"
            End If
        End If
    ElseIf mEditMode Then
        mLblQty.Caption = "Qty"
    Else
        mLblQty.Caption = "Starting qty *"
    End If
End Sub

Private Sub InsertUomBeforeAddOption(ByVal uomValue As String)
    Dim i As Long

    uomValue = UCase$(Trim$(uomValue))
    If uomValue = "" Then Exit Sub
    For i = 0 To mCmbUom.ListCount - 1
        If StrComp(CStr(mCmbUom.List(i)), uomValue, vbTextCompare) = 0 Then Exit Sub
    Next i
    If mCmbUom.ListCount > 0 Then
        mCmbUom.AddItem uomValue, mCmbUom.ListCount - 1
    Else
        mCmbUom.AddItem uomValue
    End If
End Sub

Private Sub ApplyModeLayout()
    If mBtnOK Is Nothing Then Exit Sub
    mLblEditItem.Visible = mEditMode
    mCmbEditItem.Visible = mEditMode
    If mEditMode Then
        mLblTitle.Caption = "Edit inventory item"
        mLblContext.Caption = "Edit catalog fields for an existing inventory item."
        mLblQty.Caption = "Qty"
        mBtnOK.Caption = "Save Item"
    Else
        mLblTitle.Caption = "Add inventory item"
        mLblContext.Caption = "Fill the required fields. Internal item code is generated by invSys."
        mLblQty.Caption = "Starting qty *"
        mBtnOK.Caption = "Add Item"
    End If
    If Not mEditMode Then
        mSelectedEditSku = ""
        If mGeneratedRow < 0 Then mGeneratedRow = 0
    ElseIf mCmbEditItem.ListIndex < 0 Then
        ClearEditableFields
    End If
    ApplyQuantityModeState
    RefreshGeneratedLabel
    mLblStatus.Caption = ""
End Sub

Private Sub ClearEditableFields()
    mTxtItemName.Value = ""
    mCmbUom.Value = "EA"
    mPreviousUom = "EA"
    mTxtQty.Value = ""
    mTxtLocation.Value = ""
    mTxtDescription.Value = ""
    mTxtCategory.Value = ""
    mTxtVendorName.Value = ""
    mTxtVendorCode.Value = ""
    mTxtExternalCode.Value = ""
    ApplyQuantityModeState
    ShowImagePathPlaceholder
    mTxtCustomName.Value = ""
    mTxtCustomValue.Value = ""
    mLstCustomFields.Clear
End Sub

Private Sub LoadSelectedEditItem()
    Dim sku As String
    Dim item As Object

    If mCmbEditItem.ListIndex < 0 Then Exit Sub
    sku = CStr(mCmbEditItem.List(mCmbEditItem.ListIndex, 1))
    Set item = FindCatalogItemBySku(sku)
    If item Is Nothing Then Exit Sub

    mSelectedEditSku = sku
    mGeneratedRow = CLng(Val(CatalogField(item, "ROW")))
    mTxtItemName.Value = CatalogField(item, "ITEM")
    If CatalogField(item, "UOM") <> "" Then
        InsertUomBeforeAddOption CatalogField(item, "UOM")
        mCmbUom.Value = CatalogField(item, "UOM")
        mPreviousUom = CatalogField(item, "UOM")
    Else
        mCmbUom.Value = "EA"
        mPreviousUom = "EA"
    End If
    mTxtQty.Value = ""
    mTxtLocation.Value = CatalogField(item, "LOCATION")
    mTxtDescription.Value = CatalogField(item, "DESCRIPTION")
    mTxtCategory.Value = CatalogField(item, "CATEGORY")
    mTxtVendorName.Value = CatalogField(item, "VENDOR(s)")
    mTxtVendorCode.Value = CatalogField(item, "VENDOR_CODE")
    mTxtExternalCode.Value = CatalogField(item, "EXTERNAL_CODE")
    If CatalogItemIsNonCounted(item) Then
        Select Case UCase$(CatalogField(item, "ITEM_KIND"))
            Case "UTILITY"
                mTxtQty.Value = QTY_OPTION_UTILITY
            Case "SERVICE"
                mTxtQty.Value = QTY_OPTION_SERVICE
            Case Else
                mTxtQty.Value = QTY_OPTION_NOT_COUNTED
        End Select
    End If
    ApplyQuantityModeState
    SetImagePathValue CatalogField(item, "IMAGE_PATH")
    mLstCustomFields.Clear
    RefreshGeneratedLabel
    mLblStatus.Caption = ""
End Sub

Private Sub mBtnCancel_Click()
    mAccepted = False
    Me.Hide
End Sub

Private Function ValidateForm() As Boolean
    If mGeneratedSku = "" Then
        mLblStatus.Caption = "Generated internal code is missing."
        Exit Function
    End If
    If mEditMode And mSelectedEditSku = "" Then
        mLblStatus.Caption = "Choose an inventory item to edit."
        Exit Function
    End If
    If (Not mEditMode) And mGeneratedRow <= 0 Then
        mLblStatus.Caption = "Generated ROW id is missing."
        Exit Function
    End If
    If ItemName = "" Then
        mLblStatus.Caption = "Item name is required."
        Exit Function
    End If
    If Uom = "" Then
        mLblStatus.Caption = "UOM is required."
        Exit Function
    End If
    If Not mEditMode Then
        If NonCountedItem Then
            ValidateForm = True
            Exit Function
        End If
        If Not IsNumeric(CStr(mTxtQty.Value)) Then
            mLblStatus.Caption = "Starting quantity must be numeric or a mode like Utility."
            Exit Function
        End If
        If StartingQty <= 0 Then
            mLblStatus.Caption = "Starting quantity must be greater than zero."
            Exit Function
        End If
    End If
    ValidateForm = True
End Function

Private Function CatalogItemIsNonCounted(ByVal item As Variant) As Boolean
    Dim trackQty As String
    Dim itemKind As String

    trackQty = UCase$(CatalogField(item, "TRACK_QTY"))
    itemKind = UCase$(CatalogField(item, "ITEM_KIND"))
    CatalogItemIsNonCounted = (trackQty = "FALSE" Or trackQty = "NO" Or trackQty = "0" _
                               Or itemKind = "UTILITY" Or itemKind = "SERVICE" Or itemKind = "NON_COUNTED")
End Function

Private Function FindCatalogItemBySku(ByVal sku As String) As Object
    Dim item As Variant

    If mCatalogItems Is Nothing Then Exit Function
    For Each item In mCatalogItems
        If StrComp(CatalogField(item, "SKU"), sku, vbTextCompare) = 0 Then
            Set FindCatalogItemBySku = item
            Exit Function
        End If
    Next item
End Function

Private Function CatalogField(ByVal item As Variant, ByVal fieldName As String) As String
    On Error Resume Next
    CatalogField = Trim$(CStr(item(fieldName)))
    On Error GoTo 0
End Function

Private Function FindCustomFieldRow(ByVal fieldName As String) As Long
    Dim i As Long

    FindCustomFieldRow = -1
    For i = 0 To mLstCustomFields.ListCount - 1
        If StrComp(CStr(mLstCustomFields.List(i, 0)), fieldName, vbTextCompare) = 0 Then
            FindCustomFieldRow = i
            Exit Function
        End If
    Next i
End Function

Private Function NormalizeCustomFieldName(ByVal fieldName As String) As String
    fieldName = Replace(fieldName, vbCr, " ")
    fieldName = Replace(fieldName, vbLf, " ")
    fieldName = Replace(fieldName, vbTab, " ")
    fieldName = Replace(fieldName, "[", "(")
    fieldName = Replace(fieldName, "]", ")")
    Do While InStr(1, fieldName, "  ", vbBinaryCompare) > 0
        fieldName = Replace(fieldName, "  ", " ")
    Loop
    fieldName = Trim$(fieldName)
    If Len(fieldName) > 48 Then fieldName = Left$(fieldName, 48)
    NormalizeCustomFieldName = fieldName
End Function

Private Function IsReservedCustomField(ByVal fieldName As String) As Boolean
    Select Case UCase$(Trim$(fieldName))
        Case "SKU", "ROW", "ITEM_CODE", "ITEM", "UOM", "LOCATION", "QTY", "TOTAL INV", _
             "QTYAVAILABLE", "DESCRIPTION", "VENDOR(S)", "VENDOR_CODE", "CATEGORY", _
             "EXTERNAL_CODE", "IMAGE_PATH", "TRACK_QTY", "ITEM_KIND", "NOTE", "IOTYPE"
            IsReservedCustomField = True
    End Select
End Function
