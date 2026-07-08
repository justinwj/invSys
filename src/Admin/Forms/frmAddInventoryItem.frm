VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddInventoryItem
   Caption         =   "invSys Admin - Add Inventory Item"
   ClientHeight    =   5700
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

Private mLblTitle As MSForms.Label
Private mLblContext As MSForms.Label
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
Private mCmbUom As MSForms.ComboBox
Private mTxtQty As MSForms.TextBox
Private mTxtLocation As MSForms.TextBox
Private mTxtDescription As MSForms.TextBox
Private mTxtCategory As MSForms.TextBox
Private mTxtVendorName As MSForms.TextBox
Private mTxtVendorCode As MSForms.TextBox
Private mTxtExternalCode As MSForms.TextBox
Private mTxtImagePath As MSForms.TextBox
Private mTxtCustomName As MSForms.TextBox
Private mTxtCustomValue As MSForms.TextBox
Private mLstCustomFields As MSForms.ListBox

Private mAccepted As Boolean
Private mGeneratedSku As String
Private mGeneratedRow As Long
Private mWarehouseId As String
Private mStationId As String
Private mUserId As String

Public Property Get Accepted() As Boolean
    Accepted = mAccepted
End Property

Public Property Get GeneratedSku() As String
    GeneratedSku = mGeneratedSku
End Property

Public Property Get GeneratedRow() As Long
    GeneratedRow = mGeneratedRow
End Property

Public Property Get ItemName() As String
    ItemName = Trim$(CStr(mTxtItemName.Value))
End Property

Public Property Get Uom() As String
    Uom = Trim$(CStr(mCmbUom.Value))
End Property

Public Property Get StartingQty() As Double
    StartingQty = CDbl(Val(CStr(mTxtQty.Value)))
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
    ImagePath = Trim$(CStr(mTxtImagePath.Value))
End Property

Public Property Get CustomFields() As Object
    Dim result As Object
    Dim i As Long

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    If Not mLstCustomFields Is Nothing Then
        For i = 0 To mLstCustomFields.ListCount - 1
            If Trim$(CStr(mLstCustomFields.List(i, 0))) <> "" Then
                result(Trim$(CStr(mLstCustomFields.List(i, 0)))) = Trim$(CStr(mLstCustomFields.List(i, 1)))
            End If
        Next i
    End If
    Set CustomFields = result
End Property

Public Sub Configure(ByVal warehouseId As String, _
                     ByVal stationId As String, _
                     ByVal userId As String, _
                     ByVal generatedSku As String, _
                     ByVal generatedRow As Long, _
                     ByVal defaultLocation As String)
    EnsureControls
    mWarehouseId = Trim$(warehouseId)
    mStationId = Trim$(stationId)
    mUserId = Trim$(userId)
    mGeneratedSku = Trim$(generatedSku)
    mGeneratedRow = generatedRow
    mAccepted = False

    mTxtItemName.Value = ""
    mCmbUom.Value = "EA"
    mTxtQty.Value = "1"
    mTxtLocation.Value = defaultLocation
    mTxtDescription.Value = ""
    mTxtCategory.Value = ""
    mTxtVendorName.Value = ""
    mTxtVendorCode.Value = ""
    mTxtExternalCode.Value = ""
    mTxtImagePath.Value = ""
    mTxtCustomName.Value = ""
    mTxtCustomValue.Value = ""
    mLstCustomFields.Clear
    RefreshGeneratedLabel
    mLblStatus.Caption = ""
End Sub

Private Sub UserForm_Initialize()
    EnsureControls
End Sub

Private Sub EnsureControls()
    If Not mBtnOK Is Nothing Then Exit Sub

    Me.Caption = "invSys Admin - Add Inventory Item"
    Me.Width = 575
    Me.Height = 500

    Set mLblTitle = AddLabel("lblTitle", 14, 12, 530, 20, "Add inventory item")
    mLblTitle.Font.Bold = True
    Set mLblContext = AddLabel("lblContext", 14, 34, 530, 18, "Fill the required fields. Internal item code is generated by invSys.")
    Set mLblGenerated = AddLabel("lblGenerated", 14, 58, 530, 32, "")
    mLblGenerated.WordWrap = True

    Set mLblItemName = AddLabel("lblItemName", 14, 102, 126, 18, "Item name *")
    Set mTxtItemName = AddTextBox("txtItemName", 146, 98, 392, 22)
    Set mLblUom = AddLabel("lblUom", 14, 134, 126, 18, "UOM *")
    Set mCmbUom = AddCombo("cmbUom", 146, 130, 120, 22)
    LoadUomOptions
    Set mLblQty = AddLabel("lblQty", 288, 134, 92, 18, "Starting qty *")
    Set mTxtQty = AddTextBox("txtQty", 386, 130, 152, 22)

    Set mLblLocation = AddLabel("lblLocation", 14, 166, 126, 18, "Default location")
    Set mTxtLocation = AddTextBox("txtLocation", 146, 162, 120, 22)
    Set mLblCategory = AddLabel("lblCategory", 288, 166, 92, 18, "Category")
    Set mTxtCategory = AddTextBox("txtCategory", 386, 162, 152, 22)

    Set mLblDescription = AddLabel("lblDescription", 14, 198, 126, 18, "Description")
    Set mTxtDescription = AddTextBox("txtDescription", 146, 194, 392, 22)
    Set mLblVendorName = AddLabel("lblVendorName", 14, 230, 126, 18, "Vendor(s)")
    Set mTxtVendorName = AddTextBox("txtVendorName", 146, 226, 392, 22)
    Set mLblVendorCode = AddLabel("lblVendorCode", 14, 262, 126, 18, "Vendor code")
    Set mTxtVendorCode = AddTextBox("txtVendorCode", 146, 258, 120, 22)
    Set mLblExternalCode = AddLabel("lblExternalCode", 288, 262, 92, 18, "External code")
    Set mTxtExternalCode = AddTextBox("txtExternalCode", 386, 258, 152, 22)
    Set mLblImagePath = AddLabel("lblImagePath", 14, 294, 126, 18, "Picture path/URL")
    Set mTxtImagePath = AddTextBox("txtImagePath", 146, 290, 392, 22)

    Set mLblCustomName = AddLabel("lblCustomName", 14, 332, 126, 18, "Additional field")
    Set mTxtCustomName = AddTextBox("txtCustomName", 146, 328, 144, 22)
    Set mLblCustomValue = AddLabel("lblCustomValue", 298, 332, 42, 18, "Value")
    Set mTxtCustomValue = AddTextBox("txtCustomValue", 342, 328, 132, 22)
    Set mBtnAddField = AddButton("btnAddField", 484, 327, 54, 24, "Add")

    Set mLstCustomFields = AddListBox("lstCustomFields", 146, 358, 328, 58)
    mLstCustomFields.ColumnCount = 2
    mLstCustomFields.ColumnWidths = "130 pt;190 pt"
    Set mBtnRemoveField = AddButton("btnRemoveField", 484, 358, 54, 24, "Remove")

    Set mLblStatus = AddLabel("lblStatus", 146, 424, 328, 28, "")
    mLblStatus.ForeColor = 255
    Set mBtnOK = AddButton("btnOK", 374, 454, 78, 28, "Add Item")
    Set mBtnCancel = AddButton("btnCancel", 460, 454, 78, 28, "Cancel")
End Sub

Private Sub LoadUomOptions()
    mCmbUom.Clear
    mCmbUom.AddItem "EA"
    mCmbUom.AddItem "CS"
    mCmbUom.AddItem "LB"
    mCmbUom.AddItem "OZ"
    mCmbUom.AddItem "GAL"
    mCmbUom.AddItem "L"
    mCmbUom.AddItem "ML"
    mCmbUom.AddItem "KG"
End Sub

Private Sub RefreshGeneratedLabel()
    mLblGenerated.Caption = "Internal code: " & mGeneratedSku & "    ROW: " & CStr(mGeneratedRow) & _
                            "    Warehouse: " & mWarehouseId & " / " & mStationId
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

Private Sub mBtnCancel_Click()
    mAccepted = False
    Me.Hide
End Sub

Private Function ValidateForm() As Boolean
    If mGeneratedSku = "" Then
        mLblStatus.Caption = "Generated internal code is missing."
        Exit Function
    End If
    If mGeneratedRow <= 0 Then
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
    If Not IsNumeric(CStr(mTxtQty.Value)) Then
        mLblStatus.Caption = "Starting quantity must be numeric."
        Exit Function
    End If
    If StartingQty <= 0 Then
        mLblStatus.Caption = "Starting quantity must be greater than zero."
        Exit Function
    End If
    ValidateForm = True
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
             "EXTERNAL_CODE", "IMAGE_PATH", "NOTE", "IOTYPE"
            IsReservedCustomField = True
    End Select
End Function
