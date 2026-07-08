VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProduction
   Caption         =   "Production"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13200
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProduction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@RuntimeStubUserFormCode
Option Explicit

Private WithEvents mPages As MSForms.MultiPage
Private WithEvents mLstBuilderRecipes As MSForms.ListBox
Private WithEvents mLstBuilderLines As MSForms.ListBox
Private WithEvents mCmbLineIo As MSForms.ComboBox
Private WithEvents mBtnBuilderRefresh As MSForms.CommandButton
Private WithEvents mBtnBuilderNew As MSForms.CommandButton
Private WithEvents mBtnBuilderLoad As MSForms.CommandButton
Private WithEvents mBtnBuilderSave As MSForms.CommandButton
Private WithEvents mBtnBuilderProcess As MSForms.CommandButton
Private WithEvents mBtnBuilderFormulas As MSForms.CommandButton
Private WithEvents mBtnBuilderClear As MSForms.CommandButton
Private WithEvents mBtnLineAdd As MSForms.CommandButton
Private WithEvents mBtnLineUpdate As MSForms.CommandButton
Private WithEvents mBtnLineRemove As MSForms.CommandButton

Private WithEvents mLstAssignRecipes As MSForms.ListBox
Private WithEvents mLstAssignIngredients As MSForms.ListBox
Private WithEvents mTxtInventorySearch As MSForms.TextBox
Private WithEvents mLstAssignInventory As MSForms.ListBox
Private WithEvents mLstAssignAllowed As MSForms.ListBox
Private WithEvents mBtnAssignRefresh As MSForms.CommandButton
Private WithEvents mBtnAssignRecipe As MSForms.CommandButton
Private WithEvents mBtnAssignIngredient As MSForms.CommandButton
Private WithEvents mBtnAssignAdd As MSForms.CommandButton
Private WithEvents mBtnAssignRemove As MSForms.CommandButton
Private WithEvents mBtnAssignSave As MSForms.CommandButton
Private WithEvents mBtnAssignClear As MSForms.CommandButton

Private WithEvents mLstLoaderRecipes As MSForms.ListBox
Private WithEvents mLstLoaderLines As MSForms.ListBox
Private WithEvents mLstLoaderOutput As MSForms.ListBox
Private WithEvents mBtnLoaderRefresh As MSForms.CommandButton
Private WithEvents mBtnLoaderLoad As MSForms.CommandButton
Private WithEvents mBtnLoaderClear As MSForms.CommandButton

Private WithEvents mLstManagerOutput As MSForms.ListBox
Private WithEvents mLstManagerCheck As MSForms.ListBox
Private WithEvents mBtnManagerRefresh As MSForms.CommandButton
Private WithEvents mBtnManagerPrepare As MSForms.CommandButton
Private WithEvents mBtnManagerApplyOutput As MSForms.CommandButton
Private WithEvents mBtnManagerUsed As MSForms.CommandButton
Private WithEvents mBtnManagerMade As MSForms.CommandButton
Private WithEvents mBtnManagerTotal As MSForms.CommandButton
Private WithEvents mBtnManagerNext As MSForms.CommandButton
Private WithEvents mBtnManagerPrint As MSForms.CommandButton
Private WithEvents mBtnClose As MSForms.CommandButton

Private mTxtRecipeName As MSForms.TextBox
Private mTxtRecipeId As MSForms.TextBox
Private mTxtRecipeDescription As MSForms.TextBox
Private mTxtLineProcess As MSForms.TextBox
Private mTxtLineIngredient As MSForms.TextBox
Private mTxtLinePercent As MSForms.TextBox
Private mTxtLineUom As MSForms.TextBox
Private mTxtLineAmount As MSForms.TextBox
Private mTxtOutputReal As MSForms.TextBox
Private mTxtOutputBatch As MSForms.TextBox
Private mTxtStatus As MSForms.TextBox
Private mOperatorWorkbook As Workbook
Private mBuilt As Boolean
Private mLoading As Boolean
Private mResizeInitialized As Boolean

Private Const TABLE_BUILDER_HEADER As String = "RB_AddRecipeName"
Private Const TABLE_BUILDER_LINES As String = "RecipeBuilder"
Private Const TABLE_ASSIGN_RECIPE As String = "IP_ChooseRecipe"
Private Const TABLE_ASSIGN_INGREDIENT As String = "IP_ChooseIngredient"
Private Const TABLE_ASSIGN_ITEM As String = "IP_ChooseItem"
Private Const TABLE_LOADER_CHOOSE As String = "RC_RecipeChoose"
Private Const TABLE_LOADER_LINES As String = "RecipeChooser_generated"
Private Const TABLE_MANAGER_PALETTE As String = "InventoryPalette_generated"
Private Const TABLE_MANAGER_OUTPUT As String = "ProductionOutput"
Private Const TABLE_MANAGER_CHECK As String = "Prod_invSys_Check"

Private Sub UserForm_Initialize()
    BuildLayout
End Sub

Private Sub UserForm_Activate()
    If Not mResizeInitialized Then
        On Error Resume Next
        Application.Run "modUserFormResizeWin.EnableResizableUserForm", Me
        On Error GoTo 0
        mResizeInitialized = True
    End If
End Sub

Private Sub UserForm_Terminate()
    Set mOperatorWorkbook = Nothing
End Sub

Public Sub SetOperatorWorkbook(ByVal wb As Workbook)
    If IsUsableWorkbook(wb) Then Set mOperatorWorkbook = wb
End Sub

Public Sub InitializeFromProduction()
    On Error GoTo ErrHandler

    Dim wb As Workbook
    If Not mBuilt Then BuildLayout
    Set wb = ResolveOperatorWorkbook()
    If wb Is Nothing Then
        ShowStatus "Open a Production operator workbook before using the Production form."
        Exit Sub
    End If

    mLoading = True
    On Error Resume Next
    wb.Activate
    On Error GoTo ErrHandler
    RunProductionSub1 "InitializeProductionUiForWorkbook", wb
    RefreshAllViews
    mLoading = False
    ShowStatus "Production form loaded for " & wb.Name & "."
    Exit Sub

ErrHandler:
    mLoading = False
    ShowStatus "Production form load failed: " & Err.Description
End Sub

Public Function TestInitializeForWorkbook(ByVal wb As Workbook) As Long
    SetOperatorWorkbook wb
    InitializeFromProduction
    TestInitializeForWorkbook = TestPageCount()
End Function

Public Function TestPageCount() As Long
    If mPages Is Nothing Then BuildLayout
    TestPageCount = mPages.Pages.Count
End Function

Public Function TestStatusText() As String
    If mTxtStatus Is Nothing Then BuildLayout
    TestStatusText = mTxtStatus.Text
End Function

Public Function TestFillAssignmentIoCount(ByVal values As Variant, ByVal ioValue As String) As Long
    Dim i As Long
    Dim wanted As String

    If Not mBuilt Then BuildLayout
    FillIngredientListFromArray mLstAssignIngredients, values
    wanted = UCase$(Trim$(ioValue))
    For i = 0 To mLstAssignIngredients.ListCount - 1
        If UCase$(Trim$(NzStr(mLstAssignIngredients.List(i, 4)))) = wanted Then
            TestFillAssignmentIoCount = TestFillAssignmentIoCount + 1
        End If
    Next i
End Function

Private Sub BuildLayout()
    If mBuilt Then Exit Sub

    Me.Caption = "Production"
    Me.Width = 1110
    Me.Height = 690

    Set mPages = Me.Controls.Add("Forms.MultiPage.1", "mpProduction", True)
    With mPages
        .Left = 12
        .Top = 10
        .Width = 1070
        .Height = 575
    End With
    Do While mPages.Pages.Count < 4
        mPages.Pages.Add
    Loop
    mPages.Pages(0).Caption = "Recipe Builder"
    mPages.Pages(1).Caption = "Ingredients Assignment"
    mPages.Pages(2).Caption = "Recipe Loader"
    mPages.Pages(3).Caption = "Production Manager"

    BuildRecipeBuilderPage mPages.Pages(0)
    BuildAssignmentPage mPages.Pages(1)
    BuildLoaderPage mPages.Pages(2)
    BuildManagerPage mPages.Pages(3)

    Set mTxtStatus = Me.Controls.Add("Forms.TextBox.1", "txtProductionStatus", True)
    With mTxtStatus
        .Left = 12
        .Top = 596
        .Width = 900
        .Height = 42
        .MultiLine = True
        .ScrollBars = fmScrollBarsVertical
        .Locked = True
        .Text = ""
    End With
    Set mBtnClose = AddButton(Me, "btnProductionClose", "Close", 930, 596, 150, 42)

    mBuilt = True
End Sub

Private Sub BuildRecipeBuilderPage(ByVal pg As MSForms.Page)
    AddLabel pg, "Saved Recipes", 12, 12, 180, 16
    Set mLstBuilderRecipes = AddList(pg, "lstBuilderRecipes", 12, 32, 320, 230, 3, "0 pt;130 pt;170 pt")
    AddLabel pg, "Recipe Name", 350, 12, 100, 16
    Set mTxtRecipeName = AddText(pg, "txtRecipeName", 350, 32, 240, 22)
    AddLabel pg, "Recipe ID", 610, 12, 80, 16
    Set mTxtRecipeId = AddText(pg, "txtRecipeId", 610, 32, 230, 22)
    AddLabel pg, "Description", 350, 62, 120, 16
    Set mTxtRecipeDescription = AddText(pg, "txtRecipeDescription", 350, 82, 490, 44)
    mTxtRecipeDescription.MultiLine = True

    Set mBtnBuilderRefresh = AddButton(pg, "btnBuilderRefresh", "Refresh", 860, 32, 170, 24)
    Set mBtnBuilderNew = AddButton(pg, "btnBuilderNew", "New Recipe", 860, 64, 170, 24)
    Set mBtnBuilderLoad = AddButton(pg, "btnBuilderLoad", "Load Selected", 860, 96, 170, 24)
    Set mBtnBuilderSave = AddButton(pg, "btnBuilderSave", "Save Recipe", 860, 128, 170, 24)
    Set mBtnBuilderProcess = AddButton(pg, "btnBuilderProcess", "Add Process Table", 860, 160, 170, 24)
    Set mBtnBuilderFormulas = AddButton(pg, "btnBuilderFormulas", "Save Formulas", 860, 192, 170, 24)
    Set mBtnBuilderClear = AddButton(pg, "btnBuilderClear", "Clear Builder", 860, 224, 170, 24)

    AddLabel pg, "Process", 350, 145, 70, 16
    Set mTxtLineProcess = AddText(pg, "txtLineProcess", 350, 165, 120, 22)
    AddLabel pg, "In/Out", 485, 145, 70, 16
    Set mCmbLineIo = AddCombo(pg, "cmbLineIo", 485, 165, 95, 22)
    mCmbLineIo.AddItem "USED"
    mCmbLineIo.AddItem "OUTPUT"
    mCmbLineIo.ListIndex = 0
    AddLabel pg, "Ingredient / Output", 595, 145, 150, 16
    Set mTxtLineIngredient = AddText(pg, "txtLineIngredient", 595, 165, 220, 22)
    AddLabel pg, "Percent", 350, 197, 70, 16
    Set mTxtLinePercent = AddText(pg, "txtLinePercent", 350, 217, 70, 22)
    AddLabel pg, "UOM", 435, 197, 55, 16
    Set mTxtLineUom = AddText(pg, "txtLineUom", 435, 217, 70, 22)
    AddLabel pg, "Amount", 520, 197, 70, 16
    Set mTxtLineAmount = AddText(pg, "txtLineAmount", 520, 217, 80, 22)
    Set mBtnLineAdd = AddButton(pg, "btnLineAdd", "Add Line", 620, 216, 95, 24)
    Set mBtnLineUpdate = AddButton(pg, "btnLineUpdate", "Update Line", 725, 216, 100, 24)
    Set mBtnLineRemove = AddButton(pg, "btnLineRemove", "Remove Line", 835, 256, 105, 24)

    AddLabel pg, "Recipe Builder Lines", 12, 280, 220, 16
    Set mLstBuilderLines = AddList(pg, "lstBuilderLines", 12, 300, 1018, 220, 8, "90 pt;55 pt;70 pt;210 pt;55 pt;55 pt;55 pt;70 pt")
End Sub

Private Sub BuildAssignmentPage(ByVal pg As MSForms.Page)
    AddLabel pg, "Recipes", 12, 12, 160, 16
    Set mLstAssignRecipes = AddList(pg, "lstAssignRecipes", 12, 32, 300, 180, 3, "0 pt;130 pt;150 pt")
    Set mBtnAssignRecipe = AddButton(pg, "btnAssignRecipe", "Select Recipe", 12, 220, 140, 24)
    Set mBtnAssignRefresh = AddButton(pg, "btnAssignRefresh", "Refresh", 172, 220, 140, 24)

    AddLabel pg, "Recipe Ingredients", 330, 12, 180, 16
    Set mLstAssignIngredients = AddList(pg, "lstAssignIngredients", 330, 32, 340, 212, 7, "0 pt;135 pt;45 pt;70 pt;55 pt;45 pt;45 pt")
    Set mBtnAssignIngredient = AddButton(pg, "btnAssignIngredient", "Select Ingredient", 690, 32, 150, 24)
    Set mBtnAssignSave = AddButton(pg, "btnAssignSave", "Save Assignment", 690, 64, 150, 24)
    Set mBtnAssignClear = AddButton(pg, "btnAssignClear", "Clear", 690, 96, 150, 24)

    AddLabel pg, "Search Inventory", 12, 262, 130, 16
    Set mTxtInventorySearch = AddText(pg, "txtInventorySearch", 130, 258, 230, 22)
    Set mBtnAssignAdd = AddButton(pg, "btnAssignAdd", "Add Acceptable", 380, 258, 150, 24)
    Set mBtnAssignRemove = AddButton(pg, "btnAssignRemove", "Remove Row", 548, 258, 122, 24)
    AddLabel pg, "Inventory", 12, 292, 120, 16
    Set mLstAssignInventory = AddList(pg, "lstAssignInventory", 12, 312, 510, 208, 7, "45 pt;145 pt;45 pt;58 pt;65 pt;130 pt;0 pt")
    AddLabel pg, "Acceptable Items", 540, 292, 150, 16
    Set mLstAssignAllowed = AddList(pg, "lstAssignAllowed", 540, 312, 490, 208, 6, "45 pt;160 pt;45 pt;170 pt;0 pt;0 pt")
End Sub

Private Sub BuildLoaderPage(ByVal pg As MSForms.Page)
    AddLabel pg, "Recipes", 12, 12, 140, 16
    Set mLstLoaderRecipes = AddList(pg, "lstLoaderRecipes", 12, 32, 320, 470, 3, "0 pt;140 pt;160 pt")
    Set mBtnLoaderRefresh = AddButton(pg, "btnLoaderRefresh", "Refresh", 350, 32, 150, 24)
    Set mBtnLoaderLoad = AddButton(pg, "btnLoaderLoad", "Load Recipe", 350, 64, 150, 24)
    Set mBtnLoaderClear = AddButton(pg, "btnLoaderClear", "Clear Loader", 350, 96, 150, 24)

    AddLabel pg, "Loaded Recipe Lines", 520, 12, 180, 16
    Set mLstLoaderLines = AddList(pg, "lstLoaderLines", 520, 32, 510, 260, 8, "90 pt;55 pt;70 pt;160 pt;55 pt;60 pt;70 pt;0 pt")
    AddLabel pg, "Production Output", 520, 310, 180, 16
    Set mLstLoaderOutput = AddList(pg, "lstLoaderOutput", 520, 330, 510, 172, 7, "95 pt;160 pt;45 pt;70 pt;55 pt;80 pt;45 pt")
End Sub

Private Sub BuildManagerPage(ByVal pg As MSForms.Page)
    AddLabel pg, "Production Output", 12, 12, 170, 16
    Set mLstManagerOutput = AddList(pg, "lstManagerOutput", 12, 32, 620, 260, 7, "100 pt;180 pt;45 pt;70 pt;55 pt;90 pt;45 pt")
    AddLabel pg, "Inventory Check", 12, 312, 150, 16
    Set mLstManagerCheck = AddList(pg, "lstManagerCheck", 12, 332, 620, 170, 6, "45 pt;120 pt;180 pt;45 pt;65 pt;70 pt")

    AddLabel pg, "Real Output", 660, 32, 100, 16
    Set mTxtOutputReal = AddText(pg, "txtOutputReal", 760, 28, 100, 22)
    AddLabel pg, "Batch", 660, 64, 80, 16
    Set mTxtOutputBatch = AddText(pg, "txtOutputBatch", 760, 60, 100, 22)
    Set mBtnManagerApplyOutput = AddButton(pg, "btnManagerApplyOutput", "Apply Output", 880, 44, 145, 26)

    Set mBtnManagerRefresh = AddButton(pg, "btnManagerRefresh", "Refresh", 660, 112, 165, 26)
    Set mBtnManagerPrepare = AddButton(pg, "btnManagerPrepare", "Prepare Output", 660, 154, 165, 26)
    Set mBtnManagerUsed = AddButton(pg, "btnManagerUsed", "To USED", 660, 196, 165, 26)
    Set mBtnManagerMade = AddButton(pg, "btnManagerMade", "To MADE", 660, 238, 165, 26)
    Set mBtnManagerTotal = AddButton(pg, "btnManagerTotal", "To TOTAL INV", 660, 280, 165, 26)
    Set mBtnManagerNext = AddButton(pg, "btnManagerNext", "Next Batch", 660, 322, 165, 26)
    Set mBtnManagerPrint = AddButton(pg, "btnManagerPrint", "Print Recall Codes", 660, 364, 165, 26)
End Sub

Private Function AddList(ByVal parent As Object, ByVal name As String, ByVal leftVal As Single, ByVal topVal As Single, _
                         ByVal widthVal As Single, ByVal heightVal As Single, ByVal columns As Long, _
                         ByVal widths As String) As MSForms.ListBox
    Set AddList = parent.Controls.Add("Forms.ListBox.1", name, True)
    With AddList
        .Left = leftVal
        .Top = topVal
        .Width = widthVal
        .Height = heightVal
        .ColumnCount = columns
        .ColumnWidths = widths
        .IntegralHeight = False
    End With
End Function

Private Function AddButton(ByVal parent As Object, ByVal name As String, ByVal caption As String, ByVal leftVal As Single, _
                           ByVal topVal As Single, ByVal widthVal As Single, ByVal heightVal As Single) As MSForms.CommandButton
    Set AddButton = parent.Controls.Add("Forms.CommandButton.1", name, True)
    With AddButton
        .Caption = caption
        .Left = leftVal
        .Top = topVal
        .Width = widthVal
        .Height = heightVal
        .TakeFocusOnClick = False
    End With
End Function

Private Function AddText(ByVal parent As Object, ByVal name As String, ByVal leftVal As Single, ByVal topVal As Single, _
                         ByVal widthVal As Single, ByVal heightVal As Single) As MSForms.TextBox
    Set AddText = parent.Controls.Add("Forms.TextBox.1", name, True)
    With AddText
        .Left = leftVal
        .Top = topVal
        .Width = widthVal
        .Height = heightVal
    End With
End Function

Private Function AddCombo(ByVal parent As Object, ByVal name As String, ByVal leftVal As Single, ByVal topVal As Single, _
                          ByVal widthVal As Single, ByVal heightVal As Single) As MSForms.ComboBox
    Set AddCombo = parent.Controls.Add("Forms.ComboBox.1", name, True)
    With AddCombo
        .Left = leftVal
        .Top = topVal
        .Width = widthVal
        .Height = heightVal
        .Style = fmStyleDropDownList
    End With
End Function

Private Sub AddLabel(ByVal parent As Object, ByVal caption As String, ByVal leftVal As Single, ByVal topVal As Single, _
                     ByVal widthVal As Single, ByVal heightVal As Single)
    Dim lbl As MSForms.Label
    Set lbl = parent.Controls.Add("Forms.Label.1", "lbl" & CleanControlName(caption) & CStr(parent.Controls.Count + 1), True)
    With lbl
        .Caption = caption
        .Left = leftVal
        .Top = topVal
        .Width = widthVal
        .Height = heightVal
        .Font.Bold = True
    End With
End Sub

Private Function CleanControlName(ByVal value As String) As String
    Dim i As Long
    Dim ch As String
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        If ch Like "[A-Za-z0-9]" Then CleanControlName = CleanControlName & ch
    Next i
    If CleanControlName = "" Then CleanControlName = "Control"
End Function

Private Sub RefreshAllViews()
    RefreshRecipeLists
    RefreshBuilderHeader
    RefreshBuilderLines
    RefreshAssignmentState
    RefreshLoaderState
    RefreshManagerState
End Sub

Private Sub RefreshRecipeLists()
    Dim recipes As Variant
    recipes = RunProduction0("LoadRecipeList")
    FillListFromArray mLstBuilderRecipes, recipes
    FillListFromArray mLstAssignRecipes, recipes
    FillListFromArray mLstLoaderRecipes, recipes
End Sub

Private Sub RefreshBuilderHeader()
    Dim lo As ListObject
    Set lo = ProductionTable(TABLE_BUILDER_HEADER)
    If lo Is Nothing Then Exit Sub
    mTxtRecipeName.Text = FirstRowValue(lo, "RECIPE_NAME")
    mTxtRecipeId.Text = FirstRowValue(lo, "RECIPE_ID")
    mTxtRecipeDescription.Text = FirstRowValue(lo, "DESCRIPTION")
    If Trim$(mTxtRecipeId.Text) = "" And Trim$(mTxtRecipeName.Text) = "" Then
        mTxtRecipeId.Text = NzStr(RunProduction0("GenerateRecipeIdForCurrentWorkbook"))
    End If
End Sub

Private Sub RefreshBuilderLines()
    FillListFromTable mLstBuilderLines, ProductionTable(TABLE_BUILDER_LINES), _
        Array("PROCESS", "DIAGRAM_ID", "INPUT/OUTPUT", "INGREDIENT", "PERCENT", "UOM", "AMOUNT", "INGREDIENT_ID")
End Sub

Private Sub RefreshAssignmentState()
    Dim recipeId As String
    Dim ingredientId As String
    Dim ingredients As Variant

    recipeId = NzStr(RunProduction0("GetPaletteRecipeId"))
    ingredientId = NzStr(RunProduction0("GetPaletteIngredientId"))
    If recipeId <> "" Then
        ingredients = RunProduction1("LoadIngredientListForRecipe", recipeId)
        FillIngredientListFromArray mLstAssignIngredients, ingredients
    Else
        mLstAssignIngredients.Clear
    End If
    RefreshInventoryList
    RefreshAllowedItems
End Sub

Private Sub RefreshInventoryList()
    Dim rowsData As Variant
    rowsData = RunProduction1("LoadProductionInventoryPickerItems", Trim$(mTxtInventorySearch.Text))
    FillInventoryListFromArray rowsData
End Sub

Private Sub RefreshAllowedItems()
    FillListFromTable mLstAssignAllowed, ProductionTable(TABLE_ASSIGN_ITEM), _
        Array("ROW", "ITEMS", "UOM", "DESCRIPTION", "RECIPE_ID", "INGREDIENT_ID")
End Sub

Private Sub RefreshLoaderState()
    FillListFromTable mLstLoaderLines, ProductionTable(TABLE_LOADER_LINES), _
        Array("PROCESS", "DIAGRAM_ID", "INPUT/OUTPUT", "INGREDIENT", "PERCENT", "UOM", "AMOUNT NEEDED", "INGREDIENT_ID")
    FillListFromTable mLstLoaderOutput, ProductionTable(TABLE_MANAGER_OUTPUT), _
        Array("PROCESS", "OUTPUT", "UOM", "REAL OUTPUT", "BATCH", "RECALL CODE", "ROW")
End Sub

Private Sub RefreshManagerState()
    FillListFromTable mLstManagerOutput, ProductionTable(TABLE_MANAGER_OUTPUT), _
        Array("PROCESS", "OUTPUT", "UOM", "REAL OUTPUT", "BATCH", "RECALL CODE", "ROW")
    FillListFromTable mLstManagerCheck, ProductionTable(TABLE_MANAGER_CHECK), _
        Array("ROW", "ITEM_CODE", "ITEM", "UOM", "USED", "TOTAL INV")
End Sub

Private Sub FillListFromArray(ByVal lst As MSForms.ListBox, ByVal values As Variant)
    Dim r As Long
    Dim c As Long
    Dim lb As Long
    Dim ub As Long

    lst.Clear
    If IsEmpty(values) Then Exit Sub
    If Not IsArray(values) Then Exit Sub
    lb = LBound(values, 1)
    ub = UBound(values, 1)
    For r = lb To ub
        lst.AddItem NzStr(values(r, LBound(values, 2)))
        For c = LBound(values, 2) + 1 To UBound(values, 2)
            If c - LBound(values, 2) < lst.ColumnCount Then
                lst.List(lst.ListCount - 1, c - LBound(values, 2)) = NzStr(values(r, c))
            End If
        Next c
    Next r
End Sub

Private Sub FillIngredientListFromArray(ByVal lst As MSForms.ListBox, ByVal values As Variant)
    Dim r As Long
    Dim c As Long

    lst.Clear
    If IsEmpty(values) Then Exit Sub
    If Not IsArray(values) Then Exit Sub
    For r = LBound(values, 1) To UBound(values, 1)
        lst.AddItem NzStr(values(r, LBound(values, 2)))
        For c = LBound(values, 2) + 1 To UBound(values, 2)
            If c - LBound(values, 2) < lst.ColumnCount Then
                lst.List(lst.ListCount - 1, c - LBound(values, 2)) = NzStr(values(r, c))
            End If
        Next c
    Next r
End Sub

Private Sub FillListFromTable(ByVal lst As MSForms.ListBox, ByVal lo As ListObject, ByVal headers As Variant)
    Dim arr As Variant
    Dim r As Long
    Dim c As Long
    Dim colIdx As Long

    lst.Clear
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    arr = lo.DataBodyRange.Value
    For r = 1 To UBound(arr, 1)
        lst.AddItem CellText(arr, r, lo, CStr(headers(LBound(headers))))
        For c = LBound(headers) + 1 To UBound(headers)
            colIdx = c - LBound(headers)
            If colIdx < lst.ColumnCount Then lst.List(lst.ListCount - 1, colIdx) = CellText(arr, r, lo, CStr(headers(c)))
        Next c
    Next r
End Sub

Private Sub FillInventoryList(ByVal lo As ListObject, ByVal filterText As String)
    Dim arr As Variant
    Dim r As Long
    Dim rowVal As String
    Dim itemVal As String
    Dim uomVal As String
    Dim totalVal As String
    Dim locVal As String
    Dim descVal As String
    Dim itemCode As String
    Dim haystack As String

    mLstAssignInventory.Clear
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    arr = lo.DataBodyRange.Value
    filterText = LCase$(Trim$(filterText))
    For r = 1 To UBound(arr, 1)
        rowVal = CellText(arr, r, lo, "ROW")
        itemVal = CellText(arr, r, lo, "ITEM")
        uomVal = CellText(arr, r, lo, "UOM")
        totalVal = CellText(arr, r, lo, "TOTAL INV")
        locVal = CellText(arr, r, lo, "LOCATION")
        descVal = CellText(arr, r, lo, "DESCRIPTION")
        itemCode = CellText(arr, r, lo, "ITEM_CODE")
        haystack = LCase$(rowVal & " " & itemVal & " " & descVal & " " & itemCode & " " & locVal)
        If filterText = "" Or InStr(1, haystack, filterText, vbTextCompare) > 0 Then
            mLstAssignInventory.AddItem rowVal
            mLstAssignInventory.List(mLstAssignInventory.ListCount - 1, 1) = itemVal
            mLstAssignInventory.List(mLstAssignInventory.ListCount - 1, 2) = uomVal
            mLstAssignInventory.List(mLstAssignInventory.ListCount - 1, 3) = totalVal
            mLstAssignInventory.List(mLstAssignInventory.ListCount - 1, 4) = locVal
            mLstAssignInventory.List(mLstAssignInventory.ListCount - 1, 5) = descVal
            mLstAssignInventory.List(mLstAssignInventory.ListCount - 1, 6) = itemCode
        End If
    Next r
End Sub

Private Sub FillInventoryListFromArray(ByVal values As Variant)
    Dim r As Long

    mLstAssignInventory.Clear
    If IsEmpty(values) Then Exit Sub
    If Not IsArray(values) Then Exit Sub

    For r = LBound(values, 1) To UBound(values, 1)
        mLstAssignInventory.AddItem NzStr(values(r, 1))
        If mLstAssignInventory.ColumnCount > 1 Then mLstAssignInventory.List(mLstAssignInventory.ListCount - 1, 1) = NzStr(values(r, 2))
        If mLstAssignInventory.ColumnCount > 2 Then mLstAssignInventory.List(mLstAssignInventory.ListCount - 1, 2) = NzStr(values(r, 3))
        If mLstAssignInventory.ColumnCount > 3 Then mLstAssignInventory.List(mLstAssignInventory.ListCount - 1, 3) = NzStr(values(r, 4))
        If mLstAssignInventory.ColumnCount > 4 Then mLstAssignInventory.List(mLstAssignInventory.ListCount - 1, 4) = NzStr(values(r, 5))
        If mLstAssignInventory.ColumnCount > 5 Then mLstAssignInventory.List(mLstAssignInventory.ListCount - 1, 5) = NzStr(values(r, 6))
        If mLstAssignInventory.ColumnCount > 6 Then mLstAssignInventory.List(mLstAssignInventory.ListCount - 1, 6) = NzStr(values(r, 7))
    Next r
End Sub

Private Function CellText(ByVal arr As Variant, ByVal rowIndex As Long, ByVal lo As ListObject, ByVal headerName As String) As String
    Dim colIndex As Long
    colIndex = ProductionColumnIndex(lo, headerName)
    If colIndex <= 0 Then Exit Function
    On Error Resume Next
    CellText = NzStr(arr(rowIndex, colIndex))
    On Error GoTo 0
End Function

Private Function ProductionTable(ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    Set ws = RunProductionObject0("GetProductionSheet")
    If ws Is Nothing Then Exit Function
    Set ProductionTable = RunProductionObject2("GetListObject", ws, tableName)
End Function

Private Function InventoryTable() As ListObject
    Dim ws As Worksheet
    Set ws = RunProductionObject1("SheetExists", "InventoryManagement")
    If ws Is Nothing Then Exit Function
    Set InventoryTable = RunProductionObject2("GetListObject", ws, "invSys")
End Function

Private Function ResolveOperatorWorkbook() As Workbook
    If IsUsableWorkbook(mOperatorWorkbook) Then
        Set ResolveOperatorWorkbook = mOperatorWorkbook
        Exit Function
    End If
    If IsUsableWorkbook(Application.ActiveWorkbook) Then
        Set mOperatorWorkbook = Application.ActiveWorkbook
        Set ResolveOperatorWorkbook = mOperatorWorkbook
        Exit Function
    End If

    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If IsUsableWorkbook(wb) Then
            If WorkbookHasSheet(wb, "Production") Then
                Set mOperatorWorkbook = wb
                Set ResolveOperatorWorkbook = wb
                Exit Function
            End If
        End If
    Next wb
End Function

Private Function IsUsableWorkbook(ByVal wb As Workbook) As Boolean
    On Error GoTo CleanFail
    If wb Is Nothing Then Exit Function
    If wb.IsAddin Then Exit Function
    IsUsableWorkbook = True
    Exit Function
CleanFail:
    IsUsableWorkbook = False
End Function

Private Function WorkbookHasSheet(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    On Error Resume Next
    WorkbookHasSheet = Not wb.Worksheets(sheetName) Is Nothing
    On Error GoTo 0
End Function

Private Sub WriteRecipeHeaderFromForm()
    Dim lo As ListObject
    Set lo = ProductionTable(TABLE_BUILDER_HEADER)
    If lo Is Nothing Then Exit Sub
    EnsureTableRow lo
    If Trim$(mTxtRecipeId.Text) = "" Then mTxtRecipeId.Text = NzStr(RunProduction0("GenerateRecipeIdForCurrentWorkbook"))
    SetFirstRowValue lo, "RECIPE_NAME", mTxtRecipeName.Text
    SetFirstRowValue lo, "RECIPE_ID", mTxtRecipeId.Text
    SetFirstRowValue lo, "DESCRIPTION", mTxtRecipeDescription.Text
End Sub

Private Sub ClearRecipeBuilderForNew()
    Dim loHeader As ListObject
    Dim loLines As ListObject

    Set loHeader = ProductionTable(TABLE_BUILDER_HEADER)
    Set loLines = ProductionTable(TABLE_BUILDER_LINES)
    If Not loHeader Is Nothing Then
        If Not loHeader.DataBodyRange Is Nothing Then loHeader.DataBodyRange.ClearContents
    End If
    If Not loLines Is Nothing Then ClearListRows loLines
    mTxtRecipeName.Text = ""
    mTxtRecipeId.Text = NzStr(RunProduction0("GenerateRecipeIdForCurrentWorkbook"))
    mTxtRecipeDescription.Text = ""
    ClearLineInputs
    RefreshBuilderLines
    ShowStatus "Ready for a new recipe."
End Sub

Private Sub ClearLineInputs()
    mTxtLineProcess.Text = ""
    If Not mCmbLineIo Is Nothing Then mCmbLineIo.ListIndex = 0
    mTxtLineIngredient.Text = ""
    mTxtLinePercent.Text = ""
    mTxtLineUom.Text = ""
    mTxtLineAmount.Text = ""
End Sub

Private Sub LoadSelectedBuilderLine()
    Dim idx As Long

    idx = mLstBuilderLines.ListIndex
    If idx < 0 Then Exit Sub
    mTxtLineProcess.Text = NzStr(mLstBuilderLines.List(idx, 0))
    SetLineIo NzStr(mLstBuilderLines.List(idx, 2))
    mTxtLineIngredient.Text = NzStr(mLstBuilderLines.List(idx, 3))
    mTxtLinePercent.Text = NzStr(mLstBuilderLines.List(idx, 4))
    mTxtLineUom.Text = NzStr(mLstBuilderLines.List(idx, 5))
    mTxtLineAmount.Text = NzStr(mLstBuilderLines.List(idx, 6))
End Sub

Private Sub SetLineIo(ByVal ioValue As String)
    Dim v As String
    v = UCase$(Trim$(ioValue))
    If v = "OUTPUT" Or v = "MADE" Then
        mCmbLineIo.ListIndex = 1
    Else
        mCmbLineIo.ListIndex = 0
    End If
End Sub

Private Function LineIoValue() As String
    If mCmbLineIo.ListIndex = 1 Then
        LineIoValue = "OUTPUT"
    Else
        LineIoValue = "USED"
    End If
End Function

Private Sub AddRecipeBuilderLine()
    Dim lo As ListObject
    Dim lr As ListRow

    If Trim$(mTxtLineIngredient.Text) = "" Then
        ShowStatus "Enter an ingredient or output name before adding a recipe line."
        Exit Sub
    End If
    Set lo = ProductionTable(TABLE_BUILDER_LINES)
    If lo Is Nothing Then
        ShowStatus "RecipeBuilder table is missing."
        Exit Sub
    End If
    Set lr = lo.ListRows.Add
    WriteRecipeLineToRow lo, lr.Index
    RefreshBuilderLines
    ClearLineInputs
    ShowStatus "Recipe line added."
End Sub

Private Sub UpdateSelectedRecipeBuilderLine()
    Dim lo As ListObject
    Dim idx As Long

    idx = mLstBuilderLines.ListIndex
    If idx < 0 Then
        ShowStatus "Select a recipe line to update."
        Exit Sub
    End If
    Set lo = ProductionTable(TABLE_BUILDER_LINES)
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    If idx + 1 > lo.ListRows.Count Then Exit Sub
    WriteRecipeLineToRow lo, idx + 1
    RefreshBuilderLines
    mLstBuilderLines.ListIndex = idx
    ShowStatus "Recipe line updated."
End Sub

Private Sub RemoveSelectedRecipeBuilderLine()
    Dim lo As ListObject
    Dim idx As Long

    idx = mLstBuilderLines.ListIndex
    If idx < 0 Then
        ShowStatus "Select a recipe line to remove."
        Exit Sub
    End If
    Set lo = ProductionTable(TABLE_BUILDER_LINES)
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    If idx + 1 <= lo.ListRows.Count Then lo.ListRows(idx + 1).Delete
    RefreshBuilderLines
    ClearLineInputs
    ShowStatus "Recipe line removed."
End Sub

Private Sub WriteRecipeLineToRow(ByVal lo As ListObject, ByVal rowIndex As Long)
    If lo Is Nothing Then Exit Sub
    If rowIndex < 1 Then Exit Sub
    EnsureTableRow lo
    Do While lo.ListRows.Count < rowIndex
        lo.ListRows.Add AlwaysInsert:=True
    Loop
    SetCellByHeader lo, rowIndex, "PROCESS", mTxtLineProcess.Text
    SetCellByHeader lo, rowIndex, "INPUT/OUTPUT", LineIoValue()
    SetCellByHeader lo, rowIndex, "INGREDIENT", mTxtLineIngredient.Text
    SetCellByHeader lo, rowIndex, "PERCENT", mTxtLinePercent.Text
    SetCellByHeader lo, rowIndex, "UOM", mTxtLineUom.Text
    SetCellByHeader lo, rowIndex, "AMOUNT", mTxtLineAmount.Text
    If Trim$(CellByHeader(lo, rowIndex, "RECIPE_LIST_ROW")) = "" Then SetCellByHeader lo, rowIndex, "RECIPE_LIST_ROW", rowIndex
    If Trim$(CellByHeader(lo, rowIndex, "INGREDIENT_ID")) = "" Then SetCellByHeader lo, rowIndex, "INGREDIENT_ID", BuildFormGuid()
    If Trim$(CellByHeader(lo, rowIndex, "GUID")) = "" Then SetCellByHeader lo, rowIndex, "GUID", BuildFormGuid()
End Sub

Private Sub SelectAssignmentRecipeFromList()
    Dim idx As Long
    Dim recipeId As String
    Dim lo As ListObject

    idx = mLstAssignRecipes.ListIndex
    If idx < 0 Then
        ShowStatus "Select a recipe first."
        Exit Sub
    End If
    recipeId = NzStr(mLstAssignRecipes.List(idx, 0))
    Set lo = ProductionTable(TABLE_ASSIGN_RECIPE)
    If lo Is Nothing Then
        ShowStatus "IP_ChooseRecipe table is missing."
        Exit Sub
    End If
    EnsureTableRow lo
    SetFirstRowValue lo, "RECIPE_ID", recipeId
    SetFirstRowValue lo, "RECIPE_NAME", NzStr(mLstAssignRecipes.List(idx, 1))
    SetFirstRowValue lo, "DESCRIPTION", NzStr(mLstAssignRecipes.List(idx, 2))
    SetFirstRowValue lo, "GUID", recipeId
    RunProductionSub1 "HandlePaletteRecipeSelected", recipeId
    RefreshAssignmentState
    ShowStatus "Selected assignment recipe: " & NzStr(mLstAssignRecipes.List(idx, 1))
End Sub

Private Sub SelectAssignmentIngredientFromList()
    Dim idx As Long
    Dim recipeId As String
    Dim ingredientId As String
    Dim lo As ListObject

    idx = mLstAssignIngredients.ListIndex
    If idx < 0 Then
        ShowStatus "Select an ingredient first."
        Exit Sub
    End If
    recipeId = NzStr(RunProduction0("GetPaletteRecipeId"))
    ingredientId = NzStr(mLstAssignIngredients.List(idx, 0))
    If recipeId = "" Or ingredientId = "" Then
        ShowStatus "Select a recipe and ingredient first."
        Exit Sub
    End If
    Set lo = ProductionTable(TABLE_ASSIGN_INGREDIENT)
    If lo Is Nothing Then
        ShowStatus "IP_ChooseIngredient table is missing."
        Exit Sub
    End If
    EnsureTableRow lo
    SetFirstRowValue lo, "RECIPE_ID", recipeId
    SetFirstRowValue lo, "INGREDIENT_ID", ingredientId
    SetFirstRowValue lo, "INGREDIENT", NzStr(mLstAssignIngredients.List(idx, 1))
    SetFirstRowValue lo, "UOM", NzStr(mLstAssignIngredients.List(idx, 2))
    SetFirstRowValue lo, "PROCESS", NzStr(mLstAssignIngredients.List(idx, 3))
    SetFirstRowValue lo, "DESCRIPTION", ""
    SetFirstRowValue lo, "QUANTITY", NzStr(mLstAssignIngredients.List(idx, 5))
    RunProductionSub2 "HandlePaletteIngredientSelected", recipeId, ingredientId
    RefreshAllowedItems
    ShowStatus "Selected ingredient: " & NzStr(mLstAssignIngredients.List(idx, 1))
End Sub

Private Sub AddSelectedInventoryToAllowed()
    Dim idx As Long
    Dim lo As ListObject
    Dim lr As ListRow
    Dim recipeId As String
    Dim ingredientId As String

    idx = mLstAssignInventory.ListIndex
    If idx < 0 Then
        ShowStatus "Select an inventory row first."
        Exit Sub
    End If
    recipeId = NzStr(RunProduction0("GetPaletteRecipeId"))
    ingredientId = NzStr(RunProduction0("GetPaletteIngredientId"))
    If recipeId = "" Or ingredientId = "" Then
        ShowStatus "Select a recipe and ingredient before adding acceptable inventory."
        Exit Sub
    End If
    Set lo = ProductionTable(TABLE_ASSIGN_ITEM)
    If lo Is Nothing Then
        ShowStatus "IP_ChooseItem table is missing."
        Exit Sub
    End If
    Set lr = lo.ListRows.Add
    SetRowValue lr, lo, "ROW", NzStr(mLstAssignInventory.List(idx, 0))
    SetRowValue lr, lo, "ITEMS", NzStr(mLstAssignInventory.List(idx, 1))
    SetRowValue lr, lo, "UOM", NzStr(mLstAssignInventory.List(idx, 2))
    SetRowValue lr, lo, "DESCRIPTION", NzStr(mLstAssignInventory.List(idx, 5))
    SetRowValue lr, lo, "RECIPE_ID", recipeId
    SetRowValue lr, lo, "INGREDIENT_ID", ingredientId
    RefreshAllowedItems
    ShowStatus "Added acceptable ingredient row " & NzStr(mLstAssignInventory.List(idx, 0)) & "."
End Sub

Private Sub RemoveSelectedAllowedRow()
    Dim idx As Long
    Dim lo As ListObject

    idx = mLstAssignAllowed.ListIndex
    If idx < 0 Then
        ShowStatus "Select an acceptable item row to remove."
        Exit Sub
    End If
    Set lo = ProductionTable(TABLE_ASSIGN_ITEM)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If idx + 1 <= lo.ListRows.Count Then lo.ListRows(idx + 1).Delete
    RefreshAllowedItems
    ShowStatus "Removed acceptable item row."
End Sub

Private Sub LoadSelectedRecipeIntoBuilder()
    Dim idx As Long
    idx = mLstBuilderRecipes.ListIndex
    If idx < 0 Then
        ShowStatus "Select a saved recipe first."
        Exit Sub
    End If
    RunProductionSub1 "LoadRecipeFromRecipes", NzStr(mLstBuilderRecipes.List(idx, 0))
    RefreshBuilderHeader
    RefreshBuilderLines
    ShowStatus "Loaded recipe into builder: " & NzStr(mLstBuilderRecipes.List(idx, 1))
End Sub

Private Sub LoadSelectedRecipeIntoLoader()
    Dim idx As Long
    Dim prepared As Variant

    idx = mLstLoaderRecipes.ListIndex
    If idx < 0 Then
        ShowStatus "Select a recipe first."
        Exit Sub
    End If
    RunProductionSub1 "LoadRecipeChooser", NzStr(mLstLoaderRecipes.List(idx, 0))
    prepared = RunProduction0("PrepareProductionOutputForCurrentRecipe")
    RefreshLoaderState
    RefreshManagerState
    If CBool(prepared) Then
        ShowStatus "Loaded recipe and prepared output rows: " & NzStr(mLstLoaderRecipes.List(idx, 1))
    Else
        ShowStatus "Loaded recipe, but no output rows were prepared: " & NzStr(mLstLoaderRecipes.List(idx, 1))
    End If
End Sub

Private Sub PrepareProductionOutput()
    Dim result As Variant

    result = RunProduction0("PrepareProductionOutputForCurrentRecipe")
    RefreshLoaderState
    RefreshManagerState
    If CBool(result) Then
        ShowStatus "Production output prepared."
    Else
        ShowStatus "Production output was not prepared. Load a recipe with OUTPUT lines first."
    End If
End Sub

Private Sub LoadSelectedProductionOutput()
    Dim idx As Long

    idx = mLstManagerOutput.ListIndex
    If idx < 0 Then Exit Sub
    mTxtOutputReal.Text = NzStr(mLstManagerOutput.List(idx, 3))
    mTxtOutputBatch.Text = NzStr(mLstManagerOutput.List(idx, 4))
End Sub

Private Sub ApplySelectedProductionOutput()
    Dim idx As Long
    Dim lo As ListObject

    idx = mLstManagerOutput.ListIndex
    If idx < 0 Then
        ShowStatus "Select a ProductionOutput row first."
        Exit Sub
    End If
    Set lo = ProductionTable(TABLE_MANAGER_OUTPUT)
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    If idx + 1 > lo.ListRows.Count Then Exit Sub
    SetCellByHeader lo, idx + 1, "REAL OUTPUT", mTxtOutputReal.Text
    SetCellByHeader lo, idx + 1, "BATCH", mTxtOutputBatch.Text
    RefreshManagerState
    If idx < mLstManagerOutput.ListCount Then mLstManagerOutput.ListIndex = idx
    ShowStatus "Production output row updated."
End Sub

Private Sub EnsureTableRow(ByVal lo As ListObject)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then lo.ListRows.Add AlwaysInsert:=True
End Sub

Private Sub ClearListRows(ByVal lo As ListObject)
    If lo Is Nothing Then Exit Sub
    On Error Resume Next
    Do While lo.ListRows.Count > 0
        lo.ListRows(1).Delete
    Loop
    On Error GoTo 0
End Sub

Private Function CellByHeader(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal headerName As String) As String
    Dim colIndex As Long
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    If rowIndex < 1 Or rowIndex > lo.DataBodyRange.Rows.Count Then Exit Function
    colIndex = ProductionColumnIndex(lo, headerName)
    If colIndex <= 0 Then Exit Function
    CellByHeader = NzStr(lo.DataBodyRange.Cells(rowIndex, colIndex).Value)
End Function

Private Sub SetCellByHeader(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal headerName As String, ByVal value As Variant)
    Dim colIndex As Long
    If lo Is Nothing Then Exit Sub
    If rowIndex < 1 Then Exit Sub
    EnsureTableRow lo
    Do While lo.ListRows.Count < rowIndex
        lo.ListRows.Add AlwaysInsert:=True
    Loop
    colIndex = ProductionColumnIndex(lo, headerName)
    If colIndex <= 0 Then Exit Sub
    lo.DataBodyRange.Cells(rowIndex, colIndex).Value = value
End Sub

Private Function BuildFormGuid() As String
    On Error Resume Next
    BuildFormGuid = NzStr(Application.Run("modUR_Snapshot.GenerateGUID"))
    If BuildFormGuid = "" Then
        BuildFormGuid = Format$(Now, "yyyymmddhhnnss") & "-" & CStr(Int(Rnd() * 1000000#))
    End If
    On Error GoTo 0
End Function

Private Function FirstRowValue(ByVal lo As ListObject, ByVal headerName As String) As String
    Dim colIndex As Long
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    colIndex = ProductionColumnIndex(lo, headerName)
    If colIndex <= 0 Then Exit Function
    FirstRowValue = NzStr(lo.DataBodyRange.Cells(1, colIndex).Value)
End Function

Private Sub SetFirstRowValue(ByVal lo As ListObject, ByVal headerName As String, ByVal value As Variant)
    Dim colIndex As Long
    If lo Is Nothing Then Exit Sub
    EnsureTableRow lo
    colIndex = ProductionColumnIndex(lo, headerName)
    If colIndex <= 0 Then Exit Sub
    lo.DataBodyRange.Cells(1, colIndex).Value = value
End Sub

Private Sub SetRowValue(ByVal lr As ListRow, ByVal lo As ListObject, ByVal headerName As String, ByVal value As Variant)
    Dim colIndex As Long
    If lr Is Nothing Or lo Is Nothing Then Exit Sub
    colIndex = ProductionColumnIndex(lo, headerName)
    If colIndex <= 0 Then Exit Sub
    lr.Range.Cells(1, colIndex).Value = value
End Sub

Private Function ProductionColumnIndex(ByVal lo As ListObject, ByVal headerName As String) As Long
    On Error Resume Next
    ProductionColumnIndex = CLng(Application.Run("mProduction.ColumnIndex", lo, headerName))
    On Error GoTo 0
End Function

Private Function RunProduction0(ByVal procName As String) As Variant
    RunProduction0 = Application.Run("mProduction." & procName)
End Function

Private Function RunProduction1(ByVal procName As String, ByVal arg1 As Variant) As Variant
    RunProduction1 = Application.Run("mProduction." & procName, arg1)
End Function

Private Sub RunProductionSub0(ByVal procName As String)
    Application.Run "mProduction." & procName
End Sub

Private Sub RunProductionSub1(ByVal procName As String, ByVal arg1 As Variant)
    Application.Run "mProduction." & procName, arg1
End Sub

Private Sub RunProductionSub2(ByVal procName As String, ByVal arg1 As Variant, ByVal arg2 As Variant)
    Application.Run "mProduction." & procName, arg1, arg2
End Sub

Private Function RunProductionObject0(ByVal procName As String) As Object
    Set RunProductionObject0 = Application.Run("mProduction." & procName)
End Function

Private Function RunProductionObject1(ByVal procName As String, ByVal arg1 As Variant) As Object
    Set RunProductionObject1 = Application.Run("mProduction." & procName, arg1)
End Function

Private Function RunProductionObject2(ByVal procName As String, ByVal arg1 As Variant, ByVal arg2 As Variant) As Object
    Set RunProductionObject2 = Application.Run("mProduction." & procName, arg1, arg2)
End Function

Private Sub ShowStatus(ByVal messageText As String)
    If mTxtStatus Is Nothing Then Exit Sub
    mTxtStatus.Text = messageText
End Sub

Private Function NzStr(ByVal value As Variant) As String
    If IsError(value) Then Exit Function
    If IsNull(value) Then Exit Function
    If IsEmpty(value) Then Exit Function
    NzStr = CStr(value)
End Function

Private Sub mLstBuilderRecipes_Click()
    If mLoading Then Exit Sub
    If mLstBuilderRecipes.ListIndex < 0 Then Exit Sub
    mTxtRecipeId.Text = NzStr(mLstBuilderRecipes.List(mLstBuilderRecipes.ListIndex, 0))
    mTxtRecipeName.Text = NzStr(mLstBuilderRecipes.List(mLstBuilderRecipes.ListIndex, 1))
    mTxtRecipeDescription.Text = NzStr(mLstBuilderRecipes.List(mLstBuilderRecipes.ListIndex, 2))
End Sub

Private Sub mLstBuilderLines_Click()
    If mLoading Then Exit Sub
    LoadSelectedBuilderLine
End Sub

Private Sub mBtnBuilderRefresh_Click()
    RefreshRecipeLists
    RefreshBuilderHeader
    RefreshBuilderLines
    ShowStatus "Recipe Builder refreshed."
End Sub

Private Sub mBtnBuilderNew_Click()
    ClearRecipeBuilderForNew
End Sub

Private Sub mBtnBuilderLoad_Click()
    LoadSelectedRecipeIntoBuilder
End Sub

Private Sub mBtnBuilderSave_Click()
    WriteRecipeHeaderFromForm
    RunProductionSub0 "BtnSaveRecipe"
    RefreshRecipeLists
    RefreshBuilderLines
    ShowStatus "Save Recipe completed."
End Sub

Private Sub mBtnBuilderProcess_Click()
    WriteRecipeHeaderFromForm
    RunProductionSub0 "BtnBuildRecipeProcessTables"
    RefreshBuilderLines
    ShowStatus "Process table action completed."
End Sub

Private Sub mBtnBuilderFormulas_Click()
    WriteRecipeHeaderFromForm
    RunProductionSub0 "BtnSaveFormulas"
    ShowStatus "Save Formulas completed."
End Sub

Private Sub mBtnBuilderClear_Click()
    RunProductionSub0 "BtnClearRecipeBuilder"
    RefreshBuilderHeader
    RefreshBuilderLines
    ShowStatus "Recipe Builder cleared."
End Sub

Private Sub mBtnLineAdd_Click()
    AddRecipeBuilderLine
End Sub

Private Sub mBtnLineUpdate_Click()
    UpdateSelectedRecipeBuilderLine
End Sub

Private Sub mBtnLineRemove_Click()
    RemoveSelectedRecipeBuilderLine
End Sub

Private Sub mLstAssignRecipes_Click()
    If mLoading Then Exit Sub
    SelectAssignmentRecipeFromList
End Sub

Private Sub mLstAssignIngredients_Click()
    If mLoading Then Exit Sub
    SelectAssignmentIngredientFromList
End Sub

Private Sub mTxtInventorySearch_Change()
    If mLoading Then Exit Sub
    RefreshInventoryList
End Sub

Private Sub mBtnAssignRefresh_Click()
    RefreshRecipeLists
    RefreshAssignmentState
    ShowStatus "Ingredients Assignment refreshed."
End Sub

Private Sub mBtnAssignRecipe_Click()
    SelectAssignmentRecipeFromList
End Sub

Private Sub mBtnAssignIngredient_Click()
    SelectAssignmentIngredientFromList
End Sub

Private Sub mBtnAssignAdd_Click()
    AddSelectedInventoryToAllowed
End Sub

Private Sub mBtnAssignRemove_Click()
    RemoveSelectedAllowedRow
End Sub

Private Sub mBtnAssignSave_Click()
    RunProductionSub0 "BtnSavePalette"
    RefreshAllowedItems
    ShowStatus "Save Assignment completed. " & NzStr(RunProduction0("GetPaletteSaveDiagnostic"))
End Sub

Private Sub mBtnAssignClear_Click()
    RunProductionSub0 "BtnClearPaletteBuilder"
    RefreshAssignmentState
    ShowStatus "Ingredients Assignment cleared."
End Sub

Private Sub mBtnLoaderRefresh_Click()
    RefreshRecipeLists
    RefreshLoaderState
    ShowStatus "Recipe Loader refreshed."
End Sub

Private Sub mBtnLoaderLoad_Click()
    LoadSelectedRecipeIntoLoader
End Sub

Private Sub mBtnLoaderClear_Click()
    RunProductionSub0 "BtnClearRecipeChooser"
    RefreshLoaderState
    RefreshManagerState
    ShowStatus "Recipe Loader cleared."
End Sub

Private Sub mBtnManagerRefresh_Click()
    RefreshManagerState
    ShowStatus "Production Manager refreshed."
End Sub

Private Sub mLstManagerOutput_Click()
    If mLoading Then Exit Sub
    LoadSelectedProductionOutput
End Sub

Private Sub mBtnManagerPrepare_Click()
    PrepareProductionOutput
End Sub

Private Sub mBtnManagerApplyOutput_Click()
    ApplySelectedProductionOutput
End Sub

Private Sub mBtnManagerUsed_Click()
    RunProductionSub0 "BtnToUsed"
    RefreshManagerState
    ShowStatus "To USED completed."
End Sub

Private Sub mBtnManagerMade_Click()
    If mLstManagerOutput.ListIndex >= 0 Then ApplySelectedProductionOutput
    RunProductionSub0 "BtnToMade"
    RefreshManagerState
    ShowStatus "To MADE completed."
End Sub

Private Sub mBtnManagerTotal_Click()
    If mLstManagerOutput.ListIndex >= 0 Then ApplySelectedProductionOutput
    RunProductionSub0 "BtnToTotalInv"
    RefreshManagerState
    ShowStatus "To TOTAL INV completed."
End Sub

Private Sub mBtnManagerNext_Click()
    RunProductionSub0 "BtnNextBatch"
    RefreshLoaderState
    RefreshManagerState
    ShowStatus "Next Batch completed."
End Sub

Private Sub mBtnManagerPrint_Click()
    RunProductionSub0 "BtnPrintRecallCodes"
    ShowStatus "Print Recall Codes completed. " & NzStr(RunProduction0("GetRecallPrintDiagnostic"))
End Sub

Private Sub mBtnClose_Click()
    Unload Me
End Sub
