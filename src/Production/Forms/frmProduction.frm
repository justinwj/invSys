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

Private Const RUN_LOADER_RECIPE_WIDTHS As String = "0 pt;120 pt;130 pt"
Private Const RUN_LOADER_LINE_WIDTHS As String = "85 pt;0 pt;55 pt;155 pt;50 pt;45 pt;65 pt;0 pt"
Private Const RUN_PALETTE_WIDTHS As String = "0 pt;0 pt;125 pt;35 pt;145 pt;48 pt;58 pt;38 pt;70 pt;90 pt"
Private Const RUN_OUTPUT_WIDTHS As String = "75 pt;115 pt;35 pt;50 pt;45 pt;65 pt;35 pt"
Private Const RUN_CHECK_WIDTHS As String = "38 pt;95 pt;200 pt;45 pt;60 pt;70 pt"

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
Private WithEvents mLstRunPalette As MSForms.ListBox
Private WithEvents mLstRunTree As MSForms.ListBox
Private WithEvents mCmbRunProcess As MSForms.ComboBox
Private WithEvents mCmbTreeRunProcess As MSForms.ComboBox
Private WithEvents mCmbRunLocation As MSForms.ComboBox
Private WithEvents mCmbTreeRunLocation As MSForms.ComboBox
Private WithEvents mBtnLoaderRefresh As MSForms.CommandButton
Private WithEvents mBtnLoaderLoad As MSForms.CommandButton
Private WithEvents mBtnLoaderClear As MSForms.CommandButton
Private WithEvents mBtnRunTreeExpandAll As MSForms.CommandButton
Private WithEvents mBtnRunTreeCollapseAll As MSForms.CommandButton
Private WithEvents mBtnRunTreeApplyPalette As MSForms.CommandButton

Private WithEvents mLstManagerOutput As MSForms.ListBox
Private WithEvents mLstManagerCheck As MSForms.ListBox
Private WithEvents mBtnRunApplyPalette As MSForms.CommandButton
Private WithEvents mBtnManagerCheckIn As MSForms.CommandButton
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
Private mTxtRecipeRowBudget As MSForms.TextBox
Private mTxtLineProcess As MSForms.TextBox
Private mTxtLineIngredient As MSForms.TextBox
Private mTxtLinePercent As MSForms.TextBox
Private mTxtLineUom As MSForms.TextBox
Private mTxtLineAmount As MSForms.TextBox
Private WithEvents mTxtPaletteSplit As MSForms.TextBox
Private WithEvents mTxtPaletteQty As MSForms.TextBox
Private WithEvents mTxtTreePaletteSplit As MSForms.TextBox
Private WithEvents mTxtTreePaletteQty As MSForms.TextBox
Private WithEvents mTxtOutputReal As MSForms.TextBox
Private mTxtOutputBatch As MSForms.TextBox
Private mTxtOutputTotal As MSForms.TextBox
Private mTxtStatus As MSForms.TextBox
Private mOperatorWorkbook As Workbook
Private mInventoryRows As Variant
Private mInventoryCacheLoaded As Boolean
Private mBuilt As Boolean
Private mLoading As Boolean
Private mResizeInitialized As Boolean
Private mResizingLayout As Boolean
Private mRunTreeCollapsed As Object
Private mRunSplitOverrides As Object
Private mRunBaseQtyByKey As Object
Private mUpdatingPaletteInputs As Boolean
Private mPaletteInputSource As String

Private Const ASSIGN_INVENTORY_MAX_VISIBLE As Long = 250
Private Const PRODUCTION_BASE_WIDTH As Double = 1110
Private Const PRODUCTION_BASE_HEIGHT As Double = 690
Private Const PRODUCTION_DEFAULT_ROW_BUDGET As Long = 50
Private Const PRODUCTION_MAX_ROW_BUDGET As Long = 1000
Private Const RUN_TREE_PARENT_MARKER As String = "__RUN_TREE_PARENT__"

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
    On Error GoTo FailInitialize
    BuildLayout
    Exit Sub

FailInitialize:
    MsgBox "Production form failed to initialize: " & Err.Description, vbExclamation, "invSys Production"
End Sub

Private Sub UserForm_Activate()
    On Error GoTo FailActivate
    If Not mResizeInitialized Then
        On Error Resume Next
        Application.Run "modUserFormResizeWin.EnableResizableUserForm", Me
        On Error GoTo FailActivate
        mResizeInitialized = True
    End If
    ResizeProductionLayout
    Exit Sub

FailActivate:
    On Error Resume Next
    ShowStatus "Production form activation skipped: " & Err.Description
    On Error GoTo 0
End Sub

Private Sub UserForm_Resize()
    On Error Resume Next
    ResizeProductionLayout
    On Error GoTo 0
End Sub

Private Sub UserForm_Terminate()
    Set mOperatorWorkbook = Nothing
    Set mRunTreeCollapsed = Nothing
    Set mRunSplitOverrides = Nothing
    Set mRunBaseQtyByKey = Nothing
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

Public Function TestFilterAssignmentInventoryCount(ByVal values As Variant, ByVal filterText As String) As Long
    If Not mBuilt Then BuildLayout
    FillInventoryListFromArray values, filterText
    TestFilterAssignmentInventoryCount = mLstAssignInventory.ListCount
End Function

Private Sub BuildLayout()
    If mBuilt Then Exit Sub

    Me.Caption = "Production"
    Me.Width = PRODUCTION_BASE_WIDTH
    Me.Height = PRODUCTION_BASE_HEIGHT
    Me.ScrollBars = fmScrollBarsBoth
    Me.KeepScrollBarsVisible = fmScrollBarsNone
    Me.ScrollWidth = PRODUCTION_BASE_WIDTH - 20
    Me.ScrollHeight = PRODUCTION_BASE_HEIGHT - 35

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
    mPages.Pages(2).Caption = "Production Run - List"
    mPages.Pages(3).Caption = "Production Run - Tree"

    BuildRecipeBuilderPage mPages.Pages(0)
    BuildAssignmentPage mPages.Pages(1)
    BuildLoaderPage mPages.Pages(2)
    BuildRunTreePage mPages.Pages(3)

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
    ResizeProductionLayout
End Sub

Private Sub BuildRecipeBuilderPage(ByVal pg As MSForms.Page)
    AddLabel pg, "Saved Recipes", 12, 12, 180, 16
    Set mLstBuilderRecipes = AddList(pg, "lstBuilderRecipes", 12, 32, 320, 230, 3, "0 pt;130 pt;170 pt")
    AddLabel pg, "Recipe Name", 350, 12, 100, 16
    Set mTxtRecipeName = AddText(pg, "txtRecipeName", 350, 32, 240, 22)
    AddLabel pg, "Recipe ID", 610, 12, 80, 16
    Set mTxtRecipeId = AddText(pg, "txtRecipeId", 610, 32, 230, 22)
    AddLabel pg, "Row Budget", 610, 62, 90, 16
    Set mTxtRecipeRowBudget = AddText(pg, "txtRecipeRowBudget", 700, 58, 140, 22)
    AddLabel pg, "Description", 350, 62, 120, 16
    Set mTxtRecipeDescription = AddText(pg, "txtRecipeDescription", 350, 96, 490, 30)
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
    AddColumnHeaders pg, "LoaderRecipes", Array("", "Recipe", "Description"), 12, 32, RUN_LOADER_RECIPE_WIDTHS
    Set mLstLoaderRecipes = AddList(pg, "lstLoaderRecipes", 12, 50, 270, 147, 3, RUN_LOADER_RECIPE_WIDTHS)
    Set mBtnLoaderRefresh = AddButton(pg, "btnLoaderRefresh", "Refresh", 300, 32, 130, 24)
    Set mBtnLoaderLoad = AddButton(pg, "btnLoaderLoad", "Load Recipe", 300, 64, 130, 24)
    Set mBtnLoaderClear = AddButton(pg, "btnLoaderClear", "Clear Run", 300, 96, 130, 24)

    AddLabel pg, "Loaded Recipe Lines", 455, 12, 180, 16
    AddColumnHeaders pg, "LoaderLines", Array("Process", "", "I/O", "Ingredient", "%", "UOM", "Amount", ""), 455, 32, RUN_LOADER_LINE_WIDTHS
    Set mLstLoaderLines = AddList(pg, "lstLoaderLines", 455, 50, 575, 147, 8, RUN_LOADER_LINE_WIDTHS)

    AddLabel pg, "Acceptable Inventory For Run", 12, 215, 230, 16
    AddRunPaletteHeader pg, 12, 235
    Set mLstRunPalette = AddList(pg, "lstRunPalette", 12, 253, 650, 122, 10, RUN_PALETTE_WIDTHS)
    AddLabel pg, "Process", 680, 144, 70, 16
    Set mCmbRunProcess = AddCombo(pg, "cmbRunProcess", 790, 140, 200, 22)
    AddLabel pg, "Run Location", 680, 174, 90, 16
    Set mCmbRunLocation = AddCombo(pg, "cmbRunLocation", 790, 170, 200, 22)
    AddLabel pg, "% of Requirement", 680, 214, 110, 16
    Set mTxtPaletteSplit = AddText(pg, "txtPaletteSplit", 680, 234, 90, 22)
    AddLabel pg, "Qty", 790, 214, 45, 16
    Set mTxtPaletteQty = AddText(pg, "txtPaletteQty", 790, 234, 90, 22)
    Set mBtnRunApplyPalette = AddButton(pg, "btnRunApplyPalette", "Apply", 900, 233, 90, 24)

    AddLabel pg, "Production Output", 680, 272, 170, 16
    AddColumnHeaders pg, "ManagerOutput", Array("Process", "Output", "UOM", "Real", "Batch", "Recall", "ROW"), 680, 292, RUN_OUTPUT_WIDTHS
    Set mLstManagerOutput = AddList(pg, "lstManagerOutput", 680, 310, 350, 88, 7, RUN_OUTPUT_WIDTHS)
    AddLabel pg, "Real Output", 680, 410, 80, 16
    Set mTxtOutputReal = AddText(pg, "txtOutputReal", 760, 406, 80, 22)
    AddLabel pg, "Batch", 850, 410, 50, 16
    Set mTxtOutputBatch = AddText(pg, "txtOutputBatch", 900, 406, 80, 22)
    AddLabel pg, "Output Total", 680, 434, 80, 16
    Set mTxtOutputTotal = AddText(pg, "txtOutputTotal", 760, 430, 220, 22)
    mTxtOutputTotal.Locked = True
    Set mBtnManagerCheckIn = AddButton(pg, "btnManagerCheckIn", "Check In", 680, 462, 95, 24)
    Set mBtnManagerApplyOutput = AddButton(pg, "btnManagerApplyOutput", "Complete Run", 790, 462, 120, 24)

    AddLabel pg, "Inventory Check", 12, 395, 150, 16
    AddColumnHeaders pg, "ManagerCheck", Array("ROW", "Code", "Item", "UOM", "Used", "Total Inv"), 12, 415, RUN_CHECK_WIDTHS
    Set mLstManagerCheck = AddList(pg, "lstManagerCheck", 12, 433, 650, 87, 6, RUN_CHECK_WIDTHS)

    Set mBtnManagerRefresh = AddButton(pg, "btnManagerRefresh", "Refresh", 930, 462, 95, 24)
    Set mBtnManagerNext = AddButton(pg, "btnManagerNext", "Next Batch", 680, 506, 120, 26)
    Set mBtnManagerPrint = AddButton(pg, "btnManagerPrint", "Print Recall", 820, 506, 120, 26)
End Sub

Private Sub BuildRunTreePage(ByVal pg As MSForms.Page)
    AddLabel pg, "Recipe Ingredient / Inventory Choices", 12, 12, 260, 16
    AddLabel pg, "% of Requirement", 300, 10, 110, 16
    Set mTxtTreePaletteSplit = AddText(pg, "txtTreePaletteSplit", 410, 8, 70, 22)
    AddLabel pg, "Qty", 492, 10, 35, 16
    Set mTxtTreePaletteQty = AddText(pg, "txtTreePaletteQty", 530, 8, 80, 22)
    Set mBtnRunTreeApplyPalette = AddButton(pg, "btnRunTreeApplyPalette", "Apply", 622, 7, 70, 24)
    AddLabel pg, "Process", 704, 10, 55, 16
    Set mCmbTreeRunProcess = AddCombo(pg, "cmbTreeRunProcess", 760, 8, 100, 22)
    AddLabel pg, "Location", 870, 10, 58, 16
    Set mCmbTreeRunLocation = AddCombo(pg, "cmbTreeRunLocation", 930, 8, 100, 22)
    Set mBtnRunTreeExpandAll = AddButton(pg, "btnRunTreeExpandAll", "Expand", 850, 528, 70, 24)
    Set mBtnRunTreeCollapseAll = AddButton(pg, "btnRunTreeCollapseAll", "Collapse", 930, 528, 80, 24)
    AddRunPaletteHeader pg, 12, 42
    Set mLstRunTree = AddList(pg, "lstRunTree", 12, 60, 1018, 460, 10, "0 pt;0 pt;300 pt;42 pt;165 pt;58 pt;68 pt;48 pt;80 pt;110 pt")
    mLstRunTree.Font.Size = 11
End Sub

Private Sub AddRunPaletteHeader(ByVal pg As MSForms.Page, ByVal leftVal As Single, ByVal topVal As Single)
    AddLabel pg, "Ingredient", leftVal + 2, topVal, 135, 14
    AddLabel pg, "ROW", leftVal + 130, topVal, 35, 14
    AddLabel pg, "Inventory Item", leftVal + 168, topVal, 145, 14
    AddLabel pg, "% Req", leftVal + 315, topVal, 46, 14
    AddLabel pg, "Qty", leftVal + 365, topVal, 50, 14
    AddLabel pg, "UOM", leftVal + 425, topVal, 40, 14
    AddLabel pg, "Inv", leftVal + 465, topVal, 45, 14
    AddLabel pg, "Location", leftVal + 540, topVal, 70, 14
End Sub

Private Sub AddColumnHeaders(ByVal parent As Object, ByVal groupName As String, ByVal labels As Variant, _
                             ByVal leftVal As Single, ByVal topVal As Single, ByVal widths As String)
    Dim i As Long
    Dim idx As Long
    Dim x As Single
    Dim colWidth As Single
    Dim caption As String
    Dim lbl As MSForms.Label

    x = leftVal
    For i = LBound(labels) To UBound(labels)
        idx = i - LBound(labels)
        colWidth = ColumnWidthPoints(widths, idx, 60)
        caption = Trim$(CStr(labels(i)))
        If colWidth > 0 And caption <> "" Then
            Set lbl = parent.Controls.Add("Forms.Label.1", "hdr" & CleanControlName(groupName) & CStr(idx + 1), True)
            With lbl
                .Caption = caption
                .Left = x + 2
                .Top = topVal
                .Width = MaxDoubleForm(1, colWidth - 4)
                .Height = 14
                .Font.Bold = True
            End With
        End If
        x = x + colWidth
    Next i
End Sub

Private Sub PositionColumnHeaders(ByVal parent As Object, ByVal groupName As String, _
                                  ByVal leftVal As Single, ByVal topVal As Single, ByVal widths As String)
    Dim parts As Variant
    Dim i As Long
    Dim x As Single
    Dim colWidth As Single
    Dim lbl As MSForms.Label

    parts = Split(widths, ";")
    x = leftVal
    For i = LBound(parts) To UBound(parts)
        colWidth = ColumnWidthPoints(widths, i, 60)
        Set lbl = Nothing
        On Error Resume Next
        Set lbl = parent.Controls("hdr" & CleanControlName(groupName) & CStr(i + 1))
        On Error GoTo 0
        If Not lbl Is Nothing Then
            lbl.Left = x + 2
            lbl.Top = topVal
            lbl.Width = MaxDoubleForm(1, colWidth - 4)
        End If
        x = x + colWidth
    Next i
End Sub

Private Function ColumnWidthPoints(ByVal widths As String, ByVal zeroBasedIndex As Long, _
                                   Optional ByVal defaultWidth As Single = 60) As Single
    Dim parts As Variant
    Dim rawValue As String

    parts = Split(widths, ";")
    If zeroBasedIndex < LBound(parts) Or zeroBasedIndex > UBound(parts) Then
        ColumnWidthPoints = defaultWidth
        Exit Function
    End If
    rawValue = Replace$(LCase$(Trim$(CStr(parts(zeroBasedIndex)))), "pt", "")
    ColumnWidthPoints = CSng(Val(rawValue))
End Function

Private Sub ResizeProductionLayout()
    On Error GoTo CleanExit
    If Not mBuilt Then Exit Sub
    If mResizingLayout Then Exit Sub
    mResizingLayout = True

    Dim layoutW As Double
    Dim layoutH As Double
    layoutW = MaxDoubleForm(PRODUCTION_BASE_WIDTH - 20, Me.InsideWidth)
    layoutH = MaxDoubleForm(PRODUCTION_BASE_HEIGHT - 35, Me.InsideHeight)

    Me.ScrollWidth = layoutW
    Me.ScrollHeight = layoutH

    If Not mPages Is Nothing Then
        mPages.Move 12, 10, MaxDoubleForm(720, layoutW - 40), MaxDoubleForm(460, layoutH - 115)
    End If
    If Not mTxtStatus Is Nothing Then
        mTxtStatus.Move 12, layoutH - 84, MaxDoubleForm(360, layoutW - 210), 42
    End If
    If Not mBtnClose Is Nothing Then
        mBtnClose.Move layoutW - 170, layoutH - 84, 150, 42
    End If

    ResizeProductionPages

CleanExit:
    mResizingLayout = False
End Sub

Private Sub ResizeProductionPages()
    Dim pageW As Double
    Dim pageH As Double
    Dim runRightPanelLeft As Double

    If mPages Is Nothing Then Exit Sub
    pageW = MaxDoubleForm(700, mPages.Width - 20)
    pageH = MaxDoubleForm(420, mPages.Height - 45)
    runRightPanelLeft = 680

    If Not mLstBuilderRecipes Is Nothing Then mLstBuilderRecipes.Height = MaxDoubleForm(130, pageH - 290)
    If Not mLstBuilderLines Is Nothing Then
        mLstBuilderLines.Width = MaxDoubleForm(520, pageW - 40)
        mLstBuilderLines.Height = MaxDoubleForm(120, pageH - mLstBuilderLines.Top - 18)
    End If

    If Not mLstAssignRecipes Is Nothing Then mLstAssignRecipes.Height = MaxDoubleForm(115, pageH - 335)
    If Not mLstAssignIngredients Is Nothing Then mLstAssignIngredients.Width = MaxDoubleForm(300, pageW - mLstAssignIngredients.Left - 380)
    If Not mTxtInventorySearch Is Nothing Then mTxtInventorySearch.Width = MaxDoubleForm(150, pageW - 810)
    If Not mLstAssignInventory Is Nothing Then
        mLstAssignInventory.Width = MaxDoubleForm(360, (pageW - 58) / 2)
        mLstAssignInventory.Height = MaxDoubleForm(130, pageH - mLstAssignInventory.Top - 18)
    End If
    If Not mLstAssignAllowed Is Nothing And Not mLstAssignInventory Is Nothing Then
        mLstAssignAllowed.Left = mLstAssignInventory.Left + mLstAssignInventory.Width + 18
        mLstAssignAllowed.Width = MaxDoubleForm(320, pageW - mLstAssignAllowed.Left - 28)
        mLstAssignAllowed.Height = mLstAssignInventory.Height
    End If

    If Not mLstLoaderLines Is Nothing Then mLstLoaderLines.Width = MaxDoubleForm(420, pageW - mLstLoaderLines.Left - 30)
    If Not mLstRunPalette Is Nothing Then
        mLstRunPalette.Width = MaxDoubleForm(430, runRightPanelLeft - mLstRunPalette.Left - 28)
        mLstRunPalette.Height = MaxDoubleForm(90, (pageH - 330) / 2)
    End If
    If Not mLstManagerCheck Is Nothing And Not mLstRunPalette Is Nothing Then
        mLstManagerCheck.Top = mLstRunPalette.Top + mLstRunPalette.Height + 68
        mLstManagerCheck.Width = mLstRunPalette.Width
        mLstManagerCheck.Height = MaxDoubleForm(70, pageH - mLstManagerCheck.Top - 18)
        PositionColumnHeaders mPages.Pages(2), "ManagerCheck", mLstManagerCheck.Left, mLstManagerCheck.Top - 18, RUN_CHECK_WIDTHS
    End If
    If Not mLstManagerOutput Is Nothing Then mLstManagerOutput.Width = MaxDoubleForm(350, pageW - mLstManagerOutput.Left - 28)

    If Not mLstRunTree Is Nothing Then
        mLstRunTree.Width = MaxDoubleForm(520, pageW - 40)
        mLstRunTree.Height = MaxDoubleForm(260, pageH - mLstRunTree.Top - 18)
    End If
    If Not mBtnRunTreeCollapseAll Is Nothing Then mBtnRunTreeCollapseAll.Left = MaxDoubleForm(150, pageW - 85)
    If Not mBtnRunTreeExpandAll Is Nothing And Not mBtnRunTreeCollapseAll Is Nothing Then mBtnRunTreeExpandAll.Left = mBtnRunTreeCollapseAll.Left - 66
    If Not mCmbTreeRunLocation Is Nothing And Not mBtnRunTreeExpandAll Is Nothing Then mCmbTreeRunLocation.Left = mBtnRunTreeExpandAll.Left - 115
    If Not mBtnRunTreeApplyPalette Is Nothing And Not mCmbTreeRunLocation Is Nothing Then mBtnRunTreeApplyPalette.Left = mCmbTreeRunLocation.Left - 88
    If Not mTxtTreePaletteQty Is Nothing And Not mBtnRunTreeApplyPalette Is Nothing Then mTxtTreePaletteQty.Left = mBtnRunTreeApplyPalette.Left - 92
    If Not mTxtTreePaletteSplit Is Nothing And Not mTxtTreePaletteQty Is Nothing Then mTxtTreePaletteSplit.Left = mTxtTreePaletteQty.Left - 120
End Sub

Private Function MaxDoubleForm(ByVal leftValue As Double, ByVal rightValue As Double) As Double
    If leftValue >= rightValue Then
        MaxDoubleForm = leftValue
    Else
        MaxDoubleForm = rightValue
    End If
End Function

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
    mTxtRecipeRowBudget.Text = NormalizeRecipeRowBudgetText(FirstRowValue(lo, "ROW_BUDGET"))
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

Private Sub RefreshInventoryList(Optional ByVal forceReload As Boolean = False)
    If forceReload Or Not mInventoryCacheLoaded Then
        mInventoryRows = RunProduction1("LoadProductionInventoryPickerItems", "")
        mInventoryCacheLoaded = True
    End If
    FillInventoryListFromArray mInventoryRows, Trim$(mTxtInventorySearch.Text)
End Sub

Private Sub RefreshAllowedItems()
    FillListFromTable mLstAssignAllowed, ProductionTable(TABLE_ASSIGN_ITEM), _
        Array("ROW", "ITEMS", "UOM", "DESCRIPTION", "RECIPE_ID", "INGREDIENT_ID")
End Sub

Private Sub RefreshLoaderState()
    FillListFromTable mLstLoaderLines, ProductionTable(TABLE_LOADER_LINES), _
        Array("PROCESS", "DIAGRAM_ID", "INPUT/OUTPUT", "INGREDIENT", "PERCENT", "UOM", "AMOUNT NEEDED", "INGREDIENT_ID")
    RefreshRunProcessChoices
    If Not mLstLoaderOutput Is Nothing Then
        FillListFromTable mLstLoaderOutput, ProductionTable(TABLE_MANAGER_OUTPUT), _
            Array("PROCESS", "OUTPUT", "UOM", "REAL OUTPUT", "BATCH", "RECALL CODE", "ROW")
    End If
    RefreshRunPaletteState
End Sub

Private Sub RefreshManagerState()
    FillListFromTable mLstManagerOutput, ProductionTable(TABLE_MANAGER_OUTPUT), _
        Array("PROCESS", "OUTPUT", "UOM", "REAL OUTPUT", "BATCH", "RECALL CODE", "ROW")
    FillListFromTable mLstManagerCheck, ProductionTable(TABLE_MANAGER_CHECK), _
        Array("ROW", "ITEM_CODE", "ITEM", "UOM", "USED", "TOTAL INV")
    RefreshRunPaletteState
End Sub

Private Sub RefreshRunPaletteState()
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim choices As Variant
    Dim filterIngredientId As String
    Dim filterIngredientName As String
    Dim filterProcess As String

    If mLstRunPalette Is Nothing Then Exit Sub
    mLstRunPalette.Clear
    If Not mLstRunTree Is Nothing Then mLstRunTree.Clear
    GetSelectedRunIngredientFilter filterIngredientId, filterIngredientName
    filterProcess = ActiveRunProcess()
    EnsureRunInventoryCache
    RefreshRunLocationChoices

    choices = RunProduction1("LoadProductionRunIngredientChoices", NzStr(RunProduction0("GetCurrentProductionRunRecipeId")))
    If Not IsEmpty(choices) Then
        AddRunChoiceRows choices, filterIngredientId, filterIngredientName, filterProcess
        BuildRunTreeFromPaletteList
        Exit Sub
    End If

    Set ws = RunProductionObject0("GetProductionSheet")
    If ws Is Nothing Then Exit Sub

    For Each lo In ws.ListObjects
        If IsRunPaletteTable(lo) Then AddRunPaletteRows lo, filterIngredientId, filterIngredientName, filterProcess
    Next lo
    BuildRunTreeFromPaletteList
End Sub

Private Sub AddRunChoiceRows(ByVal values As Variant, Optional ByVal filterIngredientId As String = "", _
                             Optional ByVal filterIngredientName As String = "", Optional ByVal filterProcess As String = "")
    Dim r As Long
    Dim listRow As Long
    Dim rowVal As String
    Dim itemVal As String
    Dim uomVal As String
    Dim locVal As String
    Dim invVal As String
    Dim ingredientId As String
    Dim ingredientName As String
    Dim processVal As String

    If IsEmpty(values) Then Exit Sub
    If Not IsArray(values) Then Exit Sub
    EnsureRunInventoryCache
    For r = LBound(values, 1) To UBound(values, 1)
        ingredientId = NzStr(values(r, 2))
        ingredientName = NzStr(values(r, 3))
        processVal = NzStr(values(r, 1))
        If Not RunProcessMatchesFilter(processVal, filterProcess) Then GoTo NextRow
        If Not RunIngredientMatchesFilter(ingredientId, ingredientName, filterIngredientId, filterIngredientName) Then GoTo NextRow

        rowVal = NzStr(values(r, 4))
        itemVal = NzStr(values(r, 5))
        uomVal = NzStr(values(r, 8))
        locVal = NzStr(values(r, 9))
        HydrateRunInventoryDisplay rowVal, itemVal, uomVal, invVal, locVal

        mLstRunPalette.AddItem processVal
        listRow = mLstRunPalette.ListCount - 1
        mLstRunPalette.List(listRow, 1) = ingredientId
        mLstRunPalette.List(listRow, 2) = ingredientName
        mLstRunPalette.List(listRow, 3) = rowVal
        mLstRunPalette.List(listRow, 4) = itemVal
        mLstRunPalette.List(listRow, 5) = NzStr(values(r, 6))
        mLstRunPalette.List(listRow, 6) = NzStr(values(r, 7))
        mLstRunPalette.List(listRow, 7) = uomVal
        mLstRunPalette.List(listRow, 8) = invVal
        mLstRunPalette.List(listRow, 9) = locVal
        StoreRunBaseQty mLstRunPalette, listRow, NzStr(values(r, 10))
        ApplyRunAllocationOverride mLstRunPalette, listRow
NextRow:
    Next r
End Sub

Private Function IsRunPaletteTable(ByVal lo As ListObject) As Boolean
    If lo Is Nothing Then Exit Function
    If StrComp(lo.Name, TABLE_MANAGER_PALETTE, vbTextCompare) = 0 Then
        If lo.Range.Row >= 500000 Then Exit Function
    End If
    IsRunPaletteTable = ProductionColumnIndex(lo, "ITEM") > 0 _
        And ProductionColumnIndex(lo, "QUANTITY") > 0 _
        And ProductionColumnIndex(lo, "PROCESS") > 0 _
        And ProductionColumnIndex(lo, "ROW") > 0 _
        And ProductionColumnIndex(lo, "INPUT/OUTPUT") > 0
End Function

Private Sub AddRunPaletteRows(ByVal lo As ListObject, Optional ByVal filterIngredientId As String = "", _
                              Optional ByVal filterIngredientName As String = "", Optional ByVal filterProcess As String = "")
    Dim arr As Variant
    Dim r As Long
    Dim listRow As Long
    Dim rowVal As String
    Dim ingredientVal As String
    Dim ingredientId As String
    Dim processVal As String
    Dim itemVal As String
    Dim splitVal As String
    Dim qtyVal As String
    Dim uomVal As String
    Dim locVal As String
    Dim invVal As String
    Dim baseQty As String
    Dim ioVal As String

    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    EnsureRunInventoryCache
    arr = lo.DataBodyRange.Value
    For r = 1 To UBound(arr, 1)
        ingredientVal = CellText(arr, r, lo, "INGREDIENT")
        ingredientId = CellText(arr, r, lo, "INGREDIENT_ID")
        processVal = CellText(arr, r, lo, "PROCESS")
        If Not RunProcessMatchesFilter(processVal, filterProcess) Then GoTo NextRow
        ioVal = UCase$(Trim$(CellText(arr, r, lo, "INPUT/OUTPUT")))
        If ioVal <> "" And ioVal <> "USED" Then GoTo NextRow
        If Not RunIngredientMatchesFilter(ingredientId, IngredientDisplayName(ingredientVal, processVal), filterIngredientId, filterIngredientName) Then GoTo NextRow

        rowVal = CellText(arr, r, lo, "ROW")
        itemVal = CellText(arr, r, lo, "ITEM")
        splitVal = CellText(arr, r, lo, "SPLIT %")
        qtyVal = CellText(arr, r, lo, "QUANTITY")
        uomVal = CellText(arr, r, lo, "UOM")
        locVal = CellText(arr, r, lo, "LOCATION")
        baseQty = CellText(arr, r, lo, "BASE QUANTITY")
        HydrateRunInventoryDisplay rowVal, itemVal, uomVal, invVal, locVal

        mLstRunPalette.AddItem lo.Name
        listRow = mLstRunPalette.ListCount - 1
        mLstRunPalette.List(listRow, 1) = CStr(r)
        mLstRunPalette.List(listRow, 2) = IngredientDisplayName(ingredientVal, processVal)
        mLstRunPalette.List(listRow, 3) = rowVal
        mLstRunPalette.List(listRow, 4) = itemVal
        mLstRunPalette.List(listRow, 5) = splitVal
        mLstRunPalette.List(listRow, 6) = qtyVal
        mLstRunPalette.List(listRow, 7) = uomVal
        mLstRunPalette.List(listRow, 8) = invVal
        mLstRunPalette.List(listRow, 9) = locVal
        StoreRunBaseQty mLstRunPalette, listRow, baseQty
        ApplyRunAllocationOverride mLstRunPalette, listRow
NextRow:
    Next r
End Sub

Private Sub EnsureRunSplitOverrides()
    If mRunSplitOverrides Is Nothing Then Set mRunSplitOverrides = CreateObject("Scripting.Dictionary")
End Sub

Private Sub EnsureRunBaseQtyMap()
    If mRunBaseQtyByKey Is Nothing Then Set mRunBaseQtyByKey = CreateObject("Scripting.Dictionary")
End Sub

Private Function RunAllocationKeyFromList(ByVal lst As MSForms.ListBox, ByVal rowIndex As Long) As String
    If lst Is Nothing Then Exit Function
    If rowIndex < 0 Then Exit Function
    RunAllocationKeyFromList = Trim$(NzStr(lst.List(rowIndex, 0))) & "|" & _
                               Trim$(NzStr(lst.List(rowIndex, 1))) & "|" & _
                               Trim$(NzStr(lst.List(rowIndex, 3))) & "|" & _
                               Trim$(NzStr(lst.List(rowIndex, 4)))
End Function

Private Sub StoreRunBaseQty(ByVal lst As MSForms.ListBox, ByVal rowIndex As Long, ByVal baseQtyText As String)
    Dim key As String
    key = RunAllocationKeyFromList(lst, rowIndex)
    If key = "" Then Exit Sub
    EnsureRunBaseQtyMap
    mRunBaseQtyByKey(key) = baseQtyText
End Sub

Private Function RunBaseQtyFromList(ByVal lst As MSForms.ListBox, ByVal rowIndex As Long) As Double
    Dim key As String
    Dim baseText As String

    If lst Is Nothing Then Exit Function
    If rowIndex < 0 Then Exit Function
    If mRunBaseQtyByKey Is Nothing Then Exit Function
    key = RunAllocationKeyFromList(lst, rowIndex)
    If key = "" Then Exit Function
    If Not mRunBaseQtyByKey.Exists(key) Then Exit Function
    baseText = NzStr(mRunBaseQtyByKey(key))
    If IsNumeric(baseText) Then RunBaseQtyFromList = CDbl(baseText)
End Function

Private Function RunAllocationGroupKeyFromList(ByVal lst As MSForms.ListBox, ByVal rowIndex As Long) As String
    Dim r As Long

    If lst Is Nothing Then Exit Function
    If rowIndex < 0 Then Exit Function
    If Not mLstRunTree Is Nothing Then
        If lst Is mLstRunTree Then
            If IsRunTreeParentRow(lst, rowIndex) Then
                RunAllocationGroupKeyFromList = Trim$(NzStr(lst.List(rowIndex, 1)))
                Exit Function
            End If
            For r = rowIndex To 0 Step -1
                If IsRunTreeParentRow(lst, r) Then
                    RunAllocationGroupKeyFromList = Trim$(NzStr(lst.List(r, 1)))
                    Exit Function
                End If
            Next r
        End If
    End If
    RunAllocationGroupKeyFromList = Trim$(NzStr(lst.List(rowIndex, 1)))
    If RunAllocationGroupKeyFromList = "" Or IsNumeric(RunAllocationGroupKeyFromList) Then
        RunAllocationGroupKeyFromList = Trim$(NzStr(lst.List(rowIndex, 2)))
    End If
End Function

Private Function RunAllocationPercentForGroup(ByVal groupKey As String, Optional ByVal excludeAllocationKey As String = "") As Double
    Dim i As Long
    Dim splitText As String

    If mLstRunPalette Is Nothing Then Exit Function
    groupKey = Trim$(groupKey)
    For i = 0 To mLstRunPalette.ListCount - 1
        If StrComp(RunAllocationGroupKeyFromList(mLstRunPalette, i), groupKey, vbTextCompare) = 0 Then
            If excludeAllocationKey = "" Or RunAllocationKeyFromList(mLstRunPalette, i) <> excludeAllocationKey Then
                splitText = Trim$(NzStr(mLstRunPalette.List(i, 5)))
                If IsNumeric(splitText) Then RunAllocationPercentForGroup = RunAllocationPercentForGroup + CDbl(splitText)
            End If
        End If
    Next i
End Function

Private Function RunAllocationStatusText(ByVal groupKey As String) As String
    Dim totalPct As Double

    totalPct = RunAllocationPercentForGroup(groupKey)
    If totalPct >= 99.999 And totalPct <= 100.001 Then
        RunAllocationStatusText = "FILLED 100%"
    ElseIf totalPct > 0 Then
        RunAllocationStatusText = FormatRunNumber(totalPct) & "% filled"
    Else
        RunAllocationStatusText = "0% filled"
    End If
End Function

Private Sub StoreRunAllocationOverride(ByVal lst As MSForms.ListBox, ByVal rowIndex As Long, _
                                       ByVal splitText As String, ByVal qtyText As String)
    Dim key As String
    key = RunAllocationKeyFromList(lst, rowIndex)
    If key = "" Then Exit Sub
    EnsureRunSplitOverrides
    mRunSplitOverrides(key) = Array(splitText, qtyText)
End Sub

Private Sub ApplyRunAllocationOverride(ByVal lst As MSForms.ListBox, ByVal rowIndex As Long)
    Dim key As String
    Dim values As Variant

    If lst Is Nothing Then Exit Sub
    If rowIndex < 0 Then Exit Sub
    If mRunSplitOverrides Is Nothing Then Exit Sub
    key = RunAllocationKeyFromList(lst, rowIndex)
    If key = "" Then Exit Sub
    If Not mRunSplitOverrides.Exists(key) Then Exit Sub
    values = mRunSplitOverrides(key)
    lst.List(rowIndex, 5) = NzStr(values(0))
    lst.List(rowIndex, 6) = NzStr(values(1))
End Sub

Private Sub GetSelectedRunIngredientFilter(ByRef ingredientId As String, ByRef ingredientName As String)
    Dim idx As Long

    ingredientId = ""
    ingredientName = ""
    If mLstLoaderLines Is Nothing Then Exit Sub
    idx = mLstLoaderLines.ListIndex
    If idx < 0 Then Exit Sub
    ingredientName = NzStr(mLstLoaderLines.List(idx, 3))
    ingredientId = NzStr(mLstLoaderLines.List(idx, 7))
End Sub

Private Function RunIngredientMatchesFilter(ByVal rowIngredientId As String, ByVal rowIngredientName As String, _
                                            ByVal filterIngredientId As String, ByVal filterIngredientName As String) As Boolean
    rowIngredientId = Trim$(rowIngredientId)
    rowIngredientName = Trim$(rowIngredientName)
    filterIngredientId = Trim$(filterIngredientId)
    filterIngredientName = Trim$(filterIngredientName)

    If filterIngredientId = "" And filterIngredientName = "" Then
        RunIngredientMatchesFilter = True
    ElseIf filterIngredientId <> "" And rowIngredientId <> "" Then
        RunIngredientMatchesFilter = (StrComp(rowIngredientId, filterIngredientId, vbTextCompare) = 0)
    ElseIf filterIngredientName <> "" Then
        RunIngredientMatchesFilter = (StrComp(rowIngredientName, filterIngredientName, vbTextCompare) = 0)
    End If
End Function

Private Function RunProcessMatchesFilter(ByVal rowProcess As String, ByVal filterProcess As String) As Boolean
    rowProcess = Trim$(rowProcess)
    filterProcess = Trim$(filterProcess)
    If filterProcess = "" Then
        RunProcessMatchesFilter = True
    Else
        RunProcessMatchesFilter = (StrComp(rowProcess, filterProcess, vbTextCompare) = 0)
    End If
End Function

Private Sub BuildRunTreeFromPaletteList()
    Dim i As Long
    Dim key As String
    Dim lastKey As String
    Dim parentText As String
    Dim childText As String
    Dim rowIndex As Long
    Dim collapsed As Boolean

    If mLstRunTree Is Nothing Or mLstRunPalette Is Nothing Then Exit Sub
    EnsureRunTreeState
    mLstRunTree.Clear
    For i = 0 To mLstRunPalette.ListCount - 1
        key = RunTreeGroupKey(mLstRunPalette, i)
        If key <> lastKey Then
            parentText = NzStr(mLstRunPalette.List(i, 2))
            If parentText = "" Then parentText = "Ingredient"
            collapsed = RunTreeGroupCollapsed(key)
            mLstRunTree.AddItem RUN_TREE_PARENT_MARKER
            rowIndex = mLstRunTree.ListCount - 1
            mLstRunTree.List(rowIndex, 1) = key
            mLstRunTree.List(rowIndex, 2) = RunTreeParentCaption(parentText, collapsed, key)
            lastKey = key
        End If

        If collapsed Then GoTo NextPaletteRow
        childText = "  " & NzStr(mLstRunPalette.List(i, 4))
        If Trim$(childText) = "" Then childText = "  ROW " & NzStr(mLstRunPalette.List(i, 3))
        mLstRunTree.AddItem NzStr(mLstRunPalette.List(i, 0))
        rowIndex = mLstRunTree.ListCount - 1
        CopyRunPaletteListRow mLstRunPalette, i, mLstRunTree, rowIndex
        mLstRunTree.List(rowIndex, 2) = childText
NextPaletteRow:
    Next i
End Sub

Private Sub EnsureRunTreeState()
    If mRunTreeCollapsed Is Nothing Then Set mRunTreeCollapsed = CreateObject("Scripting.Dictionary")
End Sub

Private Function RunTreeGroupKey(ByVal lst As MSForms.ListBox, ByVal rowIndex As Long) As String
    RunTreeGroupKey = RunAllocationGroupKeyFromList(lst, rowIndex)
End Function

Private Function RunTreeGroupCollapsed(ByVal groupKey As String) As Boolean
    EnsureRunTreeState
    If groupKey = "" Then Exit Function
    If mRunTreeCollapsed.Exists(groupKey) Then RunTreeGroupCollapsed = CBool(mRunTreeCollapsed(groupKey))
End Function

Private Function RunTreeParentCaption(ByVal parentText As String, ByVal collapsed As Boolean, ByVal groupKey As String) As String
    Dim statusText As String
    statusText = RunAllocationStatusText(groupKey)
    If collapsed Then
        RunTreeParentCaption = "[ SHOW CHOICES ]  " & parentText & "  --  " & statusText
    Else
        RunTreeParentCaption = "[ HIDE CHOICES ]  " & parentText & "  --  " & statusText
    End If
End Function

Private Function IsRunTreeParentRow(ByVal lst As MSForms.ListBox, ByVal rowIndex As Long) As Boolean
    If lst Is Nothing Then Exit Function
    If rowIndex < 0 Then Exit Function
    IsRunTreeParentRow = (NzStr(lst.List(rowIndex, 0)) = RUN_TREE_PARENT_MARKER)
End Function

Private Sub ToggleSelectedRunTreeParent()
    Dim idx As Long
    Dim key As String

    If mLstRunTree Is Nothing Then Exit Sub
    idx = mLstRunTree.ListIndex
    If Not IsRunTreeParentRow(mLstRunTree, idx) Then Exit Sub
    EnsureRunTreeState
    key = NzStr(mLstRunTree.List(idx, 1))
    If key = "" Then Exit Sub
    mRunTreeCollapsed(key) = Not RunTreeGroupCollapsed(key)
    BuildRunTreeFromPaletteList
    If RunTreeGroupCollapsed(key) Then
        ShowStatus "Ingredient choices hidden."
    Else
        ShowStatus "Ingredient choices shown."
    End If
End Sub

Private Sub SetAllRunTreeGroupsCollapsed(ByVal collapsed As Boolean)
    Dim i As Long
    Dim key As String

    If mLstRunPalette Is Nothing Then Exit Sub
    EnsureRunTreeState
    If Not collapsed Then
        mRunTreeCollapsed.RemoveAll
    Else
        For i = 0 To mLstRunPalette.ListCount - 1
            key = RunTreeGroupKey(mLstRunPalette, i)
            If key <> "" Then mRunTreeCollapsed(key) = True
        Next i
    End If
    BuildRunTreeFromPaletteList
End Sub

Private Sub CopyRunPaletteListRow(ByVal sourceList As MSForms.ListBox, ByVal sourceRow As Long, _
                                  ByVal targetList As MSForms.ListBox, ByVal targetRow As Long)
    Dim c As Long
    For c = 0 To sourceList.ColumnCount - 1
        If c < targetList.ColumnCount Then targetList.List(targetRow, c) = NzStr(sourceList.List(sourceRow, c))
    Next c
End Sub

Private Function IngredientDisplayName(ByVal ingredientVal As String, ByVal processVal As String) As String
    IngredientDisplayName = Trim$(ingredientVal)
    If IngredientDisplayName = "" Then IngredientDisplayName = Trim$(processVal)
End Function

Private Sub EnsureRunInventoryCache()
    If Not mInventoryCacheLoaded Then
        mInventoryRows = RunProduction1("LoadProductionInventoryPickerItems", "")
        mInventoryCacheLoaded = True
    End If
End Sub

Private Sub RefreshRunLocationChoices()
    Dim dict As Object
    Dim r As Long
    Dim locVal As String
    Dim selectedLoc As String
    Dim wasLoading As Boolean

    If mCmbRunLocation Is Nothing And mCmbTreeRunLocation Is Nothing Then Exit Sub
    If IsEmpty(mInventoryRows) Then Exit Sub
    If Not IsArray(mInventoryRows) Then Exit Sub

    selectedLoc = ActiveRunLocation()
    Set dict = CreateObject("Scripting.Dictionary")
    For r = LBound(mInventoryRows, 1) To UBound(mInventoryRows, 1)
        locVal = Trim$(NzStr(mInventoryRows(r, 5)))
        If locVal <> "" Then
            If Not dict.Exists(LCase$(locVal)) Then dict.Add LCase$(locVal), locVal
        End If
    Next r

    wasLoading = mLoading
    mLoading = True
    PopulateRunLocationCombo mCmbRunLocation, dict, selectedLoc
    PopulateRunLocationCombo mCmbTreeRunLocation, dict, selectedLoc
    mLoading = wasLoading
End Sub

Private Sub RefreshRunProcessChoices()
    Dim dict As Object
    Dim i As Long
    Dim procVal As String
    Dim selectedProcess As String
    Dim wasLoading As Boolean

    If mCmbRunProcess Is Nothing And mCmbTreeRunProcess Is Nothing Then Exit Sub
    If mLstLoaderLines Is Nothing Then Exit Sub

    selectedProcess = ActiveRunProcess()
    Set dict = CreateObject("Scripting.Dictionary")
    For i = 0 To mLstLoaderLines.ListCount - 1
        procVal = Trim$(NzStr(mLstLoaderLines.List(i, 0)))
        If procVal <> "" Then
            If Not dict.Exists(LCase$(procVal)) Then dict.Add LCase$(procVal), procVal
        End If
    Next i

    wasLoading = mLoading
    mLoading = True
    PopulateRunProcessCombo mCmbRunProcess, dict, selectedProcess
    PopulateRunProcessCombo mCmbTreeRunProcess, dict, selectedProcess
    mLoading = wasLoading
End Sub

Private Sub PopulateRunProcessCombo(ByVal cmb As MSForms.ComboBox, ByVal dict As Object, ByVal selectedProcess As String)
    Dim key As Variant
    Dim i As Long

    If cmb Is Nothing Then Exit Sub
    cmb.Clear
    cmb.AddItem ""
    If Not dict Is Nothing Then
        For Each key In dict.Keys
            cmb.AddItem CStr(dict(key))
        Next key
    End If
    cmb.ListIndex = 0
    selectedProcess = Trim$(selectedProcess)
    If selectedProcess <> "" Then
        For i = 0 To cmb.ListCount - 1
            If StrComp(NzStr(cmb.List(i)), selectedProcess, vbTextCompare) = 0 Then
                cmb.ListIndex = i
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub PopulateRunLocationCombo(ByVal cmb As MSForms.ComboBox, ByVal dict As Object, ByVal selectedLoc As String)
    Dim key As Variant
    Dim i As Long

    If cmb Is Nothing Then Exit Sub
    cmb.Clear
    cmb.AddItem ""
    If Not dict Is Nothing Then
        For Each key In dict.Keys
            cmb.AddItem CStr(dict(key))
        Next key
    End If
    cmb.ListIndex = 0
    selectedLoc = Trim$(selectedLoc)
    If selectedLoc <> "" Then
        For i = 0 To cmb.ListCount - 1
            If StrComp(NzStr(cmb.List(i)), selectedLoc, vbTextCompare) = 0 Then
                cmb.ListIndex = i
                Exit For
            End If
        Next i
    End If
End Sub

Private Function ActiveRunLocation() As String
    If Not mPages Is Nothing Then
        If mPages.Value = 3 Then
            ActiveRunLocation = ComboText(mCmbTreeRunLocation)
            If ActiveRunLocation <> "" Then Exit Function
        End If
    End If
    ActiveRunLocation = ComboText(mCmbRunLocation)
    If ActiveRunLocation = "" Then ActiveRunLocation = ComboText(mCmbTreeRunLocation)
End Function

Private Function ActiveRunProcess() As String
    If Not mPages Is Nothing Then
        If mPages.Value = 3 Then
            ActiveRunProcess = ComboText(mCmbTreeRunProcess)
            If ActiveRunProcess <> "" Then Exit Function
        End If
    End If
    ActiveRunProcess = ComboText(mCmbRunProcess)
    If ActiveRunProcess = "" Then ActiveRunProcess = ComboText(mCmbTreeRunProcess)
End Function

Private Function ComboText(ByVal cmb As MSForms.ComboBox) As String
    If cmb Is Nothing Then Exit Function
    If cmb.ListIndex < 0 Then Exit Function
    ComboText = Trim$(NzStr(cmb.Value))
End Function

Private Function RunLocationChoiceRequired() As Boolean
    If Not mCmbRunLocation Is Nothing Then
        If mCmbRunLocation.ListCount > 1 Then
            RunLocationChoiceRequired = True
            Exit Function
        End If
    End If
    If Not mCmbTreeRunLocation Is Nothing Then
        If mCmbTreeRunLocation.ListCount > 1 Then RunLocationChoiceRequired = True
    End If
End Function

Private Sub SyncRunLocationCombo(ByVal sourceCombo As MSForms.ComboBox, ByVal targetCombo As MSForms.ComboBox)
    Dim i As Long
    Dim locVal As String
    Dim wasLoading As Boolean

    If mLoading Then Exit Sub
    If sourceCombo Is Nothing Or targetCombo Is Nothing Then Exit Sub
    locVal = ComboText(sourceCombo)
    wasLoading = mLoading
    mLoading = True
    targetCombo.ListIndex = 0
    For i = 0 To targetCombo.ListCount - 1
        If StrComp(NzStr(targetCombo.List(i)), locVal, vbTextCompare) = 0 Then
            targetCombo.ListIndex = i
            Exit For
        End If
    Next i
    mLoading = wasLoading
End Sub

Private Sub SyncRunProcessCombo(ByVal sourceCombo As MSForms.ComboBox, ByVal targetCombo As MSForms.ComboBox)
    Dim i As Long
    Dim procVal As String
    Dim wasLoading As Boolean

    If mLoading Then Exit Sub
    If sourceCombo Is Nothing Or targetCombo Is Nothing Then Exit Sub
    procVal = ComboText(sourceCombo)
    wasLoading = mLoading
    mLoading = True
    targetCombo.ListIndex = 0
    For i = 0 To targetCombo.ListCount - 1
        If StrComp(NzStr(targetCombo.List(i)), procVal, vbTextCompare) = 0 Then
            targetCombo.ListIndex = i
            Exit For
        End If
    Next i
    mLoading = wasLoading
End Sub

Private Sub HydrateRunInventoryDisplay(ByVal rowVal As String, ByRef itemVal As String, _
                                       ByRef uomVal As String, ByRef invVal As String, _
                                       ByRef locVal As String)
    Dim r As Long
    Dim rowKey As String
    Dim totalVal As String
    Dim rawLoc As String

    If IsEmpty(mInventoryRows) Then Exit Sub
    If Not IsArray(mInventoryRows) Then Exit Sub

    rowKey = NormalizeRunRowKey(rowVal)
    If rowKey = "" Then Exit Sub
    For r = LBound(mInventoryRows, 1) To UBound(mInventoryRows, 1)
        If NormalizeRunRowKey(NzStr(mInventoryRows(r, 1))) = rowKey Then
            If Trim$(itemVal) = "" Then itemVal = NzStr(mInventoryRows(r, 2))
            If Trim$(uomVal) = "" Then uomVal = NzStr(mInventoryRows(r, 3))
            totalVal = NzStr(mInventoryRows(r, 4))
            rawLoc = NzStr(mInventoryRows(r, 5))
            invVal = RunInventoryAvailableDisplay(totalVal, uomVal)
            locVal = rawLoc
            Exit For
        End If
    Next r
End Sub

Private Function RunInventoryAvailableDisplay(ByVal totalVal As String, ByVal uomVal As String) As String
    totalVal = Trim$(totalVal)
    uomVal = Trim$(uomVal)
    If StrComp(totalVal, "utility", vbTextCompare) = 0 Then
        RunInventoryAvailableDisplay = "utility"
        Exit Function
    End If
    If totalVal <> "" Then
        RunInventoryAvailableDisplay = totalVal
        If uomVal <> "" Then RunInventoryAvailableDisplay = RunInventoryAvailableDisplay & " " & uomVal
    End If
End Function

Private Function NormalizeRunRowKey(ByVal value As String) As String
    NormalizeRunRowKey = Trim$(value)
    If IsNumeric(NormalizeRunRowKey) Then NormalizeRunRowKey = CStr(CLng(Val(NormalizeRunRowKey)))
End Function

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
        If Not TableArrayRowHasAnyValue(arr, r, lo, headers) Then GoTo NextTableRow
        lst.AddItem CellText(arr, r, lo, CStr(headers(LBound(headers))))
        For c = LBound(headers) + 1 To UBound(headers)
            colIdx = c - LBound(headers)
            If colIdx < lst.ColumnCount Then lst.List(lst.ListCount - 1, colIdx) = CellText(arr, r, lo, CStr(headers(c)))
        Next c
NextTableRow:
    Next r
End Sub

Private Function TableArrayRowHasAnyValue(ByVal arr As Variant, ByVal rowIndex As Long, ByVal lo As ListObject, ByVal headers As Variant) As Boolean
    Dim c As Long

    For c = LBound(headers) To UBound(headers)
        If Trim$(CellText(arr, rowIndex, lo, CStr(headers(c)))) <> "" Then
            TableArrayRowHasAnyValue = True
            Exit Function
        End If
    Next c
End Function

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

Private Sub FillInventoryListFromArray(ByVal values As Variant, Optional ByVal filterText As String = "")
    Dim r As Long
    Dim shown As Long
    Dim normalizedFilter As String

    mLstAssignInventory.Clear
    If IsEmpty(values) Then Exit Sub
    If Not IsArray(values) Then Exit Sub

    normalizedFilter = NormalizeInventorySearch(filterText)
    For r = LBound(values, 1) To UBound(values, 1)
        If Not InventoryRowMatchesSearch(values, r, normalizedFilter) Then GoTo NextInventoryRow
        shown = shown + 1
        mLstAssignInventory.AddItem NzStr(values(r, 1))
        If mLstAssignInventory.ColumnCount > 1 Then mLstAssignInventory.List(mLstAssignInventory.ListCount - 1, 1) = NzStr(values(r, 2))
        If mLstAssignInventory.ColumnCount > 2 Then mLstAssignInventory.List(mLstAssignInventory.ListCount - 1, 2) = NzStr(values(r, 3))
        If mLstAssignInventory.ColumnCount > 3 Then mLstAssignInventory.List(mLstAssignInventory.ListCount - 1, 3) = NzStr(values(r, 4))
        If mLstAssignInventory.ColumnCount > 4 Then mLstAssignInventory.List(mLstAssignInventory.ListCount - 1, 4) = NzStr(values(r, 5))
        If mLstAssignInventory.ColumnCount > 5 Then mLstAssignInventory.List(mLstAssignInventory.ListCount - 1, 5) = NzStr(values(r, 6))
        If mLstAssignInventory.ColumnCount > 6 Then mLstAssignInventory.List(mLstAssignInventory.ListCount - 1, 6) = NzStr(values(r, 7))
        If shown >= ASSIGN_INVENTORY_MAX_VISIBLE Then Exit For
NextInventoryRow:
    Next r
End Sub

Private Function InventoryRowMatchesSearch(ByVal values As Variant, ByVal rowIndex As Long, ByVal normalizedFilter As String) As Boolean
    Dim haystack As String

    If normalizedFilter = "" Then
        InventoryRowMatchesSearch = True
        Exit Function
    End If

    haystack = NormalizeInventorySearch(NzStr(values(rowIndex, 1)) & " " & _
        NzStr(values(rowIndex, 2)) & " " & NzStr(values(rowIndex, 3)) & " " & _
        NzStr(values(rowIndex, 5)) & " " & NzStr(values(rowIndex, 6)) & " " & _
        NzStr(values(rowIndex, 7)))
    InventoryRowMatchesSearch = (InStr(1, haystack, normalizedFilter, vbTextCompare) > 0)
End Function

Private Function NormalizeInventorySearch(ByVal filterText As String) As String
    Dim textOut As String

    textOut = Trim$(filterText)
    If textOut = "" Then Exit Function
    textOut = Replace$(textOut, vbCr, " ")
    textOut = Replace$(textOut, vbLf, " ")
    textOut = Replace$(textOut, vbTab, " ")
    Do While InStr(textOut, "  ") > 0
        textOut = Replace$(textOut, "  ", " ")
    Loop
    NormalizeInventorySearch = LCase$(textOut)
End Function

Private Sub ResetInventoryCache()
    mInventoryRows = Empty
    mInventoryCacheLoaded = False
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
    mTxtRecipeRowBudget.Text = NormalizeRecipeRowBudgetText(mTxtRecipeRowBudget.Text)
    SetFirstRowValue lo, "RECIPE_NAME", mTxtRecipeName.Text
    SetFirstRowValue lo, "RECIPE_ID", mTxtRecipeId.Text
    SetFirstRowValue lo, "DESCRIPTION", mTxtRecipeDescription.Text
    SetFirstRowValue lo, "ROW_BUDGET", CLng(Val(mTxtRecipeRowBudget.Text))
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
    mTxtRecipeRowBudget.Text = CStr(PRODUCTION_DEFAULT_ROW_BUDGET)
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
    Set lr = lo.ListRows.Add(AlwaysInsert:=False)
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
    If Not EnsureTableRows(lo, rowIndex) Then Exit Sub
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
    If Not mRunTreeCollapsed Is Nothing Then mRunTreeCollapsed.RemoveAll
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
    UpdateOutputRunningTotalDisplay
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
    UpdateOutputRunningTotalDisplay
    ShowStatus "Production output row updated."
End Sub

Private Sub UpdateOutputRunningTotalDisplay()
    Dim idx As Long
    Dim rowVal As String
    Dim procVal As String
    Dim outputVal As String
    Dim uomVal As String
    Dim batchVal As String
    Dim totalQty As Double
    Dim currentQty As Double
    Dim loggedCount As Long

    If mTxtOutputTotal Is Nothing Then Exit Sub
    mTxtOutputTotal.Text = ""
    If mLstManagerOutput Is Nothing Then Exit Sub
    idx = mLstManagerOutput.ListIndex
    If idx < 0 Then Exit Sub

    procVal = Trim$(NzStr(mLstManagerOutput.List(idx, 0)))
    outputVal = Trim$(NzStr(mLstManagerOutput.List(idx, 1)))
    uomVal = Trim$(NzStr(mLstManagerOutput.List(idx, 2)))
    batchVal = Trim$(NzStr(mLstManagerOutput.List(idx, 4)))
    rowVal = Trim$(NzStr(mLstManagerOutput.List(idx, 6)))

    totalQty = LoggedOutputTotal(rowVal, procVal, outputVal, loggedCount)
    If Not mTxtOutputReal Is Nothing Then
        If IsNumeric(mTxtOutputReal.Text) Then
            currentQty = CDbl(mTxtOutputReal.Text)
            If currentQty > 0 Then
                If Not OutputBatchAlreadyLogged(rowVal, procVal, outputVal, batchVal, currentQty) Then
                    totalQty = totalQty + currentQty
                End If
            End If
        End If
    End If
    mTxtOutputTotal.Text = "Output Total: " & FormatRunNumber(totalQty) & IIf(uomVal <> "", " " & uomVal, "")
End Sub

Private Function LoggedOutputTotal(ByVal rowVal As String, ByVal procVal As String, ByVal outputVal As String, _
                                   Optional ByRef loggedCount As Long = 0) As Double
    Dim loLog As ListObject
    Dim r As Long
    Dim cReal As Long
    Dim cProc As Long
    Dim cItem As Long
    Dim cOutput As Long
    Dim cRow As Long
    Dim logRowVal As String
    Dim logProc As String
    Dim logOutput As String

    Set loLog = ProductionLogTable()
    If loLog Is Nothing Then Exit Function
    If loLog.DataBodyRange Is Nothing Then Exit Function

    cReal = ProductionColumnIndex(loLog, "REAL OUTPUT")
    If cReal = 0 Then Exit Function
    cProc = ProductionColumnIndex(loLog, "PROCESS")
    cItem = ProductionColumnIndex(loLog, "ITEM")
    cOutput = ProductionColumnIndex(loLog, "OUTPUT")
    cRow = ProductionColumnIndex(loLog, "ROW")

    For r = 1 To loLog.ListRows.Count
        If Not ProductionLogRowMatchesOutput(loLog, r, rowVal, procVal, outputVal, cRow, cProc, cItem, cOutput) Then GoTo NextRow
        If IsNumeric(loLog.DataBodyRange.Cells(r, cReal).Value) Then
            LoggedOutputTotal = LoggedOutputTotal + CDbl(loLog.DataBodyRange.Cells(r, cReal).Value)
            loggedCount = loggedCount + 1
        End If
NextRow:
    Next r
End Function

Private Function OutputBatchAlreadyLogged(ByVal rowVal As String, ByVal procVal As String, ByVal outputVal As String, _
                                          ByVal batchVal As String, ByVal realQty As Double) As Boolean
    Dim loLog As ListObject
    Dim r As Long
    Dim cReal As Long
    Dim cBatch As Long
    Dim cProc As Long
    Dim cItem As Long
    Dim cOutput As Long
    Dim cRow As Long
    Dim logBatch As String
    Dim logQty As Double

    If Trim$(batchVal) = "" Then Exit Function
    Set loLog = ProductionLogTable()
    If loLog Is Nothing Then Exit Function
    If loLog.DataBodyRange Is Nothing Then Exit Function

    cReal = ProductionColumnIndex(loLog, "REAL OUTPUT")
    cBatch = ProductionColumnIndex(loLog, "BATCH")
    If cReal = 0 Or cBatch = 0 Then Exit Function
    cProc = ProductionColumnIndex(loLog, "PROCESS")
    cItem = ProductionColumnIndex(loLog, "ITEM")
    cOutput = ProductionColumnIndex(loLog, "OUTPUT")
    cRow = ProductionColumnIndex(loLog, "ROW")

    For r = 1 To loLog.ListRows.Count
        If Not ProductionLogRowMatchesOutput(loLog, r, rowVal, procVal, outputVal, cRow, cProc, cItem, cOutput) Then GoTo NextRow
        logBatch = Trim$(NzStr(loLog.DataBodyRange.Cells(r, cBatch).Value))
        If StrComp(logBatch, batchVal, vbTextCompare) = 0 Then
            logQty = NzDblLocal(loLog.DataBodyRange.Cells(r, cReal).Value)
            If Abs(logQty - realQty) < 0.0000001 Then
                OutputBatchAlreadyLogged = True
                Exit Function
            End If
        End If
NextRow:
    Next r
End Function

Private Function ProductionLogRowMatchesOutput(ByVal loLog As ListObject, ByVal rowIndex As Long, _
                                               ByVal rowVal As String, ByVal procVal As String, ByVal outputVal As String, _
                                               ByVal cRow As Long, ByVal cProc As Long, ByVal cItem As Long, ByVal cOutput As Long) As Boolean
    Dim logRowVal As String
    Dim logProc As String
    Dim logOutput As String

    If loLog Is Nothing Then Exit Function
    If loLog.DataBodyRange Is Nothing Then Exit Function
    If cRow > 0 Then
        logRowVal = NormalizeRunRowKey(NzStr(loLog.DataBodyRange.Cells(rowIndex, cRow).Value))
        If NormalizeRunRowKey(rowVal) <> "" And logRowVal = NormalizeRunRowKey(rowVal) Then
            ProductionLogRowMatchesOutput = True
            Exit Function
        End If
    End If

    If cProc > 0 Then logProc = Trim$(NzStr(loLog.DataBodyRange.Cells(rowIndex, cProc).Value))
    If cItem > 0 Then logOutput = Trim$(NzStr(loLog.DataBodyRange.Cells(rowIndex, cItem).Value))
    If logOutput = "" And cOutput > 0 Then logOutput = Trim$(NzStr(loLog.DataBodyRange.Cells(rowIndex, cOutput).Value))
    ProductionLogRowMatchesOutput = (StrComp(logProc, procVal, vbTextCompare) = 0 _
        And StrComp(logOutput, outputVal, vbTextCompare) = 0)
End Function

Private Function ProductionLogTable() As ListObject
    Dim ws As Worksheet

    Set ws = RunProductionObject1("SheetExists", "ProductionLog")
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    Set ProductionLogTable = ws.ListObjects("ProductionLog")
    If ProductionLogTable Is Nothing Then Set ProductionLogTable = ws.ListObjects("Table46")
    On Error GoTo 0
    If Not ProductionLogTable Is Nothing Then Exit Function

    Dim lo As ListObject
    For Each lo In ws.ListObjects
        If ProductionColumnIndex(lo, "PROCESS") > 0 _
           And ProductionColumnIndex(lo, "REAL OUTPUT") > 0 _
           And ProductionColumnIndex(lo, "BATCH") > 0 Then
            Set ProductionLogTable = lo
            Exit Function
        End If
    Next lo
End Function

Private Sub CheckInProductionRun()
    Dim usedPayloadJson As String
    Dim stagedTotal As Double

    If mLstLoaderLines.ListIndex >= 0 Then
        mLstLoaderLines.ListIndex = -1
        RefreshRunPaletteState
    End If
    If mLstRunPalette Is Nothing Or mLstRunPalette.ListCount = 0 Then
        ShowStatus "Load a recipe and choose acceptable inventory before checking inventory into Production."
        Exit Sub
    End If
    If Not ValidateRunAllocationsComplete() Then Exit Sub

    usedPayloadJson = BuildRunUsedPayloadJson(stagedTotal)
    If stagedTotal <= 0 Then
        ShowStatus "No inventory was checked in. Enter allocation quantities first."
        Exit Sub
    End If
    If Not WriteProductionCheckRowsFromRunPalette() Then
        ShowStatus "Check In failed. The Inventory Check list could not be updated. Refresh Production Run before completing."
        Exit Sub
    End If

    RefreshManagerState
    ShowStatus "Checked in " & FormatRunNumber(stagedTotal) & " units to Production. Complete Run will consume these checked-in quantities."
End Sub

Private Sub CompleteProductionRun()
    Dim outputIndex As Long
    Dim processName As String
    Dim prepared As Variant
    Dim completionReport As String

    If Not HasProductionCheckRows() Then
        ShowStatus "Check inventory into Production before completing the run."
        Exit Sub
    End If

    processName = ActiveRunProcess()
    If mLstManagerOutput.ListCount = 0 Then
        prepared = RunProduction0("PrepareProductionOutputForCurrentRecipe")
        RefreshManagerState
        If Not CBool(prepared) Then
            ShowStatus "No production output rows were prepared. Confirm the recipe has an OUTPUT line."
            Exit Sub
        End If
    End If
    If processName <> "" Then
        If Not SelectProductionOutputForProcess(processName) Then
            ShowStatus "No Production Output row matches process '" & processName & "'."
            Exit Sub
        End If
    End If

    outputIndex = mLstManagerOutput.ListIndex
    If outputIndex < 0 And mLstManagerOutput.ListCount = 1 Then
        mLstManagerOutput.ListIndex = 0
        LoadSelectedProductionOutput
        outputIndex = 0
    End If
    If outputIndex < 0 Then
        ShowStatus "Select the Production Output row for this run."
        Exit Sub
    End If
    If Trim$(mTxtOutputBatch.Text) = "" Then mTxtOutputBatch.Text = "1"
    If Trim$(mTxtOutputReal.Text) = "" Then
        ShowStatus "Enter the real output quantity before completing the run."
        Exit Sub
    End If
    If Not IsNumeric(mTxtOutputReal.Text) Or CDbl(mTxtOutputReal.Text) <= 0 Then
        ShowStatus "Real Output must be a number greater than zero."
        Exit Sub
    End If

    ApplySelectedProductionOutput
    If Not CBool(Application.Run("mProduction.CompleteProductionRunAfterCheckInForOutput", outputIndex + 1, completionReport)) Then
        If Trim$(completionReport) = "" Then completionReport = "Complete Run failed."
        ShowStatus completionReport
        Exit Sub
    End If
    RefreshLoaderState
    RefreshManagerState
    ShowStatus "Production run completed. Checked-in inventory was consumed, Real Output was added to inventory, and the batch was logged." & IIf(Trim$(completionReport) <> "", " " & completionReport, "")
End Sub

Private Function HasProductionCheckRows() As Boolean
    Dim lo As ListObject

    Set lo = ProductionTable(TABLE_MANAGER_CHECK)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    HasProductionCheckRows = (lo.ListRows.Count > 0 And Trim$(FirstNonBlankCheckRow(lo)) <> "")
End Function

Private Sub ClearProductionCheckRows()
    Dim lo As ListObject

    Set lo = ProductionTable(TABLE_MANAGER_CHECK)
    If lo Is Nothing Then Exit Sub
    ClearTableContentsKeepRows lo
    If Not mLstManagerCheck Is Nothing Then mLstManagerCheck.Clear
End Sub

Private Function FirstNonBlankCheckRow(ByVal lo As ListObject) As String
    Dim r As Long
    Dim cRow As Long
    Dim cUsed As Long

    cRow = ProductionColumnIndex(lo, "ROW")
    cUsed = ProductionColumnIndex(lo, "USED")
    If cRow = 0 Or cUsed = 0 Then Exit Function
    For r = 1 To lo.ListRows.Count
        If Trim$(NzStr(lo.DataBodyRange.Cells(r, cRow).Value)) <> "" _
           And NzDblLocal(lo.DataBodyRange.Cells(r, cUsed).Value) > 0 Then
            FirstNonBlankCheckRow = NzStr(lo.DataBodyRange.Cells(r, cRow).Value)
            Exit Function
        End If
    Next r
End Function

Private Function SelectProductionOutputForProcess(ByVal processName As String) As Boolean
    Dim i As Long

    If mLstManagerOutput Is Nothing Then Exit Function
    processName = Trim$(processName)
    If processName = "" Then Exit Function
    For i = 0 To mLstManagerOutput.ListCount - 1
        If StrComp(Trim$(NzStr(mLstManagerOutput.List(i, 0))), processName, vbTextCompare) = 0 Then
            mLstManagerOutput.ListIndex = i
            LoadSelectedProductionOutput
            SelectProductionOutputForProcess = True
            Exit Function
        End If
    Next i
End Function

Private Function ValidateRunAllocationsComplete() As Boolean
    Dim seen As Object
    Dim items As Variant
    Dim i As Long
    Dim groupKey As String
    Dim totalPct As Double

    If mLstRunPalette Is Nothing Then Exit Function
    Set seen = CreateObject("Scripting.Dictionary")
    For i = 0 To mLstRunPalette.ListCount - 1
        groupKey = RunAllocationGroupKeyFromList(mLstRunPalette, i)
        If groupKey <> "" Then
            If Not seen.Exists(LCase$(groupKey)) Then seen.Add LCase$(groupKey), groupKey
        End If
    Next i

    items = seen.Items
    For i = LBound(items) To UBound(items)
        groupKey = CStr(items(i))
        totalPct = RunAllocationPercentForGroup(groupKey)
        If totalPct < 99.999 Or totalPct > 100.001 Then
            ShowStatus "Cannot complete run. Ingredient '" & groupKey & "' is " & FormatRunNumber(totalPct) & "% allocated; it must be 100%."
            Exit Function
        End If
    Next i
    ValidateRunAllocationsComplete = (seen.Count > 0)
End Function

Private Function BuildRunUsedPayloadJson(ByRef stagedTotal As Double) As String
    Dim agg As Object
    Dim i As Long
    Dim rowVal As String
    Dim qtyVal As Double
    Dim locVal As String
    Dim key As Variant
    Dim payloadItems As Collection
    Dim payloadItem As Object

    stagedTotal = 0
    Set agg = CreateObject("Scripting.Dictionary")
    For i = 0 To mLstRunPalette.ListCount - 1
        rowVal = Trim$(NzStr(mLstRunPalette.List(i, 3)))
        If rowVal = "" Then GoTo NextChoice
        If Not IsNumeric(NzStr(mLstRunPalette.List(i, 6))) Then GoTo NextChoice
        qtyVal = CDbl(NzStr(mLstRunPalette.List(i, 6)))
        If qtyVal <= 0 Then GoTo NextChoice
        If RunChoiceWouldExceedInventory(i, qtyVal) Then
            ShowStatus "Cannot complete run. ROW " & rowVal & " requires " & FormatRunNumber(qtyVal) & " but only " & NzStr(mLstRunPalette.List(i, 8)) & " is available."
            Exit Function
        End If
        locVal = NzStr(mLstRunPalette.List(i, 9))
        If agg.Exists(rowVal) Then
            Dim existingAgg As Variant
            existingAgg = agg(rowVal)
            agg(rowVal) = Array(CDbl(existingAgg(0)) + qtyVal, locVal)
        Else
            agg.Add rowVal, Array(qtyVal, locVal)
        End If
NextChoice:
    Next i

    If agg.Count = 0 Then Exit Function
    Set payloadItems = New Collection
    For Each key In agg.Keys
        qtyVal = CDbl(agg(key)(0))
        locVal = NzStr(agg(key)(1))
        Set payloadItem = CreateObject("Scripting.Dictionary")
        payloadItem.CompareMode = vbTextCompare
        payloadItem("Row") = CLng(Val(CStr(key)))
        payloadItem("SKU") = ""
        payloadItem("Qty") = qtyVal
        payloadItem("Location") = locVal
        payloadItem("IoType") = "USED"
        payloadItem("Note") = "Production run input"
        payloadItems.Add payloadItem
        stagedTotal = stagedTotal + qtyVal
    Next key
    BuildRunUsedPayloadJson = NzStr(Application.Run("modRoleEventWriter.BuildPayloadJsonFromCollection", payloadItems))
End Function

Private Function WriteProductionCheckRowsFromRunPalette() As Boolean
    Dim lo As ListObject
    Dim agg As Object
    Dim i As Long
    Dim rowVal As String
    Dim qtyVal As Double
    Dim entry As Variant
    Dim key As Variant
    Dim outRow As Long

    Set lo = ProductionTable(TABLE_MANAGER_CHECK)
    If lo Is Nothing Then Exit Function
    Set agg = CreateObject("Scripting.Dictionary")

    For i = 0 To mLstRunPalette.ListCount - 1
        rowVal = Trim$(NzStr(mLstRunPalette.List(i, 3)))
        If rowVal = "" Then GoTo NextChoice
        If Not IsNumeric(NzStr(mLstRunPalette.List(i, 6))) Then GoTo NextChoice
        qtyVal = CDbl(NzStr(mLstRunPalette.List(i, 6)))
        If qtyVal <= 0 Then GoTo NextChoice
        If agg.Exists(rowVal) Then
            entry = agg(rowVal)
            entry(3) = CDbl(entry(3)) + qtyVal
            agg(rowVal) = entry
        Else
            agg.Add rowVal, Array( _
                NzStr(mLstRunPalette.List(i, 4)), _
                NzStr(mLstRunPalette.List(i, 7)), _
                NzStr(mLstRunPalette.List(i, 8)), _
                qtyVal)
        End If
NextChoice:
    Next i

    If agg.Count = 0 Then Exit Function
    ClearTableContentsKeepRows lo
    If Not EnsureTableRows(lo, MaxLongLocal(agg.Count, CurrentRecipeRowBudget())) Then Exit Function
    For Each key In agg.Keys
        outRow = outRow + 1
        entry = agg(key)
        SetCellByHeader lo, outRow, "ROW", CStr(key)
        SetCellByHeader lo, outRow, "ITEM_CODE", ""
        SetCellByHeader lo, outRow, "ITEM", NzStr(entry(0))
        SetCellByHeader lo, outRow, "UOM", NzStr(entry(1))
        SetCellByHeader lo, outRow, "USED", CDbl(entry(3))
        SetCellByHeader lo, outRow, "TOTAL INV", NzStr(entry(2))
    Next key
    WriteProductionCheckRowsFromRunPalette = True
End Function

Private Function RunChoiceWouldExceedInventory(ByVal listIndex As Long, ByVal qtyVal As Double) As Boolean
    Dim invText As String
    Dim invQty As Double

    invText = Trim$(NzStr(mLstRunPalette.List(listIndex, 8)))
    If invText = "" Then Exit Function
    If Not IsNumeric(Left$(invText, 1)) Then Exit Function
    invQty = CDbl(Val(invText))
    RunChoiceWouldExceedInventory = (qtyVal > invQty + 0.0000001)
End Function

Private Function NzDblLocal(ByVal value As Variant) As Double
    If IsError(value) Then Exit Function
    If IsNull(value) Or IsEmpty(value) Then Exit Function
    If IsNumeric(value) Then NzDblLocal = CDbl(value)
End Function

Private Function CurrentRecipeRowBudget() As Long
    CurrentRecipeRowBudget = CLng(Val(NormalizeRecipeRowBudgetText(mTxtRecipeRowBudget.Text)))
End Function

Private Function NormalizeRecipeRowBudgetText(ByVal valueText As String) As String
    Dim n As Long

    n = CLng(Val(Trim$(valueText)))
    If n <= 0 Then n = PRODUCTION_DEFAULT_ROW_BUDGET
    If n > PRODUCTION_MAX_ROW_BUDGET Then n = PRODUCTION_MAX_ROW_BUDGET
    NormalizeRecipeRowBudgetText = CStr(n)
End Function

Private Function MaxLongLocal(ByVal leftValue As Long, ByVal rightValue As Long) As Long
    If leftValue >= rightValue Then
        MaxLongLocal = leftValue
    Else
        MaxLongLocal = rightValue
    End If
End Function

Private Sub LoadSelectedRunPaletteRow()
    Dim idx As Long
    Dim lst As MSForms.ListBox
    Dim splitTextBox As MSForms.TextBox
    Dim qtyTextBox As MSForms.TextBox

    Set lst = ActiveRunPaletteList()
    If lst Is Nothing Then Exit Sub
    Set splitTextBox = ActiveRunSplitTextBox()
    Set qtyTextBox = ActiveRunQtyTextBox()
    If splitTextBox Is Nothing Or qtyTextBox Is Nothing Then Exit Sub
    idx = lst.ListIndex
    If idx < 0 Then Exit Sub
    If IsRunTreeParentRow(lst, idx) Then Exit Sub
    mUpdatingPaletteInputs = True
    splitTextBox.Text = NzStr(lst.List(idx, 5))
    qtyTextBox.Text = NzStr(lst.List(idx, 6))
    mPaletteInputSource = ""
    mUpdatingPaletteInputs = False
    SetRunAllocationVisualState splitTextBox, qtyTextBox, RunAllocationPercentForGroup(RunAllocationGroupKeyFromList(lst, idx))
End Sub

Private Sub ApplySelectedRunPaletteSplit()
    Dim idx As Long
    Dim lst As MSForms.ListBox
    Dim tableName As String
    Dim rowIndex As Long
    Dim lo As ListObject
    Dim splitText As String
    Dim qtyText As String
    Dim baseQtyText As String
    Dim splitVal As Double
    Dim qtyVal As Double
    Dim baseQtyVal As Double
    Dim hasSplit As Boolean
    Dim hasQty As Boolean
    Dim splitTextBox As MSForms.TextBox
    Dim qtyTextBox As MSForms.TextBox
    Dim groupKey As String
    Dim allocationKey As String
    Dim otherPct As Double
    Dim newTotalPct As Double
    Dim runLoc As String
    Dim invLoc As String
    Dim locationWarning As String

    Set lst = ActiveRunPaletteList()
    If lst Is Nothing Then Exit Sub
    Set splitTextBox = ActiveRunSplitTextBox()
    Set qtyTextBox = ActiveRunQtyTextBox()
    If splitTextBox Is Nothing Or qtyTextBox Is Nothing Then Exit Sub
    idx = lst.ListIndex
    If idx < 0 Then
        ShowStatus "Select an acceptable inventory row first."
        Exit Sub
    End If
    If IsRunTreeParentRow(lst, idx) Then
        ShowStatus "Select an acceptable inventory child row first."
        Exit Sub
    End If

    tableName = NzStr(lst.List(idx, 0))
    rowIndex = CLng(Val(NzStr(lst.List(idx, 1))))
    runLoc = ActiveRunLocation()
    invLoc = Trim$(NzStr(lst.List(idx, 9)))
    If RunLocationChoiceRequired() And runLoc = "" Then
        locationWarning = " Choose a production run location before finalizing this run."
    End If
    If runLoc <> "" And invLoc <> "" Then
        If StrComp(runLoc, invLoc, vbTextCompare) <> 0 Then
            locationWarning = " Inventory is at " & invLoc & "; move it to " & runLoc & " before using it."
        End If
    End If

    splitText = Trim$(splitTextBox.Text)
    qtyText = Trim$(qtyTextBox.Text)
    baseQtyVal = RunBaseQtyFromList(lst, idx)

    If splitText <> "" Then
        If Not TryParseNonNegativeRunNumber(splitText, splitVal, "% of Requirement") Then Exit Sub
        hasSplit = True
    End If
    If qtyText <> "" Then
        If Not TryParseNonNegativeRunNumber(qtyText, qtyVal, "Quantity") Then Exit Sub
        hasQty = True
    End If

    If mPaletteInputSource = "QTY" And hasQty Then
        If baseQtyVal > 0 Then
            splitVal = qtyVal / baseQtyVal * 100#
            splitText = FormatRunNumber(splitVal)
            hasSplit = True
            SetPaletteTextSilently splitTextBox, splitText
        End If
    ElseIf hasSplit Then
        If baseQtyVal > 0 Then
            qtyVal = baseQtyVal * splitVal / 100#
            qtyText = FormatRunNumber(qtyVal)
            hasQty = True
            SetPaletteTextSilently qtyTextBox, qtyText
        End If
    ElseIf hasQty And baseQtyVal > 0 Then
        splitVal = qtyVal / baseQtyVal * 100#
        splitText = FormatRunNumber(splitVal)
        hasSplit = True
        SetPaletteTextSilently splitTextBox, splitText
    End If

    If Not hasSplit And Not hasQty Then
        ShowStatus "Enter % of Requirement or Qty first."
        Exit Sub
    End If
    If Not hasSplit Then
        ShowStatus "Cannot validate allocation without a % of Requirement. Select a row with a recipe requirement first."
        Exit Sub
    End If

    groupKey = RunAllocationGroupKeyFromList(lst, idx)
    allocationKey = RunAllocationKeyFromList(lst, idx)
    otherPct = RunAllocationPercentForGroup(groupKey, allocationKey)
    newTotalPct = otherPct + splitVal
    If newTotalPct > 100.001 Then
        SetRunAllocationVisualState splitTextBox, qtyTextBox, newTotalPct
        ShowStatus "Allocation rejected. Other choices already use " & FormatRunNumber(otherPct) & "%. This would make " & FormatRunNumber(newTotalPct) & "%; maximum is 100%."
        Exit Sub
    End If

    If tableName <> "" And rowIndex >= 1 Then
        Set lo = ProductionTable(tableName)
        If Not lo Is Nothing Then
            If Not lo.DataBodyRange Is Nothing Then
                If rowIndex <= lo.ListRows.Count Then
                    If hasSplit Then SetCellByHeader lo, rowIndex, "SPLIT %", splitVal
                    If hasQty Then SetCellByHeader lo, rowIndex, "QUANTITY", qtyVal
                End If
            End If
        End If
    End If

    If hasSplit Then lst.List(idx, 5) = splitText
    If hasQty Then lst.List(idx, 6) = qtyText
    StoreRunAllocationOverride lst, idx, NzStr(lst.List(idx, 5)), NzStr(lst.List(idx, 6))
    SyncRunAllocationToPaletteList lst, idx
    BuildRunTreeFromPaletteList
    SetRunAllocationVisualState splitTextBox, qtyTextBox, newTotalPct
    ShowStatus "Acceptable inventory allocation updated. Ingredient is " & FormatRunNumber(newTotalPct) & "% filled." & locationWarning
End Sub

Private Function TryParseNonNegativeRunNumber(ByVal numberText As String, ByRef numberValue As Double, ByVal labelText As String) As Boolean
    If Not IsNumeric(numberText) Then
        ShowStatus labelText & " must be numeric."
        Exit Function
    End If
    numberValue = CDbl(numberText)
    If numberValue < 0 Then
        ShowStatus labelText & " cannot be negative."
        Exit Function
    End If
    TryParseNonNegativeRunNumber = True
End Function

Private Function FormatRunNumber(ByVal numberValue As Double) As String
    FormatRunNumber = Format$(numberValue, "0.###")
End Function

Private Sub SetPaletteTextSilently(ByVal txt As MSForms.TextBox, ByVal valueText As String)
    mUpdatingPaletteInputs = True
    txt.Text = valueText
    mUpdatingPaletteInputs = False
End Sub

Private Sub SetRunAllocationVisualState(ByVal splitTextBox As MSForms.TextBox, ByVal qtyTextBox As MSForms.TextBox, ByVal totalPct As Double)
    Dim fillColor As Long

    If totalPct > 100.001 Then
        fillColor = RGB(255, 210, 210)
    ElseIf totalPct >= 99.999 Then
        fillColor = RGB(210, 245, 210)
    ElseIf totalPct > 0 Then
        fillColor = RGB(255, 245, 200)
    Else
        fillColor = vbWhite
    End If
    If Not splitTextBox Is Nothing Then
        splitTextBox.BackColor = fillColor
        splitTextBox.Font.Bold = (totalPct >= 99.999 And totalPct <= 100.001)
    End If
    If Not qtyTextBox Is Nothing Then
        qtyTextBox.BackColor = fillColor
        qtyTextBox.Font.Bold = (totalPct >= 99.999 And totalPct <= 100.001)
    End If
End Sub

Private Sub SyncRunAllocationToPaletteList(ByVal sourceList As MSForms.ListBox, ByVal sourceRow As Long)
    Dim sourceKey As String
    Dim i As Long

    If mLstRunPalette Is Nothing Or sourceList Is Nothing Then Exit Sub
    sourceKey = RunAllocationKeyFromList(sourceList, sourceRow)
    If sourceKey = "" Then Exit Sub
    For i = 0 To mLstRunPalette.ListCount - 1
        If RunAllocationKeyFromList(mLstRunPalette, i) = sourceKey Then
            mLstRunPalette.List(i, 5) = NzStr(sourceList.List(sourceRow, 5))
            mLstRunPalette.List(i, 6) = NzStr(sourceList.List(sourceRow, 6))
            StoreRunAllocationOverride mLstRunPalette, i, NzStr(mLstRunPalette.List(i, 5)), NzStr(mLstRunPalette.List(i, 6))
            Exit For
        End If
    Next i
End Sub

Private Function SelectedRunPaletteBaseQty() As Double
    Dim lst As MSForms.ListBox
    Dim idx As Long
    Dim baseText As String

    Set lst = ActiveRunPaletteList()
    If lst Is Nothing Then Exit Function
    idx = lst.ListIndex
    If idx < 0 Then Exit Function
    If IsRunTreeParentRow(lst, idx) Then Exit Function
    SelectedRunPaletteBaseQty = RunBaseQtyFromList(lst, idx)
End Function

Private Function ActiveRunPaletteList() As MSForms.ListBox
    If Not mPages Is Nothing Then
        If mPages.Value = 3 Then
            Set ActiveRunPaletteList = mLstRunTree
            Exit Function
        End If
    End If
    Set ActiveRunPaletteList = mLstRunPalette
End Function

Private Function ActiveRunSplitTextBox() As MSForms.TextBox
    If Not mPages Is Nothing Then
        If mPages.Value = 3 Then
            Set ActiveRunSplitTextBox = mTxtTreePaletteSplit
            Exit Function
        End If
    End If
    Set ActiveRunSplitTextBox = mTxtPaletteSplit
End Function

Private Function ActiveRunQtyTextBox() As MSForms.TextBox
    If Not mPages Is Nothing Then
        If mPages.Value = 3 Then
            Set ActiveRunQtyTextBox = mTxtTreePaletteQty
            Exit Function
        End If
    End If
    Set ActiveRunQtyTextBox = mTxtPaletteQty
End Function

Private Sub EnsureTableRow(ByVal lo As ListObject)
    EnsureTableRows lo, 1
End Sub

Private Function EnsureTableRows(ByVal lo As ListObject, ByVal rowCount As Long) As Boolean
    On Error GoTo Fail

    If lo Is Nothing Then Exit Function
    If rowCount < 1 Then
        EnsureTableRows = True
        Exit Function
    End If
    Do While lo.ListRows.Count < rowCount
        lo.ListRows.Add AlwaysInsert:=False
    Loop
    EnsureTableRows = True
    Exit Function

Fail:
    EnsureTableRows = False
End Function

Private Sub ClearListRows(ByVal lo As ListObject)
    If lo Is Nothing Then Exit Sub
    On Error Resume Next
    Do While lo.ListRows.Count > 0
        lo.ListRows(1).Delete
    Loop
    On Error GoTo 0
End Sub

Private Sub ClearTableContentsKeepRows(ByVal lo As ListObject)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    lo.DataBodyRange.ClearContents
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
    colIndex = ProductionColumnIndex(lo, headerName)
    If colIndex <= 0 Then Exit Sub
    If Not EnsureTableRows(lo, rowIndex) Then Exit Sub
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

Private Function RunProduction2(ByVal procName As String, ByVal arg1 As Variant, ByRef arg2 As Variant) As Variant
    RunProduction2 = Application.Run("mProduction." & procName, arg1, arg2)
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

Private Sub mCmbRunLocation_Change()
    SyncRunLocationCombo mCmbRunLocation, mCmbTreeRunLocation
End Sub

Private Sub mCmbTreeRunLocation_Change()
    SyncRunLocationCombo mCmbTreeRunLocation, mCmbRunLocation
End Sub

Private Sub mTxtPaletteSplit_Change()
    PaletteSplitTextChanged mTxtPaletteSplit, mTxtPaletteQty
End Sub

Private Sub mTxtTreePaletteSplit_Change()
    PaletteSplitTextChanged mTxtTreePaletteSplit, mTxtTreePaletteQty
End Sub

Private Sub mTxtPaletteQty_Change()
    PaletteQtyTextChanged mTxtPaletteQty, mTxtPaletteSplit
End Sub

Private Sub mTxtTreePaletteQty_Change()
    PaletteQtyTextChanged mTxtTreePaletteQty, mTxtTreePaletteSplit
End Sub

Private Sub mTxtOutputReal_Change()
    If mLoading Then Exit Sub
    UpdateOutputRunningTotalDisplay
End Sub

Private Sub PaletteSplitTextChanged(ByVal splitTextBox As MSForms.TextBox, ByVal qtyTextBox As MSForms.TextBox)
    Dim baseQty As Double
    Dim splitVal As Double
    Dim lst As MSForms.ListBox
    Dim idx As Long
    Dim totalPct As Double

    If mLoading Or mUpdatingPaletteInputs Then Exit Sub
    mPaletteInputSource = "SPLIT"
    If splitTextBox Is Nothing Or qtyTextBox Is Nothing Then Exit Sub
    If Trim$(splitTextBox.Text) = "" Then Exit Sub
    If Not IsNumeric(splitTextBox.Text) Then Exit Sub
    splitVal = CDbl(splitTextBox.Text)
    If splitVal < 0 Then Exit Sub
    baseQty = SelectedRunPaletteBaseQty()
    If baseQty <= 0 Then Exit Sub
    SetPaletteTextSilently qtyTextBox, FormatRunNumber(baseQty * splitVal / 100#)
    Set lst = ActiveRunPaletteList()
    If Not lst Is Nothing Then
        idx = lst.ListIndex
        If idx >= 0 And Not IsRunTreeParentRow(lst, idx) Then
            totalPct = RunAllocationPercentForGroup(RunAllocationGroupKeyFromList(lst, idx), RunAllocationKeyFromList(lst, idx)) + splitVal
            SetRunAllocationVisualState splitTextBox, qtyTextBox, totalPct
        End If
    End If
    mPaletteInputSource = "SPLIT"
End Sub

Private Sub PaletteQtyTextChanged(ByVal qtyTextBox As MSForms.TextBox, ByVal splitTextBox As MSForms.TextBox)
    Dim baseQty As Double
    Dim qtyVal As Double
    Dim splitVal As Double
    Dim lst As MSForms.ListBox
    Dim idx As Long
    Dim totalPct As Double

    If mLoading Or mUpdatingPaletteInputs Then Exit Sub
    mPaletteInputSource = "QTY"
    If qtyTextBox Is Nothing Or splitTextBox Is Nothing Then Exit Sub
    If Trim$(qtyTextBox.Text) = "" Then Exit Sub
    If Not IsNumeric(qtyTextBox.Text) Then Exit Sub
    qtyVal = CDbl(qtyTextBox.Text)
    If qtyVal < 0 Then Exit Sub
    baseQty = SelectedRunPaletteBaseQty()
    If baseQty <= 0 Then Exit Sub
    splitVal = qtyVal / baseQty * 100#
    SetPaletteTextSilently splitTextBox, FormatRunNumber(splitVal)
    Set lst = ActiveRunPaletteList()
    If Not lst Is Nothing Then
        idx = lst.ListIndex
        If idx >= 0 And Not IsRunTreeParentRow(lst, idx) Then
            totalPct = RunAllocationPercentForGroup(RunAllocationGroupKeyFromList(lst, idx), RunAllocationKeyFromList(lst, idx)) + splitVal
            SetRunAllocationVisualState splitTextBox, qtyTextBox, totalPct
        End If
    End If
    mPaletteInputSource = "QTY"
End Sub

Private Sub mBtnAssignRefresh_Click()
    ResetInventoryCache
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
    RefreshManagerState
    ShowStatus "Production Run refreshed."
End Sub

Private Sub mBtnLoaderLoad_Click()
    LoadSelectedRecipeIntoLoader
End Sub

Private Sub mLstLoaderLines_Click()
    If mLoading Then Exit Sub
    If mLstLoaderLines.ListIndex < 0 Then Exit Sub
    RefreshRunPaletteState
    ShowStatus "Acceptable inventory filtered for: " & NzStr(mLstLoaderLines.List(mLstLoaderLines.ListIndex, 3))
End Sub

Private Sub mBtnLoaderClear_Click()
    RunProductionSub0 "BtnClearRecipeChooser"
    RefreshLoaderState
    RefreshManagerState
    ShowStatus "Production Run cleared."
End Sub

Private Sub mLstRunPalette_Click()
    If mLoading Then Exit Sub
    LoadSelectedRunPaletteRow
End Sub

Private Sub mLstRunTree_Click()
    If mLoading Then Exit Sub
    If IsRunTreeParentRow(mLstRunTree, mLstRunTree.ListIndex) Then
        ToggleSelectedRunTreeParent
        Exit Sub
    End If
    LoadSelectedRunPaletteRow
End Sub

Private Sub mBtnRunTreeExpandAll_Click()
    SetAllRunTreeGroupsCollapsed False
    ShowStatus "All ingredient choices shown."
End Sub

Private Sub mBtnRunTreeCollapseAll_Click()
    SetAllRunTreeGroupsCollapsed True
    ShowStatus "All ingredient choices hidden."
End Sub

Private Sub mBtnRunApplyPalette_Click()
    ApplySelectedRunPaletteSplit
End Sub

Private Sub mBtnRunTreeApplyPalette_Click()
    ApplySelectedRunPaletteSplit
End Sub

Private Sub mBtnManagerRefresh_Click()
    RefreshLoaderState
    RefreshManagerState
    ShowStatus "Production Run refreshed."
End Sub

Private Sub mLstManagerOutput_Click()
    If mLoading Then Exit Sub
    LoadSelectedProductionOutput
End Sub

Private Sub mBtnManagerCheckIn_Click()
    CheckInProductionRun
End Sub

Private Sub mBtnManagerPrepare_Click()
    PrepareProductionOutput
End Sub

Private Sub mBtnManagerApplyOutput_Click()
    CompleteProductionRun
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

Private Sub mCmbRunProcess_Change()
    If mLoading Then Exit Sub
    SyncRunProcessCombo mCmbRunProcess, mCmbTreeRunProcess
    RefreshRunPaletteState
End Sub

Private Sub mCmbTreeRunProcess_Change()
    If mLoading Then Exit Sub
    SyncRunProcessCombo mCmbTreeRunProcess, mCmbRunProcess
    RefreshRunPaletteState
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
