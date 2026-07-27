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
Private Const RUN_PALETTE_WIDTHS As String = "0 pt;0 pt;180 pt;45 pt;220 pt;60 pt;70 pt;45 pt;105 pt;120 pt"
Private Const RUN_OUTPUT_WIDTHS As String = "85 pt;260 pt;45 pt;70 pt;55 pt;80 pt;105 pt;45 pt"
Private Const RUN_CHECK_WIDTHS As String = "48 pt;120 pt;320 pt;50 pt;70 pt;100 pt"

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
Private WithEvents mBtnBuilderRelease As MSForms.CommandButton
Private WithEvents mBtnLineAdd As MSForms.CommandButton
Private WithEvents mBtnLineUpdate As MSForms.CommandButton
Private WithEvents mBtnLineRemove As MSForms.CommandButton
Private WithEvents mBtnLineMoveUp As MSForms.CommandButton
Private WithEvents mBtnLineMoveDown As MSForms.CommandButton
Private WithEvents mBtnLineUomAdd As MSForms.CommandButton

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
Private mTxtLineUom As MSForms.ComboBox
Private mTxtLineAmount As MSForms.TextBox
Private WithEvents mTxtPaletteSplit As MSForms.TextBox
Private WithEvents mTxtPaletteQty As MSForms.TextBox
Private WithEvents mTxtTreePaletteSplit As MSForms.TextBox
Private WithEvents mTxtTreePaletteQty As MSForms.TextBox
Private WithEvents mTxtOutputReal As MSForms.TextBox
Private mTxtStatus As MSForms.TextBox
Private mOperatorWorkbook As Workbook
Private mInventoryRows As Variant
Private mInventoryCacheLoaded As Boolean
Private mRunInventoryRows As Variant
Private mRunInventoryCacheLoaded As Boolean
Private mBuilt As Boolean
Private mLoading As Boolean
Private mResizeInitialized As Boolean
Private mResizingLayout As Boolean
Private mRunTreeCollapsed As Object
Private mRunSplitOverrides As Object
Private mRunBaseQtyByKey As Object
Private mUpdatingPaletteInputs As Boolean
Private mPaletteInputSource As String
Private mRunProcessByKey As Object
Private mRunItemCodeByKey As Object
Private mBuilderLineTableRows() As Long
Private mBuilderLineTableRowCount As Long

Private Const ASSIGN_INVENTORY_MAX_VISIBLE As Long = 250
Private Const PRODUCTION_BASE_WIDTH As Double = 1110
Private Const PRODUCTION_BASE_HEIGHT As Double = 690
Private Const PRODUCTION_DEFAULT_ROW_BUDGET As Long = 50
Private Const PRODUCTION_MAX_ROW_BUDGET As Long = 1000
Private Const RUN_TREE_PARENT_MARKER As String = "__RUN_TREE_PARENT__"
Private Const RUN_TREE_OUTPUT_MARKER As String = "__RUN_TREE_OUTPUT__"

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
        Application.Run "modUserFormResizeWin.EnableResizableUserForm", Me, True, True
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
    Set mRunProcessByKey = Nothing
    Set mRunItemCodeByKey = Nothing
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
    ResetInventoryCache
    RefreshAllViews
    mLoading = False
    ShowStatus "Production form loaded for " & wb.Name & ". " & _
               NzStr(RunProduction0("GetProductionInventoryModeStatus")) & " " & _
               NzStr(RunProduction0("GetProductionDesignsModeStatus"))
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

Public Function TestRunTwoConsecutiveBatchesForWorkbook(ByVal operatorWb As Workbook, _
                                                        ByVal inputItemCode As String, _
                                                        ByVal inputItemName As String, _
                                                        ByVal inputQty As Double, _
                                                        ByVal inputUom As String, _
                                                        ByVal inputLocation As String, _
                                                        ByVal outputQty As Double, _
                                                        Optional ByVal activatedWb As Workbook = Nothing) As String
    Dim batchNumber As Long
    Dim batchStatus As String
    Dim activeChoices As MSForms.ListBox
    Dim splitInput As MSForms.TextBox
    Dim qtyInput As MSForms.TextBox
    Dim choiceIndex As Long

    If Not mBuilt Then BuildLayout
    SetOperatorWorkbook operatorWb

    For batchNumber = 1 To 2
        If Not PrepareRunChoiceForActionTest(inputItemCode, inputItemName, inputQty, _
                                             inputUom, inputLocation) Then
            TestRunTwoConsecutiveBatchesForWorkbook = _
                "FAIL|Batch=" & CStr(batchNumber) & "|Prepare|" & TestStatusText()
            Exit Function
        End If
        If Not activatedWb Is Nothing Then activatedWb.Activate
        SetRunLocationForActionTest inputLocation
        BuildRunTreeFromPaletteList
        Set activeChoices = ActiveRunPaletteList()
        choiceIndex = FirstSelectableRunChoice(activeChoices)
        If choiceIndex < 0 Then
            TestRunTwoConsecutiveBatchesForWorkbook = _
                "FAIL|Batch=" & CStr(batchNumber) & "|Prepare|No selectable inventory choice."
            Exit Function
        End If
        activeChoices.ListIndex = choiceIndex
        If activeChoices Is mLstRunTree Then
            mLstRunTree_Click
        Else
            mLstRunPalette_Click
        End If
        Set splitInput = ActiveRunSplitTextBox()
        Set qtyInput = ActiveRunQtyTextBox()
        splitInput.Text = "100"
        qtyInput.Text = CStr(inputQty)
        mBtnRunApplyPalette_Click
        mBtnManagerCheckIn_Click
        batchStatus = TestStatusText()
        If InStr(1, batchStatus, "Checked in ", vbTextCompare) = 0 Then
            TestRunTwoConsecutiveBatchesForWorkbook = _
                "FAIL|Batch=" & CStr(batchNumber) & "|CheckIn|" & batchStatus
            Exit Function
        End If

        RefreshManagerState
        If mLstManagerOutput.ListCount = 0 Then
            TestRunTwoConsecutiveBatchesForWorkbook = _
                "FAIL|Batch=" & CStr(batchNumber) & "|OutputMissing|" & TestStatusText()
            Exit Function
        End If
        mLstManagerOutput.ListIndex = 0
        mLstManagerOutput_Click
        mTxtOutputReal.Text = CStr(outputQty)
        mBtnManagerApplyOutput_Click
        batchStatus = TestStatusText()
        If InStr(1, batchStatus, "Production run completed.", vbTextCompare) = 0 Then
            TestRunTwoConsecutiveBatchesForWorkbook = _
                "FAIL|Batch=" & CStr(batchNumber) & "|Complete|" & batchStatus
            Exit Function
        End If
        If batchNumber = 1 Then mBtnManagerNext_Click
    Next batchNumber

    TestRunTwoConsecutiveBatchesForWorkbook = _
        "OK|Batches=2|BoundWorkbook=" & mOperatorWorkbook.Name
End Function

Private Sub SetRunLocationForActionTest(ByVal locationValue As String)
    locationValue = Trim$(locationValue)
    If locationValue = "" Then Exit Sub
    EnsureComboValueForActionTest mCmbRunLocation, locationValue
    EnsureComboValueForActionTest mCmbTreeRunLocation, locationValue
End Sub

Private Sub EnsureComboValueForActionTest(ByVal combo As MSForms.ComboBox, ByVal valueText As String)
    Dim i As Long
    Dim found As Boolean

    If combo Is Nothing Then Exit Sub
    For i = 0 To combo.ListCount - 1
        If StrComp(Trim$(NzStr(combo.List(i))), valueText, vbTextCompare) = 0 Then
            found = True
            Exit For
        End If
    Next i
    If Not found Then combo.AddItem valueText
    combo.Value = valueText
End Sub

Private Function FirstSelectableRunChoice(ByVal choices As MSForms.ListBox) As Long
    Dim i As Long

    FirstSelectableRunChoice = -1
    If choices Is Nothing Then Exit Function
    For i = 0 To choices.ListCount - 1
        If Not IsRunTreeParentRow(choices, i) And Not IsRunTreeOutputRow(choices, i) Then
            FirstSelectableRunChoice = i
            Exit Function
        End If
    Next i
End Function

Private Function PrepareRunChoiceForActionTest(ByVal itemCode As String, _
                                               ByVal itemName As String, _
                                               ByVal qtyValue As Double, _
                                               ByVal uomValue As String, _
                                               ByVal locationValue As String) As Boolean
    Dim values(1 To 1, 1 To 11) As Variant

    mLstRunPalette.Clear
    Set mRunItemCodeByKey = Nothing
    Set mRunProcessByKey = Nothing
    Set mRunBaseQtyByKey = Nothing
    values(1, 1) = "TEST"
    values(1, 2) = "TEST-INPUT"
    values(1, 3) = "Test input"
    values(1, 4) = ""
    values(1, 5) = itemName
    values(1, 6) = ""
    values(1, 7) = ""
    values(1, 8) = uomValue
    values(1, 9) = locationValue
    values(1, 10) = qtyValue
    values(1, 11) = itemCode
    AddRunChoiceRows values
    If mLstRunPalette.ListCount <> 1 Then Exit Function
    mLstRunPalette.ListIndex = 0
    PrepareRunChoiceForActionTest = True
End Function

Public Function TestSelectedProductionOutputTableRow(ByVal wb As Workbook, ByVal listIndex As Long) As Long
    If Not mBuilt Then BuildLayout
    SetOperatorWorkbook wb
    If Not wb Is Nothing Then wb.Activate
    RefreshProductionOutputList mLstManagerOutput
    If listIndex < 0 Or listIndex >= mLstManagerOutput.ListCount Then Exit Function
    mLstManagerOutput.ListIndex = listIndex
    TestSelectedProductionOutputTableRow = SelectedProductionOutputTableRow()
End Function

Public Function TestProductionOutputDisplayedBatch(ByVal wb As Workbook, ByVal listIndex As Long) As String
    If Not mBuilt Then BuildLayout
    SetOperatorWorkbook wb
    If Not wb Is Nothing Then wb.Activate
    RefreshProductionOutputList mLstManagerOutput
    If listIndex < 0 Or listIndex >= mLstManagerOutput.ListCount Then Exit Function
    TestProductionOutputDisplayedBatch = NzStr(mLstManagerOutput.List(listIndex, 4))
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

Public Function TestProductionRunLocationAllowed(ByVal runLocation As String, ByVal inventoryLocation As String, ByVal qtyValue As Double) As Long
    TestProductionRunLocationAllowed = IIf(RunChoiceLocationAllowed(runLocation, inventoryLocation, qtyValue), 1, 0)
End Function

Public Function TestRunPaletteCanonicalItemCodeStorage() As Long
    Dim values(1 To 1, 1 To 11) As Variant

    If Not mBuilt Then BuildLayout
    mLstRunPalette.Clear
    mRunInventoryRows = Empty
    mRunInventoryCacheLoaded = True
    Set mRunItemCodeByKey = Nothing

    values(1, 1) = "BREW"
    values(1, 2) = "ING-WATER"
    values(1, 3) = "Filtered Water"
    values(1, 4) = ""
    values(1, 5) = "Filtered Water"
    values(1, 7) = 500
    values(1, 8) = "LB"
    values(1, 10) = 500
    values(1, 11) = "ITEM-0061"
    AddRunChoiceRows values

    If mLstRunPalette.ColumnCount = 10 _
       And mLstRunPalette.ListCount = 1 _
       And StrComp(RunItemCodeFromList(mLstRunPalette, 0), "ITEM-0061", vbTextCompare) = 0 Then
        TestRunPaletteCanonicalItemCodeStorage = 1
    End If
End Function

Public Function TestAssignmentItemRowsGrowWithoutTableCollision(ByVal wb As Workbook) As Long
    Dim lo As ListObject
    Dim lr As ListRow
    Dim i As Long

    If Not mBuilt Then BuildLayout
    SetOperatorWorkbook wb
    If Not wb Is Nothing Then wb.Activate
    Set lo = ProductionTable(TABLE_ASSIGN_ITEM)
    If lo Is Nothing Then Exit Function
    ClearTableContentsKeepBlank lo

    For i = 1 To 3
        Set lr = WritableAssignmentItemRow(lo)
        If lr Is Nothing Then Exit Function
        SetRowValue lr, lo, "ITEMS", "Test acceptable " & CStr(i)
        SetRowValue lr, lo, "ITEM_CODE", "TEST-SKU-" & CStr(i)
    Next i

    If lo.ListRows.Count >= 3 _
       And StrComp(CellByHeader(lo, 3, "ITEM_CODE"), "TEST-SKU-3", vbTextCompare) = 0 Then
        TestAssignmentItemRowsGrowWithoutTableCollision = 1
    End If
End Function

Public Function TestProductionCheckRowsRecognizeSkuIdentity(ByVal wb As Workbook) As Long
    Dim lo As ListObject

    If Not mBuilt Then BuildLayout
    SetOperatorWorkbook wb
    If Not wb Is Nothing Then wb.Activate
    Set lo = ProductionTable(TABLE_MANAGER_CHECK)
    If lo Is Nothing Then Exit Function
    ClearTableContentsKeepBlank lo
    SetCellByHeader lo, 1, "ROW", ""
    SetCellByHeader lo, 1, "ITEM_CODE", "SKU-CHECK-ONLY"
    SetCellByHeader lo, 1, "ITEM", "SKU-only checked input"
    SetCellByHeader lo, 1, "USED", 12
    If HasProductionCheckRows() Then TestProductionCheckRowsRecognizeSkuIdentity = 1
End Function

Public Function TestRecipeBuilderSelectedLineProcessUpdate(ByVal wb As Workbook, _
                                                           ByVal newProcess As String, _
                                                           Optional ByVal selectedIndex As Long = 0) As Long
    Dim lo As ListObject
    Dim tableRow As Long

    If Not mBuilt Then BuildLayout
    SetOperatorWorkbook wb
    If Not wb Is Nothing Then wb.Activate
    RefreshBuilderLines
    If mLstBuilderLines Is Nothing Then Exit Function
    If selectedIndex < 0 Or selectedIndex >= mLstBuilderLines.ListCount Then Exit Function
    mLstBuilderLines.ListIndex = selectedIndex
    tableRow = BuilderTableRowForListIndex(selectedIndex)
    If tableRow <= 0 Then Exit Function
    LoadSelectedBuilderLine
    mTxtLineProcess.Text = newProcess
    If Not WriteSelectedRecipeBuilderLineFromForm(False) Then Exit Function
    Set lo = ProductionTable(TABLE_BUILDER_LINES)
    If Not lo Is Nothing Then
        If StrComp(CellByHeader(lo, tableRow, "PROCESS"), newProcess, vbTextCompare) = 0 Then _
            TestRecipeBuilderSelectedLineProcessUpdate = 1
    End If
End Function

Public Function TestRecipeBuilderLineMove(ByVal wb As Workbook, _
                                          ByVal selectedIndex As Long, _
                                          ByVal moveDirection As Long) As String
    Dim lo As ListObject

    If Not mBuilt Then BuildLayout
    SetOperatorWorkbook wb
    If Not wb Is Nothing Then wb.Activate
    RefreshBuilderLines
    If mLstBuilderLines Is Nothing Then Exit Function
    If selectedIndex < 0 Or selectedIndex >= mLstBuilderLines.ListCount Then Exit Function
    mLstBuilderLines.ListIndex = selectedIndex
    If Not MoveSelectedRecipeBuilderLine(moveDirection, False) Then Exit Function

    Set lo = ProductionTable(TABLE_BUILDER_LINES)
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    TestRecipeBuilderLineMove = CellByHeader(lo, 1, "INGREDIENT") & "|" & _
                                CellByHeader(lo, 2, "INGREDIENT") & "|" & _
                                CellByHeader(lo, 1, "RECIPE_LIST_ROW") & "|" & _
                                CellByHeader(lo, 2, "RECIPE_LIST_ROW")
End Function

Public Function TestRecipeBuilderHasInstructionIo() As Long
    Dim i As Long

    If Not mBuilt Then BuildLayout
    For i = 0 To mCmbLineIo.ListCount - 1
        If StrComp(NzStr(mCmbLineIo.List(i)), "INSTRUCTION", vbTextCompare) = 0 Then
            TestRecipeBuilderHasInstructionIo = 1
            Exit Function
        End If
    Next i
End Function

Public Function TestRecipeBuilderLineActionsFitLayout() As Long
    If Not mBuilt Then BuildLayout

    If mBtnLineAdd.Left + mBtnLineAdd.Width + 10 > mBtnLineUpdate.Left Then Exit Function
    If mBtnLineUpdate.Left + mBtnLineUpdate.Width + 10 > mBtnLineRemove.Left Then Exit Function
    If mBtnLineRemove.Left + mBtnLineRemove.Width + 10 > mBtnLineMoveUp.Left Then Exit Function
    If mBtnLineMoveUp.Left + mBtnLineMoveUp.Width + 10 > mBtnLineMoveDown.Left Then Exit Function
    If mBtnLineMoveDown.Left + mBtnLineMoveDown.Width + 10 > mBtnBuilderClear.Left Then Exit Function
    If mBtnLineAdd.Top + mBtnLineAdd.Height + 10 > mLstBuilderLines.Top Then Exit Function

    TestRecipeBuilderLineActionsFitLayout = 1
End Function

Public Function TestRecipeBuilderLifecycleAndHeadersReady() As Long
    If Not mBuilt Then BuildLayout
    If mBtnBuilderRelease Is Nothing Then Exit Function
    If StrComp(mBtnBuilderRelease.Caption, "Release for Production", vbTextCompare) <> 0 Then Exit Function
    If Not ControlExistsByName(mPages.Pages(0), "hdrBuilderLines1") Then Exit Function
    If Not ControlExistsByName(mPages.Pages(0), "hdrBuilderLines8") Then Exit Function
    TestRecipeBuilderLifecycleAndHeadersReady = 1
End Function

Public Function TestRecipeBuilderUomCatalogContains(ByVal uomName As String) As Long
    Dim idx As Long

    If Not mBuilt Then BuildLayout
    RefreshRecipeUomCatalog
    For idx = 0 To mTxtLineUom.ListCount - 1
        If StrComp(CStr(mTxtLineUom.List(idx)), Trim$(uomName), vbTextCompare) = 0 Then
            TestRecipeBuilderUomCatalogContains = 1
            Exit Function
        End If
    Next idx
End Function

Public Function TestRecipeBuilderSelectUom(ByVal uomName As String) As String
    If Not mBuilt Then BuildLayout
    RefreshRecipeUomCatalog uomName
    If mTxtLineUom.ListIndex >= 0 Then _
        TestRecipeBuilderSelectUom = CStr(mTxtLineUom.List(mTxtLineUom.ListIndex))
End Function

Public Function TestWriteRecipeHeaderToBoundWorkbook(ByVal wb As Workbook, _
                                                     ByVal recipeName As String, _
                                                     ByVal recipeId As String) As String
    Dim lo As ListObject

    If Not mBuilt Then BuildLayout
    SetOperatorWorkbook wb
    mTxtRecipeName.Text = recipeName
    mTxtRecipeId.Text = recipeId
    mTxtRecipeDescription.Text = "bound workbook test"
    mTxtRecipeRowBudget.Text = CStr(PRODUCTION_DEFAULT_ROW_BUDGET)
    WriteRecipeHeaderFromForm

    Set lo = ProductionTable(TABLE_BUILDER_HEADER)
    If lo Is Nothing Then Exit Function
    TestWriteRecipeHeaderToBoundWorkbook = FirstRowValue(lo, "RECIPE_ID") & "|" & _
                                           FirstRowValue(lo, "RECIPE_NAME")
End Function

Public Function TestAssignmentIngredientProcessForRecipe(ByVal wb As Workbook, ByVal recipeId As String, ByVal ingredientName As String) As String
    Dim lo As ListObject
    Dim i As Long

    If Not mBuilt Then BuildLayout
    SetOperatorWorkbook wb
    If Not wb Is Nothing Then wb.Activate

    Set lo = ProductionTable(TABLE_ASSIGN_RECIPE)
    If lo Is Nothing Then Exit Function
    EnsureTableRow lo
    SetFirstRowValue lo, "RECIPE_ID", recipeId
    RefreshAssignmentState

    For i = 0 To mLstAssignIngredients.ListCount - 1
        If StrComp(NzStr(mLstAssignIngredients.List(i, 1)), ingredientName, vbTextCompare) = 0 Then
            TestAssignmentIngredientProcessForRecipe = NzStr(mLstAssignIngredients.List(i, 3))
            Exit Function
        End If
    Next i
End Function

Public Function TestAssignmentOutputSelectionClearsStaging(ByVal wb As Workbook, ByVal recipeId As String, ByVal outputName As String) As Long
    Dim loRecipe As ListObject
    Dim loIng As ListObject
    Dim loItems As ListObject
    Dim lr As ListRow
    Dim i As Long

    If Not mBuilt Then BuildLayout
    SetOperatorWorkbook wb
    If Not wb Is Nothing Then wb.Activate

    Set loRecipe = ProductionTable(TABLE_ASSIGN_RECIPE)
    Set loIng = ProductionTable(TABLE_ASSIGN_INGREDIENT)
    Set loItems = ProductionTable(TABLE_ASSIGN_ITEM)
    If loRecipe Is Nothing Or loIng Is Nothing Or loItems Is Nothing Then Exit Function

    EnsureTableRow loRecipe
    SetFirstRowValue loRecipe, "RECIPE_ID", recipeId
    SetFirstRowValue loRecipe, "RECIPE_NAME", "Recipe"

    EnsureTableRow loIng
    SetFirstRowValue loIng, "RECIPE_ID", recipeId
    SetFirstRowValue loIng, "INGREDIENT_ID", "OLD-ING"
    SetFirstRowValue loIng, "INGREDIENT", "Old ingredient"

    Set lr = loItems.ListRows.Add(AlwaysInsert:=False)
    SetRowValue lr, loItems, "RECIPE_ID", recipeId
    SetRowValue lr, loItems, "INGREDIENT_ID", "OLD-ING"
    SetRowValue lr, loItems, "ITEMS", "Old acceptable"

    RefreshAssignmentState
    For i = 0 To mLstAssignIngredients.ListCount - 1
        If StrComp(NzStr(mLstAssignIngredients.List(i, 1)), outputName, vbTextCompare) = 0 Then
            mLstAssignIngredients.ListIndex = i
            SelectAssignmentIngredientFromList
            Exit For
        End If
    Next i
    If i >= mLstAssignIngredients.ListCount Then Exit Function

    If FirstRowValue(loIng, "INGREDIENT_ID") <> "" Then Exit Function
    If Not loItems.DataBodyRange Is Nothing Then
        For i = 1 To loItems.DataBodyRange.Rows.Count
            If CellByHeader(loItems, i, "ITEMS") <> "" Then Exit Function
            If CellByHeader(loItems, i, "ITEM") <> "" Then Exit Function
            If CellByHeader(loItems, i, "INGREDIENT_ID") <> "" Then Exit Function
        Next i
    End If
    If InStr(1, TestStatusText(), "Outputs do not accept ingredients", vbTextCompare) > 0 Then
        TestAssignmentOutputSelectionClearsStaging = 1
    End If
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
    Set mBtnBuilderRelease = AddButton(pg, "btnBuilderRelease", "Release for Production", 860, 256, 170, 24)

    AddLabel pg, "Process", 350, 145, 70, 16
    Set mTxtLineProcess = AddText(pg, "txtLineProcess", 350, 165, 120, 22)
    AddLabel pg, "In/Out", 485, 145, 70, 16
    Set mCmbLineIo = AddCombo(pg, "cmbLineIo", 485, 165, 95, 22)
    mCmbLineIo.AddItem "USED"
    mCmbLineIo.AddItem "OUTPUT"
    mCmbLineIo.AddItem "INSTRUCTION"
    mCmbLineIo.ListIndex = 0
    AddLabel pg, "Ingredient / Output / Instruction", 595, 145, 220, 16
    Set mTxtLineIngredient = AddText(pg, "txtLineIngredient", 595, 165, 220, 22)
    AddLabel pg, "Percent", 350, 197, 70, 16
    Set mTxtLinePercent = AddText(pg, "txtLinePercent", 350, 217, 70, 22)
    AddLabel pg, "UOM", 435, 197, 55, 16
    Set mTxtLineUom = AddCombo(pg, "cmbLineUom", 435, 217, 90, 22)
    RefreshRecipeUomCatalog
    Set mBtnLineUomAdd = AddButton(pg, "btnLineUomAdd", "Add UOM", 535, 216, 75, 24)
    AddLabel pg, "Amount", 625, 197, 70, 16
    Set mTxtLineAmount = AddText(pg, "txtLineAmount", 625, 217, 90, 22)
    Set mBtnLineAdd = AddButton(pg, "btnLineAdd", "Add Line", 350, 252, 90, 24)
    Set mBtnLineUpdate = AddButton(pg, "btnLineUpdate", "Update Line", 450, 252, 90, 24)
    Set mBtnLineRemove = AddButton(pg, "btnLineRemove", "Remove Line", 550, 252, 90, 24)
    Set mBtnLineMoveUp = AddButton(pg, "btnLineMoveUp", "Move Up", 650, 252, 90, 24)
    Set mBtnLineMoveDown = AddButton(pg, "btnLineMoveDown", "Move Down", 750, 252, 90, 24)

    AddLabel pg, "Recipe Builder Lines", 12, 290, 220, 16
    AddColumnHeaders pg, "BuilderLines", _
        Array("Process", "", "I/O", "Ingredient / Output / Instruction", "%", "UOM", "Amount", "Ingredient ID"), _
        12, 310, "90 pt;55 pt;70 pt;210 pt;55 pt;55 pt;55 pt;70 pt"
    Set mLstBuilderLines = AddList(pg, "lstBuilderLines", 12, 328, 1018, 192, 8, "90 pt;55 pt;70 pt;210 pt;55 pt;55 pt;55 pt;70 pt")
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
    Set mLstAssignAllowed = AddList(pg, "lstAssignAllowed", 540, 312, 490, 208, 7, "45 pt;160 pt;45 pt;170 pt;0 pt;0 pt;0 pt")
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

    AddLabel pg, "Process", 12, 170, 60, 16
    Set mCmbRunProcess = AddCombo(pg, "cmbRunProcess", 75, 166, 160, 22)
    AddLabel pg, "Run Location", 250, 170, 90, 16
    Set mCmbRunLocation = AddCombo(pg, "cmbRunLocation", 345, 166, 160, 22)
    AddLabel pg, "% of Requirement", 525, 170, 110, 16
    Set mTxtPaletteSplit = AddText(pg, "txtPaletteSplit", 640, 166, 90, 22)
    AddLabel pg, "Qty", 750, 170, 45, 16
    Set mTxtPaletteQty = AddText(pg, "txtPaletteQty", 790, 166, 90, 22)
    Set mBtnRunApplyPalette = AddButton(pg, "btnRunApplyPalette", "Apply", 900, 165, 90, 24)

    AddLabel pg, "Acceptable Inventory For Run", 12, 182, 230, 16
    AddColumnHeaders pg, "RunPalette", Array("", "", "Ingredient", "ROW", "Inventory Item", "% Req", "Qty", "UOM", "Inv", "Location"), 12, 202, RUN_PALETTE_WIDTHS
    Set mLstRunPalette = AddList(pg, "lstRunPalette", 12, 220, 1018, 80, 10, RUN_PALETTE_WIDTHS)

    AddLabel pg, "Inventory Check", 12, 316, 150, 16
    AddColumnHeaders pg, "ManagerCheck", Array("ROW", "Code", "Item", "UOM", "Used", "Total Inv"), 12, 336, RUN_CHECK_WIDTHS
    Set mLstManagerCheck = AddList(pg, "lstManagerCheck", 12, 354, 1018, 56, 6, RUN_CHECK_WIDTHS)

    AddLabel pg, "Production Output", 12, 426, 170, 16
    AddColumnHeaders pg, "ManagerOutput", Array("Process", "Output", "UOM", "Last", "Batch", "Total", "Recall", "Inventory ID"), 12, 446, RUN_OUTPUT_WIDTHS
    Set mLstManagerOutput = AddList(pg, "lstManagerOutput", 12, 464, 1018, 48, 8, RUN_OUTPUT_WIDTHS)

    AddLabel pg, "Real Output", 12, 526, 80, 16
    Set mTxtOutputReal = AddText(pg, "txtOutputReal", 100, 522, 105, 22)
    Set mBtnManagerCheckIn = AddButton(pg, "btnManagerCheckIn", "Check In", 230, 520, 95, 24)
    Set mBtnManagerApplyOutput = AddButton(pg, "btnManagerApplyOutput", "Complete Run", 340, 520, 120, 24)
    Set mBtnManagerRefresh = AddButton(pg, "btnManagerRefresh", "Refresh", 480, 520, 95, 24)
    Set mBtnManagerNext = AddButton(pg, "btnManagerNext", "Next Batch", 595, 520, 120, 24)
    Set mBtnManagerPrint = AddButton(pg, "btnManagerPrint", "Print Recall", 735, 520, 120, 24)
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
    Dim runListLeft As Double
    Dim runListWidth As Double
    Dim paletteTop As Double
    Dim paletteHeight As Double
    Dim checkTop As Double
    Dim checkHeight As Double
    Dim outputTop As Double
    Dim outputHeight As Double
    Dim controlsTop As Double
    Dim availableRunHeight As Double
    Dim remainingRunHeight As Double

    If mPages Is Nothing Then Exit Sub
    pageW = MaxDoubleForm(700, mPages.Width - 20)
    pageH = MaxDoubleForm(420, mPages.Height - 45)

    If Not mLstBuilderRecipes Is Nothing Then mLstBuilderRecipes.Height = MaxDoubleForm(130, pageH - 300)
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

    If Not mLstLoaderRecipes Is Nothing Then mLstLoaderRecipes.Height = 112
    If Not mBtnLoaderRefresh Is Nothing Then mBtnLoaderRefresh.Top = 32
    If Not mBtnLoaderLoad Is Nothing Then mBtnLoaderLoad.Top = 64
    If Not mBtnLoaderClear Is Nothing Then mBtnLoaderClear.Top = 96
    If Not mLstLoaderLines Is Nothing Then
        mLstLoaderLines.Width = MaxDoubleForm(420, pageW - mLstLoaderLines.Left - 30)
        mLstLoaderLines.Height = 112
    End If

    runListLeft = 12
    runListWidth = MaxDoubleForm(640, pageW - 40)

    If Not mCmbRunProcess Is Nothing Then mCmbRunProcess.Move 75, 166, 160, 22
    If Not mCmbRunLocation Is Nothing Then mCmbRunLocation.Move 345, 166, 160, 22
    If Not mTxtPaletteSplit Is Nothing Then mTxtPaletteSplit.Move 640, 166, 90, 22
    If Not mTxtPaletteQty Is Nothing Then mTxtPaletteQty.Move 790, 166, 90, 22
    If Not mBtnRunApplyPalette Is Nothing Then mBtnRunApplyPalette.Move 900, 165, 90, 24

    paletteTop = 220
    controlsTop = MaxDoubleForm(520, pageH - 34)
    availableRunHeight = MaxDoubleForm(225, controlsTop - paletteTop - 84)
    paletteHeight = MaxDoubleForm(82, availableRunHeight * 0.4)
    MoveLabelByCaption mPages.Pages(2), "Acceptable Inventory For Run", runListLeft, paletteTop - 38, 230, 16
    If Not mLstRunPalette Is Nothing Then
        mLstRunPalette.Move runListLeft, paletteTop, runListWidth, paletteHeight
        PositionColumnHeaders mPages.Pages(2), "RunPalette", runListLeft, paletteTop - 18, RUN_PALETTE_WIDTHS
    End If

    checkTop = paletteTop + paletteHeight + 42
    remainingRunHeight = MaxDoubleForm(120, controlsTop - checkTop - 42)
    checkHeight = MaxDoubleForm(62, remainingRunHeight * 0.45)
    MoveLabelByCaption mPages.Pages(2), "Inventory Check", runListLeft, checkTop - 38, 150, 16
    If Not mLstManagerCheck Is Nothing And Not mLstRunPalette Is Nothing Then
        mLstManagerCheck.Move runListLeft, checkTop, runListWidth, checkHeight
        PositionColumnHeaders mPages.Pages(2), "ManagerCheck", runListLeft, checkTop - 18, RUN_CHECK_WIDTHS
    End If

    outputTop = checkTop + checkHeight + 42
    outputHeight = MaxDoubleForm(58, controlsTop - outputTop - 14)
    MoveLabelByCaption mPages.Pages(2), "Production Output", runListLeft, outputTop - 38, 170, 16
    If Not mLstManagerOutput Is Nothing And Not mLstManagerCheck Is Nothing Then
        mLstManagerOutput.Move runListLeft, outputTop, runListWidth, outputHeight
        PositionColumnHeaders mPages.Pages(2), "ManagerOutput", runListLeft, outputTop - 18, RUN_OUTPUT_WIDTHS
    End If

    MoveLabelByCaption mPages.Pages(2), "Real Output", 12, controlsTop, 80, 16
    If Not mTxtOutputReal Is Nothing Then mTxtOutputReal.Move 100, controlsTop - 4, 105, 22
    If Not mBtnManagerCheckIn Is Nothing Then mBtnManagerCheckIn.Move 230, controlsTop - 6, 95, 24
    If Not mBtnManagerApplyOutput Is Nothing Then mBtnManagerApplyOutput.Move 340, controlsTop - 6, 120, 24
    If Not mBtnManagerRefresh Is Nothing Then mBtnManagerRefresh.Move 480, controlsTop - 6, 95, 24
    If Not mBtnManagerNext Is Nothing Then mBtnManagerNext.Move 595, controlsTop - 6, 120, 24
    If Not mBtnManagerPrint Is Nothing Then
        mBtnManagerPrint.Move 735, controlsTop - 6, 120, 24
        If pageW > 930 Then mBtnManagerPrint.Left = pageW - 180
    End If

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

Private Sub MoveLabelByCaption(ByVal parent As Object, ByVal caption As String, ByVal leftVal As Single, _
                               ByVal topVal As Single, ByVal widthVal As Single, ByVal heightVal As Single)
    Dim ctl As MSForms.Control
    For Each ctl In parent.Controls
        If TypeName(ctl) = "Label" Then
            If StrComp(CStr(ctl.Caption), caption, vbTextCompare) = 0 Then
                ctl.Move leftVal, topVal, widthVal, heightVal
                Exit Sub
            End If
        End If
    Next ctl
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

Private Function ControlExistsByName(ByVal parent As Object, ByVal controlName As String) As Boolean
    Dim ctl As Object

    On Error Resume Next
    Set ctl = parent.Controls(controlName)
    ControlExistsByName = Not ctl Is Nothing
    On Error GoTo 0
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
    Dim releasedRecipes As Variant

    recipes = RunProduction0("LoadRecipeList")
    releasedRecipes = RunProduction0("LoadReleasedRecipeList")
    FillListFromArray mLstBuilderRecipes, recipes
    FillListFromArray mLstAssignRecipes, recipes
    FillListFromArray mLstLoaderRecipes, releasedRecipes
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
        mTxtRecipeId.Text = GenerateRecipeIdForOperatorWorkbook()
    End If
End Sub

Private Sub RefreshBuilderLines()
    Dim lo As ListObject
    Dim headers As Variant
    Dim arr As Variant
    Dim r As Long
    Dim c As Long
    Dim colIdx As Long

    headers = Array("PROCESS", "DIAGRAM_ID", "INPUT/OUTPUT", "INGREDIENT", "PERCENT", "UOM", "AMOUNT", "INGREDIENT_ID")
    mLstBuilderLines.Clear
    mBuilderLineTableRowCount = 0
    Erase mBuilderLineTableRows

    Set lo = ProductionTable(TABLE_BUILDER_LINES)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    arr = lo.DataBodyRange.Value
    For r = 1 To UBound(arr, 1)
        If Not TableArrayRowHasAnyValue(arr, r, lo, headers) Then GoTo NextBuilderTableRow
        mLstBuilderLines.AddItem CellText(arr, r, lo, CStr(headers(LBound(headers))))
        For c = LBound(headers) + 1 To UBound(headers)
            colIdx = c - LBound(headers)
            If colIdx < mLstBuilderLines.ColumnCount Then _
                mLstBuilderLines.List(mLstBuilderLines.ListCount - 1, colIdx) = CellText(arr, r, lo, CStr(headers(c)))
        Next c

        mBuilderLineTableRowCount = mBuilderLineTableRowCount + 1
        ReDim Preserve mBuilderLineTableRows(0 To mBuilderLineTableRowCount - 1)
        mBuilderLineTableRows(mBuilderLineTableRowCount - 1) = r
NextBuilderTableRow:
    Next r
End Sub

Private Function BuilderTableRowForListIndex(ByVal listIndex As Long) As Long
    If listIndex < 0 Then Exit Function
    If listIndex >= mBuilderLineTableRowCount Then Exit Function
    BuilderTableRowForListIndex = mBuilderLineTableRows(listIndex)
End Function

Private Sub RefreshRecipeUomCatalog(Optional ByVal selectedUom As String = "")
    Dim uoms As Variant
    Dim idx As Long

    If mTxtLineUom Is Nothing Then Exit Sub
    If selectedUom = "" Then selectedUom = Trim$(CStr(mTxtLineUom.Value))
    selectedUom = Trim$(selectedUom)
    mTxtLineUom.Clear
    uoms = modUomSettings.GetConfiguredUoms()
    If IsArray(uoms) Then
        For idx = LBound(uoms) To UBound(uoms)
            mTxtLineUom.AddItem CStr(uoms(idx))
        Next idx
    End If
    If selectedUom = "" Then Exit Sub

    For idx = 0 To mTxtLineUom.ListCount - 1
        If StrComp(CStr(mTxtLineUom.List(idx)), selectedUom, vbTextCompare) = 0 Then
            mTxtLineUom.ListIndex = idx
            Exit Sub
        End If
    Next idx

    ' Preserve older recipe UOMs without silently adding them to warehouse config.
    mTxtLineUom.AddItem selectedUom
    mTxtLineUom.ListIndex = mTxtLineUom.ListCount - 1
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
        Array("ROW", "ITEMS", "UOM", "DESCRIPTION", "RECIPE_ID", "INGREDIENT_ID", "ITEM_CODE")
End Sub

Private Sub RefreshLoaderState()
    FillListFromTable mLstLoaderLines, ProductionTable(TABLE_LOADER_LINES), _
        Array("PROCESS", "DIAGRAM_ID", "INPUT/OUTPUT", "INGREDIENT", "PERCENT", "UOM", "AMOUNT NEEDED", "INGREDIENT_ID")
    RefreshRunProcessChoices
    If Not mLstLoaderOutput Is Nothing Then
        RefreshProductionOutputList mLstLoaderOutput
    End If
    RefreshRunPaletteState
End Sub

Private Sub RefreshManagerState()
    RefreshProductionOutputList mLstManagerOutput
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
    EnsureRunItemCodeMap
    mRunItemCodeByKey.RemoveAll
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
    Dim itemCode As String
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
        itemCode = vbNullString
        If UBound(values, 2) >= 11 Then itemCode = NzStr(values(r, 11))
        HydrateRunInventoryDisplay rowVal, itemCode, itemVal, uomVal, invVal, locVal

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
        StoreRunProcess mLstRunPalette, listRow, processVal
        StoreRunBaseQty mLstRunPalette, listRow, NzStr(values(r, 10))
        StoreRunItemCode mLstRunPalette, listRow, itemCode
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
    Dim itemCode As String

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
        itemCode = CellText(arr, r, lo, "ITEM_CODE")
        HydrateRunInventoryDisplay rowVal, itemCode, itemVal, uomVal, invVal, locVal

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
        StoreRunProcess mLstRunPalette, listRow, processVal
        StoreRunBaseQty mLstRunPalette, listRow, baseQty
        StoreRunItemCode mLstRunPalette, listRow, itemCode
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

Private Sub EnsureRunProcessMap()
    If mRunProcessByKey Is Nothing Then Set mRunProcessByKey = CreateObject("Scripting.Dictionary")
End Sub

Private Sub EnsureRunItemCodeMap()
    If mRunItemCodeByKey Is Nothing Then Set mRunItemCodeByKey = CreateObject("Scripting.Dictionary")
End Sub

Private Function RunAllocationKeyFromList(ByVal lst As MSForms.ListBox, ByVal rowIndex As Long) As String
    If lst Is Nothing Then Exit Function
    If rowIndex < 0 Then Exit Function
    If IsRunTreeOutputRow(lst, rowIndex) Then Exit Function
    RunAllocationKeyFromList = Trim$(NzStr(lst.List(rowIndex, 0))) & "|" & _
                               Trim$(NzStr(lst.List(rowIndex, 1))) & "|" & _
                               Trim$(NzStr(lst.List(rowIndex, 3))) & "|" & _
                               Trim$(NzStr(lst.List(rowIndex, 4))) & "|" & _
                               Trim$(NzStr(lst.List(rowIndex, 7))) & "|" & _
                               Trim$(NzStr(lst.List(rowIndex, 9)))
End Function

Private Sub StoreRunItemCode(ByVal lst As MSForms.ListBox, ByVal rowIndex As Long, ByVal itemCode As String)
    Dim key As String

    key = RunAllocationKeyFromList(lst, rowIndex)
    itemCode = Trim$(itemCode)
    If key = "" Or itemCode = "" Then Exit Sub
    EnsureRunItemCodeMap
    mRunItemCodeByKey(key) = itemCode
End Sub

Private Function RunItemCodeFromList(ByVal lst As MSForms.ListBox, ByVal rowIndex As Long) As String
    Dim key As String

    If mRunItemCodeByKey Is Nothing Then Exit Function
    key = RunAllocationKeyFromList(lst, rowIndex)
    If key = "" Then Exit Function
    If mRunItemCodeByKey.Exists(key) Then RunItemCodeFromList = NzStr(mRunItemCodeByKey(key))
End Function

Private Sub StoreRunProcess(ByVal lst As MSForms.ListBox, ByVal rowIndex As Long, ByVal processText As String)
    Dim key As String
    key = RunAllocationKeyFromList(lst, rowIndex)
    If key = "" Then Exit Sub
    EnsureRunProcessMap
    mRunProcessByKey(key) = Trim$(processText)
End Sub

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
            If IsRunTreeOutputRow(lst, rowIndex) Then Exit Function
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
    If RunAllocationGroupKeyFromList <> "" Then
        RunAllocationGroupKeyFromList = ProcessKey(RunProcessFromPaletteList(lst, rowIndex)) & "|" & RunAllocationGroupKeyFromList
    End If
End Function

Private Function RunProcessFromPaletteList(ByVal lst As MSForms.ListBox, ByVal rowIndex As Long) As String
    Dim key As String

    If lst Is Nothing Then Exit Function
    If rowIndex < 0 Then Exit Function
    key = RunAllocationKeyFromList(lst, rowIndex)
    If key <> "" Then
        If Not mRunProcessByKey Is Nothing Then
            If mRunProcessByKey.Exists(key) Then
                RunProcessFromPaletteList = Trim$(NzStr(mRunProcessByKey(key)))
                If RunProcessFromPaletteList <> "" Then Exit Function
            End If
        End If
    End If
    If IsNumeric(NzStr(lst.List(rowIndex, 1))) Then Exit Function
    RunProcessFromPaletteList = Trim$(NzStr(lst.List(rowIndex, 0)))
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
    Dim processes As Object
    Dim procKey As Variant
    Dim procName As String
    Dim rowIndex As Long
    Dim collapsed As Boolean

    If mLstRunTree Is Nothing Or mLstRunPalette Is Nothing Then Exit Sub
    EnsureRunTreeState
    mLstRunTree.Clear
    Set processes = OrderedRunProcesses()
    If processes Is Nothing Then Exit Sub

    For Each procKey In processes.Keys
        procName = NzStr(processes(procKey))
        collapsed = RunTreeGroupCollapsed("PROC|" & CStr(procKey))
        mLstRunTree.AddItem RUN_TREE_PARENT_MARKER
        rowIndex = mLstRunTree.ListCount - 1
        mLstRunTree.List(rowIndex, 1) = "PROC|" & CStr(procKey)
        If collapsed Then
            mLstRunTree.List(rowIndex, 2) = "[ SHOW PROCESS ]  " & ProcessCaption(procName)
        Else
            mLstRunTree.List(rowIndex, 2) = "[ HIDE PROCESS ]  " & ProcessCaption(procName)
        End If
        If Not collapsed Then
            AddRunTreeInputRowsForProcess procName
            AddRunTreeOutputRowsForProcess procName
        End If
    Next procKey
End Sub

Private Function OrderedRunProcesses() As Object
    Dim dict As Object
    Dim i As Long
    Dim procName As String
    Dim key As String
    Dim filterProcess As String

    filterProcess = ActiveRunProcess()
    Set dict = CreateObject("Scripting.Dictionary")
    For i = 0 To mLstRunPalette.ListCount - 1
        procName = RunProcessFromPaletteList(mLstRunPalette, i)
        If Not RunProcessMatchesFilter(procName, filterProcess) Then GoTo NextPaletteProcess
        key = ProcessKey(procName)
        If Not dict.Exists(key) Then dict.Add key, procName
NextPaletteProcess:
    Next i
    If Not mLstManagerOutput Is Nothing Then
        For i = 0 To mLstManagerOutput.ListCount - 1
            procName = NzStr(mLstManagerOutput.List(i, 0))
            If Not RunProcessMatchesFilter(procName, filterProcess) Then GoTo NextOutputProcess
            key = ProcessKey(procName)
            If Not dict.Exists(key) Then dict.Add key, procName
NextOutputProcess:
        Next i
    End If
    Set OrderedRunProcesses = dict
End Function

Private Sub AddRunTreeInputRowsForProcess(ByVal procName As String)
    Dim i As Long
    Dim key As String
    Dim lastKey As String
    Dim parentText As String
    Dim childText As String
    Dim rowIndex As Long
    Dim collapsed As Boolean

    For i = 0 To mLstRunPalette.ListCount - 1
        If StrComp(RunProcessFromPaletteList(mLstRunPalette, i), procName, vbTextCompare) <> 0 Then GoTo NextPaletteRow
        key = RunTreeGroupKey(mLstRunPalette, i)
        If key <> lastKey Then
            parentText = NzStr(mLstRunPalette.List(i, 2))
            If parentText = "" Then parentText = "Ingredient"
            collapsed = RunTreeGroupCollapsed(key)
            mLstRunTree.AddItem RUN_TREE_PARENT_MARKER
            rowIndex = mLstRunTree.ListCount - 1
            mLstRunTree.List(rowIndex, 1) = key
            mLstRunTree.List(rowIndex, 2) = RunTreeParentCaption("INPUT  " & parentText, collapsed, key)
            lastKey = key
        End If

        If collapsed Then GoTo NextPaletteRow
        childText = "    " & NzStr(mLstRunPalette.List(i, 4))
        If Trim$(childText) = "" Then childText = "    ROW " & NzStr(mLstRunPalette.List(i, 3))
        mLstRunTree.AddItem NzStr(mLstRunPalette.List(i, 0))
        rowIndex = mLstRunTree.ListCount - 1
        CopyRunPaletteListRow mLstRunPalette, i, mLstRunTree, rowIndex
        mLstRunTree.List(rowIndex, 2) = childText
NextPaletteRow:
    Next i
End Sub

Private Sub AddRunTreeOutputRowsForProcess(ByVal procName As String)
    Dim i As Long
    Dim rowIndex As Long
    Dim outputName As String
    Dim caption As String

    If mLstManagerOutput Is Nothing Then Exit Sub
    For i = 0 To mLstManagerOutput.ListCount - 1
        If StrComp(NzStr(mLstManagerOutput.List(i, 0)), procName, vbTextCompare) <> 0 Then GoTo NextOutput
        outputName = NzStr(mLstManagerOutput.List(i, 1))
        If Trim$(outputName) = "" Then GoTo NextOutput
        caption = "  OUTPUT  " & outputName
        If NzStr(mLstManagerOutput.List(i, 3)) <> "" Then caption = caption & "  Last " & NzStr(mLstManagerOutput.List(i, 3))
        caption = caption & "  Batch " & NzStr(mLstManagerOutput.List(i, 4)) & "  Total " & NzStr(mLstManagerOutput.List(i, 5))
        mLstRunTree.AddItem RUN_TREE_OUTPUT_MARKER
        rowIndex = mLstRunTree.ListCount - 1
        mLstRunTree.List(rowIndex, 1) = "OUTPUT|" & ProcessKey(procName) & "|" & NzStr(mLstManagerOutput.List(i, 7))
        mLstRunTree.List(rowIndex, 2) = caption
        mLstRunTree.List(rowIndex, 3) = NzStr(mLstManagerOutput.List(i, 7))
        mLstRunTree.List(rowIndex, 4) = outputName
        mLstRunTree.List(rowIndex, 5) = NzStr(mLstManagerOutput.List(i, 3))
        mLstRunTree.List(rowIndex, 6) = NzStr(mLstManagerOutput.List(i, 5))
        mLstRunTree.List(rowIndex, 7) = NzStr(mLstManagerOutput.List(i, 2))
        mLstRunTree.List(rowIndex, 8) = "Batch " & NzStr(mLstManagerOutput.List(i, 4))
NextOutput:
    Next i
End Sub

Private Function ProcessKey(ByVal procName As String) As String
    ProcessKey = LCase$(Trim$(procName))
    If ProcessKey = "" Then ProcessKey = "(no process)"
End Function

Private Function ProcessCaption(ByVal procName As String) As String
    ProcessCaption = Trim$(procName)
    If ProcessCaption = "" Then ProcessCaption = "Process"
End Function

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

Private Function IsRunTreeOutputRow(ByVal lst As MSForms.ListBox, ByVal rowIndex As Long) As Boolean
    If lst Is Nothing Then Exit Function
    If rowIndex < 0 Then Exit Function
    IsRunTreeOutputRow = (NzStr(lst.List(rowIndex, 0)) = RUN_TREE_OUTPUT_MARKER)
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
    If Not mRunInventoryCacheLoaded Then
        mRunInventoryRows = RunProduction1("LoadProductionRunInventoryPickerItems", "")
        mRunInventoryCacheLoaded = True
    End If
End Sub

Private Sub RefreshRunLocationChoices()
    Dim dict As Object
    Dim r As Long
    Dim locVal As String
    Dim defaultLoc As String
    Dim selectedLoc As String
    Dim wasLoading As Boolean

    If mCmbRunLocation Is Nothing And mCmbTreeRunLocation Is Nothing Then Exit Sub

    selectedLoc = ActiveRunLocation()
    Set dict = CreateObject("Scripting.Dictionary")
    If Not IsEmpty(mRunInventoryRows) And IsArray(mRunInventoryRows) Then
        For r = LBound(mRunInventoryRows, 1) To UBound(mRunInventoryRows, 1)
            locVal = Trim$(NzStr(mRunInventoryRows(r, 5)))
            AddRunLocationChoice dict, locVal
        Next r
    End If
    AddRunLocationChoice dict, selectedLoc
    defaultLoc = Trim$(NzStr(RunProduction0("GetProductionRunDefaultLocation")))
    AddRunLocationChoice dict, defaultLoc

    wasLoading = mLoading
    mLoading = True
    PopulateRunLocationCombo mCmbRunLocation, dict, selectedLoc
    PopulateRunLocationCombo mCmbTreeRunLocation, dict, selectedLoc
    mLoading = wasLoading
End Sub

Private Sub AddRunLocationChoice(ByVal dict As Object, ByVal locationText As String)
    locationText = Trim$(locationText)
    If dict Is Nothing Or locationText = "" Then Exit Sub
    If Not dict.Exists(LCase$(locationText)) Then dict.Add LCase$(locationText), locationText
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
    If targetCombo.ListCount > 0 Then
        targetCombo.ListIndex = 0
    Else
        targetCombo.Value = ""
    End If
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
    If targetCombo.ListCount > 0 Then
        targetCombo.ListIndex = 0
    Else
        targetCombo.Value = ""
    End If
    For i = 0 To targetCombo.ListCount - 1
        If StrComp(NzStr(targetCombo.List(i)), procVal, vbTextCompare) = 0 Then
            targetCombo.ListIndex = i
            Exit For
        End If
    Next i
    mLoading = wasLoading
End Sub

Private Sub HydrateRunInventoryDisplay(ByVal rowVal As String, ByRef itemCode As String, _
                                       ByRef itemVal As String, _
                                       ByRef uomVal As String, ByRef invVal As String, _
                                       ByRef locVal As String)
    Dim r As Long
    Dim selectedRow As Long
    Dim rowKey As String
    Dim totalVal As String
    Dim rawLoc As String
    Dim preferredLoc As String
    Dim candidateCode As String

    If IsEmpty(mRunInventoryRows) Then Exit Sub
    If Not IsArray(mRunInventoryRows) Then Exit Sub

    rowKey = NormalizeRunRowKey(rowVal)
    itemCode = Trim$(itemCode)
    If rowKey = "" And itemCode = "" Then Exit Sub
    selectedRow = -1
    preferredLoc = ActiveRunLocation()
    For r = LBound(mRunInventoryRows, 1) To UBound(mRunInventoryRows, 1)
        candidateCode = vbNullString
        If UBound(mRunInventoryRows, 2) >= 7 Then candidateCode = Trim$(NzStr(mRunInventoryRows(r, 7)))
        If (rowKey <> "" And NormalizeRunRowKey(NzStr(mRunInventoryRows(r, 1))) = rowKey) _
           Or (itemCode <> "" And StrComp(candidateCode, itemCode, vbTextCompare) = 0) Then
            rawLoc = NzStr(mRunInventoryRows(r, 5))
            If selectedRow < 0 Then selectedRow = r
            If Trim$(rawLoc) <> "" And Trim$(NzStr(mRunInventoryRows(selectedRow, 5))) = "" Then selectedRow = r
            If preferredLoc <> "" Then
                If StrComp(Trim$(rawLoc), preferredLoc, vbTextCompare) = 0 Then
                    selectedRow = r
                    Exit For
                End If
            End If
        End If
    Next r
    If selectedRow < 0 Then Exit Sub

    If Trim$(itemVal) = "" Then itemVal = NzStr(mRunInventoryRows(selectedRow, 2))
    If Trim$(uomVal) = "" Then uomVal = NzStr(mRunInventoryRows(selectedRow, 3))
    If itemCode = "" And UBound(mRunInventoryRows, 2) >= 7 Then _
        itemCode = NzStr(mRunInventoryRows(selectedRow, 7))
    totalVal = NzStr(mRunInventoryRows(selectedRow, 4))
    rawLoc = NzStr(mRunInventoryRows(selectedRow, 5))
    invVal = RunInventoryAvailableDisplay(totalVal, uomVal)
    locVal = rawLoc
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

Private Sub RefreshProductionOutputList(ByVal lst As MSForms.ListBox)
    Dim lo As ListObject
    Dim arr As Variant
    Dim r As Long
    Dim listRow As Long
    Dim cProc As Long
    Dim cOutput As Long
    Dim cUom As Long
    Dim cRecall As Long
    Dim cRow As Long
    Dim cItemCode As Long
    Dim procVal As String
    Dim outputVal As String
    Dim uomVal As String
    Dim recallVal As String
    Dim rowVal As String
    Dim itemCodeVal As String
    Dim identityVal As String
    Dim lastQty As Double
    Dim totalQty As Double
    Dim maxBatch As Long
    Dim loggedCount As Long

    If lst Is Nothing Then Exit Sub
    lst.Clear
    Set lo = ProductionTable(TABLE_MANAGER_OUTPUT)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    cProc = ProductionColumnIndex(lo, "PROCESS")
    cOutput = ProductionColumnIndex(lo, "OUTPUT")
    cUom = ProductionColumnIndex(lo, "UOM")
    cRecall = ProductionColumnIndex(lo, "RECALL CODE")
    cRow = ProductionColumnIndex(lo, "ROW")
    cItemCode = ProductionColumnIndex(lo, "ITEM_CODE")
    If cProc = 0 Or cOutput = 0 Then Exit Sub

    arr = lo.DataBodyRange.Value
    For r = 1 To UBound(arr, 1)
        procVal = NzStr(arr(r, cProc))
        outputVal = NzStr(arr(r, cOutput))
        If Trim$(procVal) = "" And Trim$(outputVal) = "" Then GoTo NextRow
        If cUom > 0 Then uomVal = NzStr(arr(r, cUom)) Else uomVal = ""
        If cRecall > 0 Then recallVal = NzStr(arr(r, cRecall)) Else recallVal = ""
        If cRow > 0 Then rowVal = NzStr(arr(r, cRow)) Else rowVal = ""
        If cItemCode > 0 Then itemCodeVal = NzStr(arr(r, cItemCode)) Else itemCodeVal = ""
        If Trim$(itemCodeVal) <> "" Then
            identityVal = itemCodeVal
        Else
            identityVal = rowVal
        End If

        LoggedOutputStats identityVal, procVal, outputVal, lastQty, totalQty, maxBatch, loggedCount
        lst.AddItem procVal
        listRow = lst.ListCount - 1
        lst.List(listRow, 1) = outputVal
        lst.List(listRow, 2) = uomVal
        lst.List(listRow, 3) = IIf(loggedCount > 0, FormatRunNumber(lastQty), "")
        ' The Batch column is completed-run history: 0 before the first
        ' completion, 1 after the first completion, and so on.  The next
        ' internal batch number is calculated only when Real Output is staged.
        lst.List(listRow, 4) = CStr(maxBatch)
        lst.List(listRow, 5) = IIf(loggedCount > 0, FormatRunNumber(totalQty), "0")
        lst.List(listRow, 6) = recallVal
        lst.List(listRow, 7) = identityVal
NextRow:
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
    mRunInventoryRows = Empty
    mRunInventoryCacheLoaded = False
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
    Dim wb As Workbook
    Dim ws As Worksheet

    Set wb = ResolveOperatorWorkbook()
    If wb Is Nothing Then Exit Function
    If Not WorkbookHasSheet(wb, "Production") Then Exit Function
    Set ws = wb.Worksheets("Production")
    If ws Is Nothing Then Exit Function
    Set ProductionTable = RunProductionObject2("GetListObject", ws, tableName)
End Function

Private Function InventoryTable() As ListObject
    Dim wb As Workbook
    Dim ws As Worksheet

    Set wb = ResolveOperatorWorkbook()
    If wb Is Nothing Then Exit Function
    If Not WorkbookHasSheet(wb, "InventoryManagement") Then Exit Function
    Set ws = wb.Worksheets("InventoryManagement")
    If ws Is Nothing Then Exit Function
    Set InventoryTable = RunProductionObject2("GetListObject", ws, "invSys")
End Function

Private Function GenerateRecipeIdForOperatorWorkbook() As String
    Dim wb As Workbook

    Set wb = ResolveOperatorWorkbook()
    If wb Is Nothing Then Exit Function
    GenerateRecipeIdForOperatorWorkbook = NzStr(RunProduction1("GenerateRecipeId", wb))
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

Private Function RefreshProductionInventoryReadModel(ByRef reportOut As String) As Boolean
    On Error GoTo CleanFail

    Dim wb As Workbook
    Dim resultText As String
    Dim separatorAt As Long
    Dim resultCode As String

    Set wb = ResolveOperatorWorkbook()
    If wb Is Nothing Then
        reportOut = "Open a Production operator workbook before refreshing inventory."
        Exit Function
    End If

    resultText = NzStr(RunProduction1("RefreshProductionInventoryReadModelForWorkbookResult", wb))
    separatorAt = InStr(1, resultText, vbTab, vbBinaryCompare)
    If separatorAt > 0 Then
        resultCode = Left$(resultText, separatorAt - 1)
        reportOut = Mid$(resultText, separatorAt + 1)
    Else
        resultCode = resultText
        reportOut = resultText
    End If
    If Trim$(reportOut) = "" Then reportOut = IIf(resultCode = "OK", "OK", "Inventory refresh failed.")
    RefreshProductionInventoryReadModel = (StrComp(Trim$(resultCode), "OK", vbTextCompare) = 0)
    Exit Function

CleanFail:
    reportOut = "Production inventory refresh failed: " & Err.Description
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
    If Trim$(mTxtRecipeId.Text) = "" Then mTxtRecipeId.Text = GenerateRecipeIdForOperatorWorkbook()
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
    mTxtRecipeId.Text = GenerateRecipeIdForOperatorWorkbook()
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
    mTxtLineUom.Value = ""
    mTxtLineAmount.Text = ""
    ConfigureRecipeLineInputs
End Sub

Private Sub LoadSelectedBuilderLine()
    Dim idx As Long

    idx = mLstBuilderLines.ListIndex
    If idx < 0 Then Exit Sub
    mTxtLineProcess.Text = NzStr(mLstBuilderLines.List(idx, 0))
    SetLineIo NzStr(mLstBuilderLines.List(idx, 2))
    mTxtLineIngredient.Text = NzStr(mLstBuilderLines.List(idx, 3))
    mTxtLinePercent.Text = NzStr(mLstBuilderLines.List(idx, 4))
    RefreshRecipeUomCatalog NzStr(mLstBuilderLines.List(idx, 5))
    mTxtLineAmount.Text = NzStr(mLstBuilderLines.List(idx, 6))
End Sub

Private Sub SetLineIo(ByVal ioValue As String)
    Dim i As Long
    Dim v As String

    v = UCase$(Trim$(ioValue))
    If v = "MADE" Then v = "OUTPUT"
    For i = 0 To mCmbLineIo.ListCount - 1
        If StrComp(NzStr(mCmbLineIo.List(i)), v, vbTextCompare) = 0 Then
            mCmbLineIo.ListIndex = i
            ConfigureRecipeLineInputs
            Exit Sub
        End If
    Next i
    mCmbLineIo.ListIndex = 0
    ConfigureRecipeLineInputs
End Sub

Private Function LineIoValue() As String
    If mCmbLineIo.ListIndex >= 0 Then LineIoValue = UCase$(Trim$(NzStr(mCmbLineIo.Value)))
    If LineIoValue = "" Then LineIoValue = "USED"
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
    ResequenceRecipeBuilderLines lo
    RefreshBuilderLines
    ClearLineInputs
    ShowStatus "Recipe line added."
End Sub

Private Sub UpdateSelectedRecipeBuilderLine()
    If WriteSelectedRecipeBuilderLineFromForm(True) Then
        ShowStatus "Recipe line updated."
    End If
End Sub

Private Function WriteSelectedRecipeBuilderLineFromForm(Optional ByVal refreshList As Boolean = True) As Boolean
    Dim lo As ListObject
    Dim idx As Long
    Dim tableRow As Long

    idx = mLstBuilderLines.ListIndex
    If idx < 0 Then
        If refreshList Then ShowStatus "Select a recipe line to update."
        Exit Function
    End If
    tableRow = BuilderTableRowForListIndex(idx)
    If tableRow <= 0 Then Exit Function
    Set lo = ProductionTable(TABLE_BUILDER_LINES)
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    If tableRow > lo.ListRows.Count Then Exit Function
    WriteRecipeLineToRow lo, tableRow
    If refreshList Then
        RefreshBuilderLines
        If idx < mLstBuilderLines.ListCount Then mLstBuilderLines.ListIndex = idx
    End If
    WriteSelectedRecipeBuilderLineFromForm = True
End Function

Private Sub RemoveSelectedRecipeBuilderLine()
    Dim lo As ListObject
    Dim idx As Long
    Dim tableRow As Long

    idx = mLstBuilderLines.ListIndex
    If idx < 0 Then
        ShowStatus "Select a recipe line to remove."
        Exit Sub
    End If
    tableRow = BuilderTableRowForListIndex(idx)
    If tableRow <= 0 Then Exit Sub
    Set lo = ProductionTable(TABLE_BUILDER_LINES)
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    If tableRow <= lo.ListRows.Count Then lo.ListRows(tableRow).Delete
    ResequenceRecipeBuilderLines lo
    RefreshBuilderLines
    ClearLineInputs
    ShowStatus "Recipe line removed."
End Sub

Private Sub WriteRecipeLineToRow(ByVal lo As ListObject, ByVal rowIndex As Long)
    Dim ioType As String

    If lo Is Nothing Then Exit Sub
    If rowIndex < 1 Then Exit Sub
    If Not EnsureTableRows(lo, rowIndex) Then Exit Sub
    ioType = LineIoValue()
    SetCellByHeader lo, rowIndex, "PROCESS", mTxtLineProcess.Text
    SetCellByHeader lo, rowIndex, "INPUT/OUTPUT", ioType
    SetCellByHeader lo, rowIndex, "INGREDIENT", mTxtLineIngredient.Text
    If ioType = "INSTRUCTION" Then
        SetCellByHeader lo, rowIndex, "PERCENT", vbNullString
        SetCellByHeader lo, rowIndex, "UOM", vbNullString
        SetCellByHeader lo, rowIndex, "AMOUNT", vbNullString
        SetCellByHeader lo, rowIndex, "INGREDIENT_ID", vbNullString
    Else
        SetCellByHeader lo, rowIndex, "PERCENT", mTxtLinePercent.Text
        SetCellByHeader lo, rowIndex, "UOM", CStr(mTxtLineUom.Value)
        SetCellByHeader lo, rowIndex, "AMOUNT", mTxtLineAmount.Text
        If Trim$(CellByHeader(lo, rowIndex, "INGREDIENT_ID")) = "" Then SetCellByHeader lo, rowIndex, "INGREDIENT_ID", BuildFormGuid()
    End If
    If Trim$(CellByHeader(lo, rowIndex, "RECIPE_LIST_ROW")) = "" Then SetCellByHeader lo, rowIndex, "RECIPE_LIST_ROW", rowIndex
    If Trim$(CellByHeader(lo, rowIndex, "GUID")) = "" Then SetCellByHeader lo, rowIndex, "GUID", BuildFormGuid()
End Sub

Private Sub ConfigureRecipeLineInputs()
    Dim acceptsQuantity As Boolean

    acceptsQuantity = (LineIoValue() <> "INSTRUCTION")
    If Not mTxtLinePercent Is Nothing Then mTxtLinePercent.Enabled = acceptsQuantity
    If Not mTxtLineUom Is Nothing Then mTxtLineUom.Enabled = acceptsQuantity
    If Not mBtnLineUomAdd Is Nothing Then mBtnLineUomAdd.Enabled = acceptsQuantity
    If Not mTxtLineAmount Is Nothing Then mTxtLineAmount.Enabled = acceptsQuantity
End Sub

Private Function MoveSelectedRecipeBuilderLine(ByVal moveDirection As Long, _
                                               Optional ByVal showMessages As Boolean = True) As Boolean
    Dim lo As ListObject
    Dim selectedRow As Long
    Dim targetRow As Long
    Dim selectedListIndex As Long
    Dim targetListIndex As Long
    Dim selectedValues() As Variant
    Dim targetValues() As Variant
    Dim columnIndex As Long
    Dim columnCount As Long

    If moveDirection <> -1 And moveDirection <> 1 Then Exit Function
    If mLstBuilderLines Is Nothing Then Exit Function
    If mLstBuilderLines.ListIndex < 0 Then
        If showMessages Then ShowStatus "Select a recipe line to move."
        Exit Function
    End If

    Set lo = ProductionTable(TABLE_BUILDER_LINES)
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    selectedListIndex = mLstBuilderLines.ListIndex
    targetListIndex = selectedListIndex + moveDirection
    selectedRow = BuilderTableRowForListIndex(selectedListIndex)
    targetRow = BuilderTableRowForListIndex(targetListIndex)
    If selectedRow <= 0 Or targetRow <= 0 Then
        If showMessages Then ShowStatus IIf(moveDirection < 0, "The selected line is already first.", "The selected line is already last.")
        Exit Function
    End If

    columnCount = lo.ListColumns.Count
    ReDim selectedValues(1 To columnCount)
    ReDim targetValues(1 To columnCount)
    For columnIndex = 1 To columnCount
        selectedValues(columnIndex) = lo.DataBodyRange.Cells(selectedRow, columnIndex).Value2
        targetValues(columnIndex) = lo.DataBodyRange.Cells(targetRow, columnIndex).Value2
    Next columnIndex
    For columnIndex = 1 To columnCount
        lo.DataBodyRange.Cells(selectedRow, columnIndex).Value2 = targetValues(columnIndex)
        lo.DataBodyRange.Cells(targetRow, columnIndex).Value2 = selectedValues(columnIndex)
    Next columnIndex
    ResequenceRecipeBuilderLines lo
    RefreshBuilderLines
    If targetListIndex >= 0 And targetListIndex < mLstBuilderLines.ListCount Then _
        mLstBuilderLines.ListIndex = targetListIndex
    LoadSelectedBuilderLine
    If showMessages Then ShowStatus "Recipe line moved " & IIf(moveDirection < 0, "up.", "down.")
    MoveSelectedRecipeBuilderLine = True
End Function

Private Sub ResequenceRecipeBuilderLines(ByVal lo As ListObject)
    Dim rowIndex As Long

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    For rowIndex = 1 To lo.ListRows.Count
        SetCellByHeader lo, rowIndex, "RECIPE_LIST_ROW", rowIndex
    Next rowIndex
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
    Dim ioVal As String
    Dim lo As ListObject

    idx = mLstAssignIngredients.ListIndex
    If idx < 0 Then
        ShowStatus "Select an ingredient first."
        Exit Sub
    End If
    ioVal = UCase$(Trim$(NzStr(mLstAssignIngredients.List(idx, 4))))
    If ioVal = "OUTPUT" Or ioVal = "MADE" Then
        ClearAssignmentIngredientSelection
        ShowStatus "Outputs do not accept ingredients. Assign acceptable inventory only to recipe inputs."
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
    On Error GoTo FailAdd

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
    If CBool(RunProduction2("PaletteIngredientIsOutput", recipeId, ingredientId)) Then
        ClearAssignmentIngredientSelection
        ShowStatus "Outputs do not accept ingredients. Assign acceptable inventory only to recipe inputs."
        Exit Sub
    End If
    Set lo = ProductionTable(TABLE_ASSIGN_ITEM)
    If lo Is Nothing Then
        ShowStatus "IP_ChooseItem table is missing."
        Exit Sub
    End If
    Set lr = WritableAssignmentItemRow(lo)
    If lr Is Nothing Then
        ShowStatus "Acceptable inventory could not be staged because the operator table has no available row."
        Exit Sub
    End If
    SetRowValue lr, lo, "ROW", NzStr(mLstAssignInventory.List(idx, 0))
    SetRowValue lr, lo, "ITEM_CODE", NzStr(mLstAssignInventory.List(idx, 6))
    SetRowValue lr, lo, "ITEMS", NzStr(mLstAssignInventory.List(idx, 1))
    SetRowValue lr, lo, "UOM", NzStr(mLstAssignInventory.List(idx, 2))
    SetRowValue lr, lo, "DESCRIPTION", NzStr(mLstAssignInventory.List(idx, 5))
    SetRowValue lr, lo, "RECIPE_ID", recipeId
    SetRowValue lr, lo, "INGREDIENT_ID", ingredientId
    RefreshAllowedItems
    ShowStatus "Added acceptable ingredient row " & NzStr(mLstAssignInventory.List(idx, 0)) & "."
    Exit Sub

FailAdd:
    ShowStatus "Add Acceptable failed: " & Err.Description
End Sub

Private Function WritableAssignmentItemRow(ByVal lo As ListObject) As ListRow
    Dim r As Long
    Dim priorCount As Long
    Dim insertRow As Long

    If lo Is Nothing Then Exit Function
    If Not lo.DataBodyRange Is Nothing Then
        For r = 1 To lo.ListRows.Count
            If AssignmentItemRowIsBlank(lo, r) Then
                Set WritableAssignmentItemRow = lo.ListRows(r)
                Exit Function
            End If
        Next r
    End If

    On Error Resume Next
    Set WritableAssignmentItemRow = lo.ListRows.Add(AlwaysInsert:=False)
    On Error GoTo 0
    If Not WritableAssignmentItemRow Is Nothing Then Exit Function

    On Error GoTo FailAcquire
    priorCount = lo.ListRows.Count
    insertRow = lo.Range.Row + lo.Range.Rows.Count
    lo.Parent.Rows(insertRow).Insert Shift:=xlDown
    If lo.ListRows.Count > priorCount Then
        Set WritableAssignmentItemRow = lo.ListRows(lo.ListRows.Count)
    Else
        Set WritableAssignmentItemRow = lo.ListRows.Add(AlwaysInsert:=False)
    End If
    Exit Function

FailAcquire:
    Set WritableAssignmentItemRow = Nothing
End Function

Private Function AssignmentItemRowIsBlank(ByVal lo As ListObject, ByVal rowIndex As Long) As Boolean
    AssignmentItemRowIsBlank = _
        Trim$(CellByHeader(lo, rowIndex, "ROW")) = "" _
        And Trim$(CellByHeader(lo, rowIndex, "ITEM_CODE")) = "" _
        And Trim$(CellByHeader(lo, rowIndex, "ITEMS")) = "" _
        And Trim$(CellByHeader(lo, rowIndex, "RECIPE_ID")) = "" _
        And Trim$(CellByHeader(lo, rowIndex, "INGREDIENT_ID")) = ""
End Function

Private Sub ClearAssignmentIngredientSelection()
    Dim loIng As ListObject
    Dim loItems As ListObject

    Set loIng = ProductionTable(TABLE_ASSIGN_INGREDIENT)
    Set loItems = ProductionTable(TABLE_ASSIGN_ITEM)
    ClearTableContentsKeepBlank loIng
    ClearTableContentsKeepBlank loItems
    RefreshAllowedItems
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
    Dim lo As ListObject
    Dim tableRowNumber As Long

    idx = mLstManagerOutput.ListIndex
    If idx < 0 Then Exit Sub
    Set lo = ProductionTable(TABLE_MANAGER_OUTPUT)
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    tableRowNumber = SelectedProductionOutputTableRow()
    If tableRowNumber = 0 Then Exit Sub
    mTxtOutputReal.Text = CellByHeader(lo, tableRowNumber, "REAL OUTPUT")
End Sub

Private Sub ApplySelectedProductionOutput()
    Dim idx As Long
    Dim lo As ListObject
    Dim batchVal As String
    Dim tableRowNumber As Long

    idx = mLstManagerOutput.ListIndex
    If idx < 0 Then
        ShowStatus "Select a ProductionOutput row first."
        Exit Sub
    End If
    Set lo = ProductionTable(TABLE_MANAGER_OUTPUT)
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    tableRowNumber = SelectedProductionOutputTableRow()
    If tableRowNumber = 0 Then
        ShowStatus "The selected Production Output row could not be resolved."
        Exit Sub
    End If
    batchVal = CStr(NextOutputBatchNumberForListIndex(idx))
    SetCellByHeader lo, tableRowNumber, "REAL OUTPUT", mTxtOutputReal.Text
    SetCellByHeader lo, tableRowNumber, "BATCH", batchVal
    RefreshManagerState
    If idx < mLstManagerOutput.ListCount Then mLstManagerOutput.ListIndex = idx
    ShowStatus "Real Output staged for batch " & batchVal & "."
End Sub

Private Function SelectedProductionOutputTableRow() As Long
    Dim idx As Long
    Dim lo As ListObject

    If mLstManagerOutput Is Nothing Then Exit Function
    idx = mLstManagerOutput.ListIndex
    If idx < 0 Or idx >= mLstManagerOutput.ListCount Then Exit Function
    Set lo = ProductionTable(TABLE_MANAGER_OUTPUT)
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function

    SelectedProductionOutputTableRow = FindProductionOutputTableRow( _
        lo, _
        NzStr(mLstManagerOutput.List(idx, 7)), _
        NzStr(mLstManagerOutput.List(idx, 0)), _
        NzStr(mLstManagerOutput.List(idx, 1)), _
        idx + 1)
End Function

Private Function FindProductionOutputTableRow(ByVal lo As ListObject, _
                                              ByVal rowVal As String, _
                                              ByVal procVal As String, _
                                              ByVal outputVal As String, _
                                              Optional ByVal fallbackRow As Long = 0) As Long
    Dim r As Long
    Dim cRow As Long
    Dim cProc As Long
    Dim cOutput As Long
    Dim wantedRow As String

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    cRow = ProductionColumnIndex(lo, "ROW")
    cProc = ProductionColumnIndex(lo, "PROCESS")
    cOutput = ProductionColumnIndex(lo, "OUTPUT")
    wantedRow = NormalizeRunRowKey(rowVal)

    If wantedRow <> "" And cRow > 0 Then
        For r = 1 To lo.ListRows.Count
            If NormalizeRunRowKey(NzStr(lo.DataBodyRange.Cells(r, cRow).Value)) = wantedRow Then
                FindProductionOutputTableRow = r
                Exit Function
            End If
        Next r
    End If

    If cProc > 0 And cOutput > 0 Then
        For r = 1 To lo.ListRows.Count
            If StrComp(Trim$(NzStr(lo.DataBodyRange.Cells(r, cProc).Value)), Trim$(procVal), vbTextCompare) = 0 _
               And StrComp(Trim$(NzStr(lo.DataBodyRange.Cells(r, cOutput).Value)), Trim$(outputVal), vbTextCompare) = 0 Then
                FindProductionOutputTableRow = r
                Exit Function
            End If
        Next r
    End If

    If fallbackRow >= 1 And fallbackRow <= lo.ListRows.Count Then
        FindProductionOutputTableRow = fallbackRow
    End If
End Function

Private Function NextOutputBatchNumberForListIndex(ByVal outputIndex As Long) As Long
    Dim lastQty As Double
    Dim totalQty As Double
    Dim maxBatch As Long
    Dim loggedCount As Long

    If mLstManagerOutput Is Nothing Then Exit Function
    If outputIndex < 0 Or outputIndex >= mLstManagerOutput.ListCount Then Exit Function
    LoggedOutputStats NzStr(mLstManagerOutput.List(outputIndex, 7)), _
                      NzStr(mLstManagerOutput.List(outputIndex, 0)), _
                      NzStr(mLstManagerOutput.List(outputIndex, 1)), _
                      lastQty, totalQty, maxBatch, loggedCount
    NextOutputBatchNumberForListIndex = maxBatch + 1
End Function

Private Sub LoggedOutputStats(ByVal rowVal As String, ByVal procVal As String, ByVal outputVal As String, _
                              ByRef lastQty As Double, ByRef totalQty As Double, _
                              ByRef maxBatch As Long, ByRef loggedCount As Long)
    Dim loLog As ListObject
    Dim r As Long
    Dim cReal As Long
    Dim cBatch As Long
    Dim cTime As Long
    Dim cProc As Long
    Dim cItem As Long
    Dim cOutput As Long
    Dim cRow As Long
    Dim realVal As Double
    Dim batchVal As Long
    Dim timeVal As Variant
    Dim latestTime As Date
    Dim hasLatestTime As Boolean

    lastQty = 0
    totalQty = 0
    maxBatch = 0
    loggedCount = 0
    Set loLog = ProductionLogTable()
    If loLog Is Nothing Then Exit Sub
    If loLog.DataBodyRange Is Nothing Then Exit Sub

    cReal = ProductionColumnIndex(loLog, "REAL OUTPUT")
    If cReal = 0 Then Exit Sub
    cBatch = ProductionColumnIndex(loLog, "BATCH")
    cTime = ProductionColumnIndex(loLog, "TIMESTAMP")
    cProc = ProductionColumnIndex(loLog, "PROCESS")
    cItem = ProductionColumnIndex(loLog, "ITEM")
    cOutput = ProductionColumnIndex(loLog, "OUTPUT")
    cRow = ProductionColumnIndex(loLog, "ROW")

    For r = 1 To loLog.ListRows.Count
        If Not ProductionLogRowMatchesOutput(loLog, r, rowVal, procVal, outputVal, cRow, cProc, cItem, cOutput) Then GoTo NextRow
        If IsNumeric(loLog.DataBodyRange.Cells(r, cReal).Value) Then
            realVal = CDbl(loLog.DataBodyRange.Cells(r, cReal).Value)
            totalQty = totalQty + realVal
            loggedCount = loggedCount + 1
            If cBatch > 0 Then
                batchVal = CLng(Val(loLog.DataBodyRange.Cells(r, cBatch).Value))
                If batchVal > maxBatch Then maxBatch = batchVal
            End If
            If cTime > 0 Then timeVal = loLog.DataBodyRange.Cells(r, cTime).Value Else timeVal = Empty
            If IsDate(timeVal) Then
                If Not hasLatestTime Or CDate(timeVal) >= latestTime Then
                    latestTime = CDate(timeVal)
                    hasLatestTime = True
                    lastQty = realVal
                End If
            ElseIf Not hasLatestTime Then
                lastQty = realVal
            End If
        End If
NextRow:
    Next r
End Sub

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
    ClearMismatchedRunLocationAllocations
    If Not ValidateRunAllocationsComplete() Then Exit Sub
    If Not ValidateRunAllocationLocations() Then Exit Sub

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
    Dim outputRowNumber As Long
    Dim outputRowVal As String
    Dim outputProcess As String
    Dim outputName As String
    Dim processName As String
    Dim prepared As Variant
    Dim completionReport As String
    Dim completionResult As String
    Dim reportSeparator As Long
    Dim enteredRealOutput As String
    Dim lo As ListObject

    If Not HasProductionCheckRows() Then
        ShowStatus "Check inventory into Production before completing the run."
        Exit Sub
    End If
    If Not ValidateRunAllocationLocations() Then Exit Sub

    enteredRealOutput = Trim$(mTxtOutputReal.Text)
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
        If Trim$(mTxtOutputReal.Text) = "" And enteredRealOutput <> "" Then mTxtOutputReal.Text = enteredRealOutput
    End If

    outputIndex = mLstManagerOutput.ListIndex
    If outputIndex < 0 And mLstManagerOutput.ListCount = 1 Then
        mLstManagerOutput.ListIndex = 0
        LoadSelectedProductionOutput
        If Trim$(mTxtOutputReal.Text) = "" And enteredRealOutput <> "" Then mTxtOutputReal.Text = enteredRealOutput
        outputIndex = 0
    End If
    If outputIndex < 0 Then
        ShowStatus "Select the Production Output row for this run."
        Exit Sub
    End If
    If Trim$(mTxtOutputReal.Text) = "" Then
        ShowStatus "Enter the real output quantity before completing the run."
        Exit Sub
    End If
    If Not IsNumeric(mTxtOutputReal.Text) Or CDbl(mTxtOutputReal.Text) <= 0 Then
        ShowStatus "Real Output must be a number greater than zero."
        Exit Sub
    End If

    outputRowNumber = SelectedProductionOutputTableRow()
    If outputRowNumber = 0 Then
        ShowStatus "The selected Production Output row could not be resolved. Refresh Production Run and select the output again."
        Exit Sub
    End If
    outputRowVal = NzStr(mLstManagerOutput.List(outputIndex, 7))
    outputProcess = NzStr(mLstManagerOutput.List(outputIndex, 0))
    outputName = NzStr(mLstManagerOutput.List(outputIndex, 1))

    ApplySelectedProductionOutput
    completionResult = CStr(Application.Run("mProduction.CompleteProductionRunAfterCheckInForOutputResult", outputRowNumber))
    reportSeparator = InStr(1, completionResult, vbTab, vbBinaryCompare)
    If reportSeparator > 0 Then
        completionReport = Mid$(completionResult, reportSeparator + 1)
        completionResult = Left$(completionResult, reportSeparator - 1)
    Else
        completionReport = completionResult
    End If
    If StrComp(completionResult, "OK", vbTextCompare) <> 0 Then
        If Trim$(completionReport) = "" Then completionReport = "Complete Run failed."
        ShowStatus completionReport
        MsgBox completionReport, vbExclamation, "Production Complete Run"
        Exit Sub
    End If
    ResetInventoryCache
    RefreshLoaderState
    RefreshManagerState
    Set lo = ProductionTable(TABLE_MANAGER_OUTPUT)
    outputRowNumber = FindProductionOutputTableRow(lo, outputRowVal, outputProcess, outputName, outputRowNumber)
    ClearProductionOutputEntry outputRowNumber
    mTxtOutputReal.Text = ""
    RefreshManagerState
    ShowStatus "Production run completed. Checked-in inventory was consumed, Real Output was added to inventory, and the batch was logged." & IIf(Trim$(completionReport) <> "", " " & completionReport, "")
End Sub

Private Sub ClearProductionOutputEntry(ByVal outputRowNumber As Long)
    Dim lo As ListObject

    Set lo = ProductionTable(TABLE_MANAGER_OUTPUT)
    If lo Is Nothing Then Exit Sub
    If outputRowNumber < 1 Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If outputRowNumber > lo.ListRows.Count Then Exit Sub
    SetCellByHeader lo, outputRowNumber, "REAL OUTPUT", ""
    SetCellByHeader lo, outputRowNumber, "BATCH", ""
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
    Dim cItemCode As Long
    Dim cUsed As Long
    Dim rowIdentity As String

    cRow = ProductionColumnIndex(lo, "ROW")
    cItemCode = ProductionColumnIndex(lo, "ITEM_CODE")
    cUsed = ProductionColumnIndex(lo, "USED")
    If (cRow = 0 And cItemCode = 0) Or cUsed = 0 Then Exit Function
    For r = 1 To lo.ListRows.Count
        rowIdentity = ""
        If cItemCode > 0 Then rowIdentity = Trim$(NzStr(lo.DataBodyRange.Cells(r, cItemCode).Value))
        If rowIdentity = "" And cRow > 0 Then rowIdentity = Trim$(NzStr(lo.DataBodyRange.Cells(r, cRow).Value))
        If rowIdentity <> "" And NzDblLocal(lo.DataBodyRange.Cells(r, cUsed).Value) > 0 Then
            FirstNonBlankCheckRow = rowIdentity
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

Private Function ValidateRunAllocationLocations() As Boolean
    Dim runLoc As String
    Dim i As Long
    Dim qtyVal As Double
    Dim itemVal As String
    Dim invLoc As String

    If mLstRunPalette Is Nothing Then Exit Function
    runLoc = ActiveRunLocation()
    If RunLocationChoiceRequired() And runLoc = "" Then
        ShowStatus "Choose a production run location before checking inventory into Production."
        Exit Function
    End If

    For i = 0 To mLstRunPalette.ListCount - 1
        If Not IsNumeric(NzStr(mLstRunPalette.List(i, 6))) Then GoTo NextChoice
        qtyVal = CDbl(NzStr(mLstRunPalette.List(i, 6)))
        If qtyVal <= 0 Then GoTo NextChoice
        invLoc = Trim$(NzStr(mLstRunPalette.List(i, 9)))
        If Not RunChoiceLocationAllowed(runLoc, invLoc, qtyVal) Then
            itemVal = Trim$(NzStr(mLstRunPalette.List(i, 4)))
            If itemVal = "" Then itemVal = "ROW " & Trim$(NzStr(mLstRunPalette.List(i, 3)))
            ShowStatus "Cannot use " & itemVal & " from " & invLoc & ". Production run location is " & runLoc & "; use inventory at the production location."
            Exit Function
        End If
NextChoice:
    Next i

    ValidateRunAllocationLocations = True
End Function

Private Function ClearMismatchedRunLocationAllocations(Optional ByRef clearedCount As Long = 0) As Boolean
    Dim runLoc As String
    Dim i As Long
    Dim qtyVal As Double
    Dim splitVal As Double
    Dim hasPositiveAllocation As Boolean
    Dim invLoc As String

    clearedCount = 0
    If mLstRunPalette Is Nothing Then
        ClearMismatchedRunLocationAllocations = True
        Exit Function
    End If
    runLoc = ActiveRunLocation()
    If runLoc = "" Then
        ClearMismatchedRunLocationAllocations = True
        Exit Function
    End If

    For i = 0 To mLstRunPalette.ListCount - 1
        qtyVal = 0
        splitVal = 0
        If IsNumeric(NzStr(mLstRunPalette.List(i, 6))) Then qtyVal = CDbl(NzStr(mLstRunPalette.List(i, 6)))
        If IsNumeric(NzStr(mLstRunPalette.List(i, 5))) Then splitVal = CDbl(NzStr(mLstRunPalette.List(i, 5)))
        hasPositiveAllocation = (qtyVal > 0 Or splitVal > 0)
        If hasPositiveAllocation Then
            invLoc = Trim$(NzStr(mLstRunPalette.List(i, 9)))
            If Not RunChoiceLocationAllowed(runLoc, invLoc, IIf(qtyVal > 0, qtyVal, splitVal)) Then
                ClearRunAllocationForListRow mLstRunPalette, i
                clearedCount = clearedCount + 1
            End If
        End If
    Next i

    If clearedCount > 0 Then BuildRunTreeFromPaletteList
    ClearMismatchedRunLocationAllocations = True
End Function

Private Sub ClearRunAllocationForListRow(ByVal lst As MSForms.ListBox, ByVal rowIndex As Long)
    Dim tableName As String
    Dim tableRow As Long
    Dim lo As ListObject

    If lst Is Nothing Then Exit Sub
    If rowIndex < 0 Or rowIndex >= lst.ListCount Then Exit Sub

    tableName = NzStr(lst.List(rowIndex, 0))
    tableRow = CLng(Val(NzStr(lst.List(rowIndex, 1))))
    If tableName <> "" And tableRow >= 1 Then
        Set lo = ProductionTable(tableName)
        If Not lo Is Nothing Then
            If Not lo.DataBodyRange Is Nothing Then
                If tableRow <= lo.ListRows.Count Then
                    SetCellByHeader lo, tableRow, "SPLIT %", ""
                    SetCellByHeader lo, tableRow, "QUANTITY", ""
                End If
            End If
        End If
    End If

    lst.List(rowIndex, 5) = ""
    lst.List(rowIndex, 6) = ""
    StoreRunAllocationOverride lst, rowIndex, "", ""
    SyncRunAllocationToPaletteList lst, rowIndex
End Sub

Private Function RunChoiceLocationAllowed(ByVal runLocation As String, ByVal inventoryLocation As String, ByVal qtyValue As Double) As Boolean
    runLocation = Trim$(runLocation)
    inventoryLocation = Trim$(inventoryLocation)
    If qtyValue <= 0 Then
        RunChoiceLocationAllowed = True
    ElseIf runLocation = "" Or inventoryLocation = "" Then
        RunChoiceLocationAllowed = False
    Else
        RunChoiceLocationAllowed = (StrComp(runLocation, inventoryLocation, vbTextCompare) = 0)
    End If
End Function

Private Function BuildRunUsedPayloadJson(ByRef stagedTotal As Double) As String
    Dim agg As Object
    Dim i As Long
    Dim rowVal As String
    Dim systemKey As String
    Dim itemCode As String
    Dim identityKey As String
    Dim qtyVal As Double
    Dim locVal As String
    Dim key As Variant
    Dim payloadItems As Collection
    Dim payloadItem As Object

    stagedTotal = 0
    Set agg = CreateObject("Scripting.Dictionary")
    For i = 0 To mLstRunPalette.ListCount - 1
        rowVal = Trim$(NzStr(mLstRunPalette.List(i, 3)))
        itemCode = Trim$(RunItemCodeFromList(mLstRunPalette, i))
        locVal = NzStr(mLstRunPalette.List(i, 9))
        systemKey = ResolveRunSystemKey(itemCode, NzStr(mLstRunPalette.List(i, 4)), locVal)
        If systemKey = "" Then
            ShowStatus "Cannot complete run. The selected inventory row has no immutable System_Key."
            Exit Function
        End If
        identityKey = "SYS|" & systemKey
        If Not IsNumeric(NzStr(mLstRunPalette.List(i, 6))) Then GoTo NextChoice
        qtyVal = CDbl(NzStr(mLstRunPalette.List(i, 6)))
        If qtyVal <= 0 Then GoTo NextChoice
        If RunChoiceWouldExceedInventory(i, qtyVal) Then
            ShowStatus "Cannot complete run. Inventory " & IIf(rowVal <> "", "ROW " & rowVal, itemCode) & _
                       " requires " & FormatRunNumber(qtyVal) & " but only " & _
                       NzStr(mLstRunPalette.List(i, 8)) & " is available."
            Exit Function
        End If
        If agg.Exists(identityKey) Then
            Dim existingAgg As Variant
            existingAgg = agg(identityKey)
            agg(identityKey) = Array(CDbl(existingAgg(0)) + qtyVal, locVal, systemKey, itemCode)
        Else
            agg.Add identityKey, Array(qtyVal, locVal, systemKey, itemCode)
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
        payloadItem("System_Key") = NzStr(agg(key)(2))
        payloadItem("SKU") = NzStr(agg(key)(3))
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
    Dim systemKey As String
    Dim itemCode As String
    Dim identityKey As String
    Dim qtyVal As Double
    Dim entry As Variant
    Dim key As Variant
    Dim outRow As Long

    Set lo = ProductionTable(TABLE_MANAGER_CHECK)
    If lo Is Nothing Then Exit Function
    Set agg = CreateObject("Scripting.Dictionary")

    For i = 0 To mLstRunPalette.ListCount - 1
        rowVal = Trim$(NzStr(mLstRunPalette.List(i, 3)))
        itemCode = Trim$(RunItemCodeFromList(mLstRunPalette, i))
        systemKey = ResolveRunSystemKey(itemCode, NzStr(mLstRunPalette.List(i, 4)), _
                                        NzStr(mLstRunPalette.List(i, 9)))
        If systemKey = "" Then Exit Function
        identityKey = "SYS|" & systemKey
        If Not IsNumeric(NzStr(mLstRunPalette.List(i, 6))) Then GoTo NextChoice
        qtyVal = CDbl(NzStr(mLstRunPalette.List(i, 6)))
        If qtyVal <= 0 Then GoTo NextChoice
        If agg.Exists(identityKey) Then
            entry = agg(identityKey)
            entry(5) = CDbl(entry(5)) + qtyVal
            agg(identityKey) = entry
        Else
            agg.Add identityKey, Array( _
                systemKey, _
                itemCode, _
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
        SetCellByHeader lo, outRow, "System_Key", NzStr(entry(0))
        SetCellByHeader lo, outRow, "ITEM_CODE", NzStr(entry(1))
        SetCellByHeader lo, outRow, "ITEM", NzStr(entry(2))
        SetCellByHeader lo, outRow, "UOM", NzStr(entry(3))
        SetCellByHeader lo, outRow, "USED", CDbl(entry(5))
        SetCellByHeader lo, outRow, "TOTAL INV", NzStr(entry(4))
    Next key
    WriteProductionCheckRowsFromRunPalette = True
End Function

Private Function ResolveRunSystemKey(ByVal itemCode As String, _
                                     ByVal itemName As String, _
                                     ByVal locationValue As String) As String
    Dim lo As ListObject
    Dim cSystemKey As Long
    Dim cItemCode As Long
    Dim cItem As Long
    Dim cLocation As Long
    Dim r As Long
    Dim codeMatches As Boolean
    Dim nameMatches As Boolean
    Dim locationMatches As Boolean

    Set lo = InventoryTable()
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    cSystemKey = ProductionColumnIndex(lo, "System_Key")
    cItemCode = ProductionColumnIndex(lo, "ITEM_CODE")
    If cItemCode = 0 Then cItemCode = ProductionColumnIndex(lo, "SKU")
    cItem = ProductionColumnIndex(lo, "ITEM")
    If cItem = 0 Then cItem = ProductionColumnIndex(lo, "ItemName")
    cLocation = ProductionColumnIndex(lo, "LOCATION")
    If cSystemKey = 0 Then Exit Function

    For r = 1 To lo.ListRows.Count
        codeMatches = (Trim$(itemCode) <> "" And cItemCode > 0 And _
                       StrComp(Trim$(NzStr(lo.DataBodyRange.Cells(r, cItemCode).Value)), _
                               Trim$(itemCode), vbTextCompare) = 0)
        nameMatches = (Trim$(itemName) <> "" And cItem > 0 And _
                       StrComp(Trim$(NzStr(lo.DataBodyRange.Cells(r, cItem).Value)), _
                               Trim$(itemName), vbTextCompare) = 0)
        locationMatches = (Trim$(locationValue) = "" Or cLocation = 0 Or _
                           StrComp(Trim$(NzStr(lo.DataBodyRange.Cells(r, cLocation).Value)), _
                                   Trim$(locationValue), vbTextCompare) = 0)
        If (codeMatches Or nameMatches) And locationMatches Then
            ResolveRunSystemKey = Trim$(NzStr(lo.DataBodyRange.Cells(r, cSystemKey).Value))
            If ResolveRunSystemKey <> "" Then Exit Function
        End If
    Next r
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
    If IsRunTreeOutputRow(lst, idx) Then Exit Sub
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
    If IsRunTreeOutputRow(lst, idx) Then
        ShowStatus "Select an input inventory choice row, not an output row."
        Exit Sub
    End If

    tableName = NzStr(lst.List(idx, 0))
    rowIndex = CLng(Val(NzStr(lst.List(idx, 1))))
    runLoc = ActiveRunLocation()
    invLoc = Trim$(NzStr(lst.List(idx, 9)))

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
    If (hasQty And qtyVal > 0) Or (hasSplit And splitVal > 0) Then
        If Not RunChoiceLocationAllowed(runLoc, invLoc, IIf(hasQty, qtyVal, splitVal)) Then
            ClearRunAllocationForListRow lst, idx
            BuildRunTreeFromPaletteList
            If runLoc = "" Then
                ShowStatus "Choose a production run location before allocating inventory."
            ElseIf invLoc = "" Then
                ShowStatus "This operator inventory row has no refreshed location. Click Refresh in Production Run, then select inventory at " & runLoc & "."
            Else
                ShowStatus "Allocation rejected. Inventory is at " & invLoc & "; production run location is " & runLoc & ". Use inventory at the production location. This row was cleared from the run."
            End If
            Exit Sub
        End If
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
    ShowStatus "Acceptable inventory allocation updated. Ingredient is " & FormatRunNumber(newTotalPct) & "% filled."
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
    If StrComp(headerName, "RECIPE_ID", vbTextCompare) = 0 Then
        lo.DataBodyRange.Cells(1, colIndex).NumberFormat = "@"
        lo.DataBodyRange.Cells(1, colIndex).Value2 = CStr(value)
    Else
        lo.DataBodyRange.Cells(1, colIndex).Value = value
    End If
End Sub

Private Sub SetRowValue(ByVal lr As ListRow, ByVal lo As ListObject, ByVal headerName As String, ByVal value As Variant)
    Dim colIndex As Long
    If lr Is Nothing Or lo Is Nothing Then Exit Sub
    colIndex = ProductionColumnIndex(lo, headerName)
    If colIndex <= 0 Then Exit Sub
    If StrComp(headerName, "RECIPE_ID", vbTextCompare) = 0 Then
        lr.Range.Cells(1, colIndex).NumberFormat = "@"
        lr.Range.Cells(1, colIndex).Value2 = CStr(value)
    Else
        lr.Range.Cells(1, colIndex).Value = value
    End If
End Sub

Private Sub ClearTableContentsKeepBlank(ByVal lo As ListObject)
    If lo Is Nothing Then Exit Sub
    EnsureTableRow lo
    If lo.DataBodyRange Is Nothing Then Exit Sub
    lo.DataBodyRange.ClearContents
End Sub

Private Function ProductionColumnIndex(ByVal lo As ListObject, ByVal headerName As String) As Long
    On Error Resume Next
    ProductionColumnIndex = CLng(Application.Run("mProduction.ColumnIndex", lo, headerName))
    On Error GoTo 0
End Function

Private Function RunProduction0(ByVal procName As String) As Variant
    ActivateOperatorWorkbookForRun
    RunProduction0 = Application.Run("mProduction." & procName)
End Function

Private Function RunProduction1(ByVal procName As String, ByVal arg1 As Variant) As Variant
    ActivateOperatorWorkbookForRun
    RunProduction1 = Application.Run("mProduction." & procName, arg1)
End Function

Private Function RunProduction2(ByVal procName As String, ByVal arg1 As Variant, ByVal arg2 As Variant) As Variant
    ActivateOperatorWorkbookForRun
    RunProduction2 = Application.Run("mProduction." & procName, arg1, arg2)
End Function

Private Sub RunProductionSub0(ByVal procName As String)
    ActivateOperatorWorkbookForRun
    Application.Run "mProduction." & procName
End Sub

Private Sub RunProductionSub1(ByVal procName As String, ByVal arg1 As Variant)
    ActivateOperatorWorkbookForRun
    Application.Run "mProduction." & procName, arg1
End Sub

Private Sub RunProductionSub2(ByVal procName As String, ByVal arg1 As Variant, ByVal arg2 As Variant)
    ActivateOperatorWorkbookForRun
    Application.Run "mProduction." & procName, arg1, arg2
End Sub

Private Function RunProductionObject0(ByVal procName As String) As Object
    ActivateOperatorWorkbookForRun
    Set RunProductionObject0 = Application.Run("mProduction." & procName)
End Function

Private Function RunProductionObject1(ByVal procName As String, ByVal arg1 As Variant) As Object
    ActivateOperatorWorkbookForRun
    Set RunProductionObject1 = Application.Run("mProduction." & procName, arg1)
End Function

Private Function RunProductionObject2(ByVal procName As String, ByVal arg1 As Variant, ByVal arg2 As Variant) As Object
    ActivateOperatorWorkbookForRun
    Set RunProductionObject2 = Application.Run("mProduction." & procName, arg1, arg2)
End Function

Private Sub ActivateOperatorWorkbookForRun()
    Dim wb As Workbook

    Set wb = ResolveOperatorWorkbook()
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    wb.Activate
    On Error GoTo 0
End Sub

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

Private Sub mCmbLineIo_Change()
    ConfigureRecipeLineInputs
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
    WriteSelectedRecipeBuilderLineFromForm False
    RunProductionSub0 "BtnSaveRecipe"
    RefreshRecipeLists
    RefreshBuilderLines
    RefreshAssignmentState
    RefreshLoaderState
    RefreshManagerState
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

Private Sub mBtnBuilderRelease_Click()
    Dim recipeId As String
    Dim recipeName As String
    Dim releaseReport As String

    recipeId = Trim$(mTxtRecipeId.Text)
    recipeName = Trim$(mTxtRecipeName.Text)
    If recipeId = "" Then
        ShowStatus "Select or load a saved recipe before releasing it."
        Exit Sub
    End If
    If MsgBox("Release " & IIf(recipeName = "", recipeId, recipeName) & _
              " for production runs?" & vbCrLf & vbCrLf & _
              "This publishes the latest immutable recipe version.", _
              vbQuestion Or vbYesNo Or vbDefaultButton2, _
              "Production Designs") <> vbYes Then Exit Sub

    releaseReport = NzStr(RunProduction1("ReleaseRecipeForProduction", recipeId))
    RefreshRecipeLists
    ShowStatus releaseReport
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

Private Sub mBtnLineMoveUp_Click()
    MoveSelectedRecipeBuilderLine -1
End Sub

Private Sub mBtnLineMoveDown_Click()
    MoveSelectedRecipeBuilderLine 1
End Sub

Private Sub mBtnLineUomAdd_Click()
    Dim uomName As String
    Dim report As String

    If Not modRoleUiAccess.CanCurrentUserPerformCapabilityCached("PROD_POST", report) Then
        ShowStatus report
        Exit Sub
    End If
    uomName = InputBox("Enter the UOM to add to this warehouse catalog:", _
                       "Recipe Builder - Add UOM", CStr(mTxtLineUom.Value))
    uomName = modUomSettings.NormalizeConfiguredUomName(uomName)
    If uomName = "" Then Exit Sub

    If modUomSettings.AddConfiguredUom(uomName, report) Then
        RefreshRecipeUomCatalog uomName
        ShowStatus report
    Else
        ShowStatus report
    End If
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
    Dim clearedCount As Long

    If mLoading Then Exit Sub
    SyncRunLocationCombo mCmbRunLocation, mCmbTreeRunLocation
    ClearMismatchedRunLocationAllocations clearedCount
    If clearedCount > 0 Then ShowStatus "Cleared " & CStr(clearedCount) & " allocation(s) that were not at the selected production run location."
End Sub

Private Sub mCmbTreeRunLocation_Change()
    Dim clearedCount As Long

    If mLoading Then Exit Sub
    SyncRunLocationCombo mCmbTreeRunLocation, mCmbRunLocation
    ClearMismatchedRunLocationAllocations clearedCount
    If clearedCount > 0 Then ShowStatus "Cleared " & CStr(clearedCount) & " allocation(s) that were not at the selected production run location."
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
    ' Real Output is staged when Complete Run is clicked.
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
    Dim refreshReport As String
    Dim refreshed As Boolean

    refreshed = RefreshProductionInventoryReadModel(refreshReport)
    ResetInventoryCache
    RefreshRecipeLists
    RefreshLoaderState
    RefreshManagerState
    If refreshed Then
        ShowStatus "Production Run inventory refreshed. " & refreshReport
    Else
        ShowStatus refreshReport
        MsgBox refreshReport, vbExclamation, "Production Inventory Refresh"
    End If
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
    If mLstRunPalette.ListCount = 0 Then
        ResetInventoryCache
        RefreshRunPaletteState
        If mLstRunPalette.ListCount = 0 Then
            ShowStatus "No acceptable inventory assignments were found for this recipe. Use Ingredients Assignment to select each USED ingredient, add acceptable inventory, and Save Assignment."
            Exit Sub
        End If
    End If
    If mLstRunPalette.ListIndex < 0 And mLstRunPalette.ListCount = 1 Then
        mLstRunPalette.ListIndex = 0
        LoadSelectedRunPaletteRow
    End If
    ApplySelectedRunPaletteSplit
End Sub

Private Sub mBtnRunTreeApplyPalette_Click()
    ApplySelectedRunPaletteSplit
End Sub

Private Sub mBtnManagerRefresh_Click()
    Dim refreshReport As String
    Dim refreshed As Boolean

    refreshed = RefreshProductionInventoryReadModel(refreshReport)
    ResetInventoryCache
    RefreshLoaderState
    RefreshManagerState
    If refreshed Then
        ShowStatus "Production Run inventory refreshed. " & refreshReport
    Else
        ShowStatus refreshReport
        MsgBox refreshReport, vbExclamation, "Production Inventory Refresh"
    End If
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
    RunProductionSub0 "ProductionToTotalInv"
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
    ResetInventoryCache
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
