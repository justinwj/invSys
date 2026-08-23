Attribute VB_Name = "TestDesignsDomain"
Option Explicit

Public Function TestDesignsSchema_CreatesAndValidatesAuthoritativeTables() As Long
    Dim wb As Workbook
    Dim report As String

    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail
    If Not modDesignsSchema.EnsureDesignsSchema(wb, report) Then GoTo CleanExit
    If modDesignsSchema.ValidateDesignsSchema(wb) <> "" Then GoTo CleanExit
    If FindDesignsTestTable(wb, "tblDesigns") Is Nothing Then GoTo CleanExit
    If FindDesignsTestTable(wb, "tblDesignLines") Is Nothing Then GoTo CleanExit
    If FindDesignsTestTable(wb, "tblDesignEvents") Is Nothing Then GoTo CleanExit
    If FindDesignsTestTable(wb, "tblAppliedDesignEvents") Is Nothing Then GoTo CleanExit
    If FindDesignsTestTable(wb, "tblLocks") Is Nothing Then GoTo CleanExit
    TestDesignsSchema_CreatesAndValidatesAuthoritativeTables = 1

CleanExit:
    CloseDesignsTestWorkbook wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestDesignsSchema_IsIdempotent() As Long
    Dim wb As Workbook
    Dim report As String
    Dim beforeCount As Long
    Dim afterCount As Long

    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail
    If Not modDesignsSchema.EnsureDesignsSchema(wb, report) Then GoTo CleanExit
    beforeCount = CountDesignsTestTables(wb)
    If Not modDesignsSchema.EnsureDesignsSchema(wb, report) Then GoTo CleanExit
    afterCount = CountDesignsTestTables(wb)
    If beforeCount = 13 And afterCount = beforeCount Then TestDesignsSchema_IsIdempotent = 1

CleanExit:
    CloseDesignsTestWorkbook wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestReusableProductionSchema_CreatesProcessRecipeProjections() As Long
    Dim wb As Workbook
    Dim report As String
    Dim tableNames As Variant
    Dim tableName As Variant

    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail
    If Not modDesignsSchema.EnsureDesignsSchema(wb, report) Then GoTo CleanExit
    tableNames = Array("tblProcesses", "tblProcessRequirements", _
        "tblProcessIngredientAlternatives", "tblProcessOutputs", _
        "tblProcessInstructions", "tblRecipes", "tblRecipeProcesses", _
        "tblRecipeConnections")
    For Each tableName In tableNames
        If FindDesignsTestTable(wb, CStr(tableName)) Is Nothing Then GoTo CleanExit
    Next tableName
    TestReusableProductionSchema_CreatesProcessRecipeProjections = 1

CleanExit:
    CloseDesignsTestWorkbook wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestProcessSave_AppliesReusableMultiOutputDefinition() As Long
    Dim wb As Workbook
    Dim report As String
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim evt As Object
    Dim loProcesses As ListObject
    Dim loOutputs As ListObject
    Dim payloadJson As String

    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail
    If Not modDesignsSchema.EnsureDesignsSchema(wb, report) Then GoTo CleanExit
    payloadJson = _
        "[{""RecordType"":""PROCESS"",""ProcessName"":""Blend Base""}," & _
        "{""RecordType"":""REQUIREMENT"",""RequirementId"":""REQ-RAW"",""RequirementName"":""Raw"",""Qty"":10,""UOM"":""LB""}," & _
        "{""RecordType"":""ALTERNATIVE"",""RequirementId"":""REQ-RAW"",""ITEM_CODE"":""SKU-RAW-A""}," & _
        "{""RecordType"":""OUTPUT"",""OutputId"":""OUT-MAIN"",""OutputName"":""Main Blend"",""ITEM_CODE"":""SKU-BLEND"",""Qty"":8,""UOM"":""LB""}," & _
        "{""RecordType"":""OUTPUT"",""OutputId"":""OUT-CO"",""OutputName"":""Co Product"",""ITEM_CODE"":""SKU-CO"",""Qty"":2,""UOM"":""LB""}," & _
        "{""RecordType"":""INSTRUCTION"",""InstructionOrdinal"":1,""Instruction"":""Blend""}]"
    Set evt = BuildDesignsTestEvent("PROC-EVT-1", "PROCESS_SAVE", _
        "PROC-BLEND", "1", payloadJson)
    If Not modDesignsApply.ApplyDesignEvent(evt, wb, "RUN-PROC-1", _
            statusOut, errorCode, errorMessage) Then GoTo CleanExit
    Set loProcesses = FindDesignsTestTable(wb, "tblProcesses")
    Set loOutputs = FindDesignsTestTable(wb, "tblProcessOutputs")
    If loProcesses Is Nothing Or loOutputs Is Nothing Then GoTo CleanExit
    If loProcesses.ListRows.Count <> 1 Then GoTo CleanExit
    If loOutputs.ListRows.Count <> 2 Then GoTo CleanExit
    TestProcessSave_AppliesReusableMultiOutputDefinition = 1

CleanExit:
    CloseDesignsTestWorkbook wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestProcessSave_RejectsDefinitionWithoutOutput() As Long
    Dim wb As Workbook
    Dim report As String
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim evt As Object
    Dim payloadJson As String

    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail
    If Not modDesignsSchema.EnsureDesignsSchema(wb, report) Then GoTo CleanExit
    payloadJson = _
        "[{""RecordType"":""PROCESS"",""ProcessName"":""Invalid""}," & _
        "{""RecordType"":""REQUIREMENT"",""RequirementId"":""REQ-1"",""RequirementName"":""Raw"",""Qty"":1,""UOM"":""EA""}]"
    Set evt = BuildDesignsTestEvent("PROC-EVT-NO-OUTPUT", "PROCESS_SAVE", _
        "PROC-INVALID", "1", payloadJson)
    If modDesignsApply.ApplyDesignEvent(evt, wb, "RUN-PROC-NO-OUTPUT", _
            statusOut, errorCode, errorMessage) Then GoTo CleanExit
    If StrComp(errorCode, "PROCESS_OUTPUT_REQUIRED", vbTextCompare) <> 0 Then GoTo CleanExit
    TestProcessSave_RejectsDefinitionWithoutOutput = 1

CleanExit:
    CloseDesignsTestWorkbook wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestRecipeSave_RejectsCircularProcessGraph() As Long
    Dim wb As Workbook
    Dim report As String
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim evt As Object
    Dim payloadJson As String

    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail
    If Not modDesignsSchema.EnsureDesignsSchema(wb, report) Then GoTo CleanExit
    payloadJson = _
        "[{""RecordType"":""RECIPE"",""RecipeName"":""Circular""}," & _
        "{""RecordType"":""PROCESS_NODE"",""ProcessNodeId"":""A"",""ProcessId"":""PROC-A"",""ProcessVersion"":""1"",""ExecutionOrdinal"":1}," & _
        "{""RecordType"":""PROCESS_NODE"",""ProcessNodeId"":""B"",""ProcessId"":""PROC-B"",""ProcessVersion"":""1"",""ExecutionOrdinal"":2}," & _
        "{""RecordType"":""CONNECTION"",""FromProcessNodeId"":""A"",""FromOutputId"":""OUT-A"",""ToProcessNodeId"":""B"",""ToRequirementId"":""REQ-B"",""Qty"":1,""UOM"":""EA""}," & _
        "{""RecordType"":""CONNECTION"",""FromProcessNodeId"":""B"",""FromOutputId"":""OUT-B"",""ToProcessNodeId"":""A"",""ToRequirementId"":""REQ-A"",""Qty"":1,""UOM"":""EA""}]"
    Set evt = BuildDesignsTestEvent("RECIPE-EVT-CYCLE", "RECIPE_SAVE", _
        "RECIPE-CYCLE", "1", payloadJson)
    If modDesignsApply.ApplyDesignEvent(evt, wb, "RUN-RECIPE-CYCLE", _
            statusOut, errorCode, errorMessage) Then GoTo CleanExit
    If StrComp(errorCode, "RECIPE_CYCLE", vbTextCompare) <> 0 Then GoTo CleanExit
    TestRecipeSave_RejectsCircularProcessGraph = 1

CleanExit:
    CloseDesignsTestWorkbook wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestProcessLifecycle_ReleasesObsoletesAndReusesVersions() As Long
    Dim wb As Workbook
    Dim report As String
    Dim payloadJson As String
    Dim errorCode As String

    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail
    If Not modDesignsSchema.EnsureDesignsSchema(wb, report) Then GoTo CleanExit
    payloadJson = ReusableProcessPayloadForTest( _
        "Reusable Blend", "REQ-RAW", "SKU-RAW", 4, "LB", _
        "OUT-BLEND", "SKU-BLEND", 4, "LB", True)
    If Not SaveReleaseProcessForTest(wb, "PROC-REUSE-V1", _
            "PROC-REUSE", "1", payloadJson, errorCode) Then GoTo CleanExit
    If Not ApplyReusableEventForTest(wb, "PROC-REUSE-V2-SAVE", _
            "PROCESS_SAVE", "PROC-REUSE", "2", payloadJson, errorCode) Then GoTo CleanExit
    If Not ApplyReusableEventForTest(wb, "PROC-REUSE-V1-OBSOLETE", _
            "PROCESS_OBSOLETE", "PROC-REUSE", "1", "", errorCode) Then GoTo CleanExit
    If StrComp(ReusableStatusForTest(wb, "tblProcesses", "ProcessId", _
            "ProcessVersion", "PROC-REUSE", "1"), "OBSOLETE", vbTextCompare) <> 0 Then GoTo CleanExit
    If StrComp(ReusableStatusForTest(wb, "tblProcesses", "ProcessId", _
            "ProcessVersion", "PROC-REUSE", "2"), "DRAFT", vbTextCompare) <> 0 Then GoTo CleanExit
    If CountReusableRowsForTest(wb, "tblProcessIngredientAlternatives", _
            "ProcessId", "ProcessVersion", "PROC-REUSE", "2") <> 1 Then GoTo CleanExit
    TestProcessLifecycle_ReleasesObsoletesAndReusesVersions = 1

CleanExit:
    CloseDesignsTestWorkbook wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestRecipeRelease_RejectsMissingOrUnreleasedProcessVersion() As Long
    Dim wb As Workbook
    Dim report As String
    Dim payloadJson As String
    Dim recipeJson As String
    Dim errorCode As String

    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail
    If Not modDesignsSchema.EnsureDesignsSchema(wb, report) Then GoTo CleanExit
    payloadJson = ReusableProcessPayloadForTest( _
        "Draft Only", "", "", 0, "", "OUT-DRAFT", "SKU-DRAFT", 1, "EA", False)
    If Not ApplyReusableEventForTest(wb, "PROC-DRAFT-SAVE", "PROCESS_SAVE", _
            "PROC-DRAFT", "1", payloadJson, errorCode) Then GoTo CleanExit
    recipeJson = _
        "[{""RecordType"":""RECIPE"",""RecipeName"":""Draft Reference""}," & _
        "{""RecordType"":""PROCESS_NODE"",""ProcessNodeId"":""A"",""ProcessId"":""PROC-DRAFT"",""ProcessVersion"":""1"",""ExecutionOrdinal"":1}]"
    If Not ApplyReusableEventForTest(wb, "RECIPE-DRAFT-SAVE", "RECIPE_SAVE", _
            "RECIPE-DRAFT", "1", recipeJson, errorCode) Then GoTo CleanExit
    If ApplyReusableEventForTest(wb, "RECIPE-DRAFT-RELEASE", "RECIPE_RELEASE", _
            "RECIPE-DRAFT", "1", "", errorCode) Then GoTo CleanExit
    If StrComp(errorCode, "RECIPE_PROCESS_NOT_RELEASED", vbTextCompare) <> 0 Then GoTo CleanExit
    TestRecipeRelease_RejectsMissingOrUnreleasedProcessVersion = 1

CleanExit:
    CloseDesignsTestWorkbook wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestRecipeRelease_RejectsUnresolvedExternalRequirement() As Long
    Dim wb As Workbook
    Dim report As String
    Dim processJson As String
    Dim recipeJson As String
    Dim errorCode As String

    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail
    If Not modDesignsSchema.EnsureDesignsSchema(wb, report) Then GoTo CleanExit
    processJson = ReusableProcessPayloadForTest( _
        "Needs Assignment", "REQ-OPEN", "", 1, "EA", _
        "OUT-OPEN", "SKU-OPEN", 1, "EA", False)
    If Not SaveReleaseProcessForTest(wb, "PROC-OPEN", _
            "PROC-OPEN", "1", processJson, errorCode) Then GoTo CleanExit
    recipeJson = _
        "[{""RecordType"":""RECIPE"",""RecipeName"":""Unresolved""}," & _
        "{""RecordType"":""PROCESS_NODE"",""ProcessNodeId"":""A"",""ProcessId"":""PROC-OPEN"",""ProcessVersion"":""1"",""ExecutionOrdinal"":1}]"
    If Not ApplyReusableEventForTest(wb, "RECIPE-OPEN-SAVE", "RECIPE_SAVE", _
            "RECIPE-OPEN", "1", recipeJson, errorCode) Then GoTo CleanExit
    If ApplyReusableEventForTest(wb, "RECIPE-OPEN-RELEASE", "RECIPE_RELEASE", _
            "RECIPE-OPEN", "1", "", errorCode) Then GoTo CleanExit
    If StrComp(errorCode, "RECIPE_UNRESOLVED_REQUIREMENT", vbTextCompare) <> 0 Then GoTo CleanExit
    TestRecipeRelease_RejectsUnresolvedExternalRequirement = 1

CleanExit:
    CloseDesignsTestWorkbook wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestRecipeRelease_RejectsIncompatibleConnection() As Long
    TestRecipeRelease_RejectsIncompatibleConnection = _
        RunInvalidRecipeConnectionTest("INCOMPATIBLE", 5, "LB", 5, "EA", _
            5, "LB", 1, 2, "RECIPE_CONNECTION_INCOMPATIBLE")
End Function

Public Function TestRecipeRelease_RejectsOutputOverallocation() As Long
    TestRecipeRelease_RejectsOutputOverallocation = _
        RunInvalidRecipeConnectionTest("OVERALLOCATED", 5, "LB", 6, "LB", _
            6, "LB", 1, 2, "RECIPE_OUTPUT_OVERALLOCATED")
End Function

Public Function TestRecipeRelease_RejectsContradictoryExecutionOrder() As Long
    TestRecipeRelease_RejectsContradictoryExecutionOrder = _
        RunInvalidRecipeConnectionTest("BAD-ORDER", 5, "LB", 5, "LB", _
            5, "LB", 2, 1, "RECIPE_EXECUTION_ORDER")
End Function

Public Function TestProcessObsolete_RejectsReleasedRecipeDependency() As Long
    Dim wb As Workbook
    Dim report As String
    Dim sourceJson As String
    Dim sinkJson As String
    Dim recipeJson As String
    Dim errorCode As String

    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail
    If Not modDesignsSchema.EnsureDesignsSchema(wb, report) Then GoTo CleanExit
    sourceJson = ReusableProcessPayloadForTest( _
        "Source", "", "", 0, "", "OUT-A", "SKU-A", 5, "LB", False)
    sinkJson = ReusableProcessPayloadForTest( _
        "Sink", "REQ-B", "SKU-A", 5, "LB", "OUT-B", "SKU-B", 5, "LB", True)
    If Not SaveReleaseProcessForTest(wb, "PROC-DEP-A", _
            "PROC-DEP-A", "1", sourceJson, errorCode) Then GoTo CleanExit
    If Not SaveReleaseProcessForTest(wb, "PROC-DEP-B", _
            "PROC-DEP-B", "1", sinkJson, errorCode) Then GoTo CleanExit
    recipeJson = TwoNodeRecipePayloadForTest( _
        "Dependency", "PROC-DEP-A", "PROC-DEP-B", 5, "LB", 1, 2)
    If Not ApplyReusableEventForTest(wb, "RECIPE-DEP-SAVE", "RECIPE_SAVE", _
            "RECIPE-DEP", "1", recipeJson, errorCode) Then GoTo CleanExit
    If Not ApplyReusableEventForTest(wb, "RECIPE-DEP-RELEASE", "RECIPE_RELEASE", _
            "RECIPE-DEP", "1", "", errorCode) Then GoTo CleanExit
    If ApplyReusableEventForTest(wb, "PROC-DEP-A-OBSOLETE", "PROCESS_OBSOLETE", _
            "PROC-DEP-A", "1", "", errorCode) Then GoTo CleanExit
    If StrComp(errorCode, "PROCESS_HAS_RELEASED_RECIPE_DEPENDENCY", vbTextCompare) <> 0 Then GoTo CleanExit
    TestProcessObsolete_RejectsReleasedRecipeDependency = 1

CleanExit:
    CloseDesignsTestWorkbook wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestRecipeLifecycle_ReleasesValidGraphAndThenObsoletes() As Long
    Dim wb As Workbook
    Dim report As String
    Dim sourceJson As String
    Dim sinkJson As String
    Dim recipeJson As String
    Dim errorCode As String

    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail
    If Not modDesignsSchema.EnsureDesignsSchema(wb, report) Then GoTo CleanExit
    sourceJson = ReusableProcessPayloadForTest( _
        "Source", "", "", 0, "", "OUT-A", "SKU-A", 5, "LB", False)
    sinkJson = ReusableProcessPayloadForTest( _
        "Sink", "REQ-B", "SKU-A", 5, "LB", "OUT-B", "SKU-B", 5, "LB", True)
    If Not SaveReleaseProcessForTest(wb, "PROC-VALID-A", _
            "PROC-VALID-A", "1", sourceJson, errorCode) Then GoTo CleanExit
    If Not SaveReleaseProcessForTest(wb, "PROC-VALID-B", _
            "PROC-VALID-B", "1", sinkJson, errorCode) Then GoTo CleanExit
    recipeJson = TwoNodeRecipePayloadForTest( _
        "Valid Graph", "PROC-VALID-A", "PROC-VALID-B", 5, "LB", 1, 2)
    If Not ApplyReusableEventForTest(wb, "RECIPE-VALID-SAVE", "RECIPE_SAVE", _
            "RECIPE-VALID", "1", recipeJson, errorCode) Then GoTo CleanExit
    If Not ApplyReusableEventForTest(wb, "RECIPE-VALID-RELEASE", "RECIPE_RELEASE", _
            "RECIPE-VALID", "1", "", errorCode) Then GoTo CleanExit
    If Not ApplyReusableEventForTest(wb, "RECIPE-VALID-OBSOLETE", "RECIPE_OBSOLETE", _
            "RECIPE-VALID", "1", "", errorCode) Then GoTo CleanExit
    If StrComp(ReusableStatusForTest(wb, "tblRecipes", "RecipeId", _
            "RecipeVersion", "RECIPE-VALID", "1"), "OBSOLETE", vbTextCompare) <> 0 Then GoTo CleanExit
    TestRecipeLifecycle_ReleasesValidGraphAndThenObsoletes = 1

CleanExit:
    CloseDesignsTestWorkbook wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestDesignsQueries_ListDesignsAndGetBOMAreReadOnly() As Long
    Dim wb As Workbook
    Dim report As String
    Dim loDesigns As ListObject
    Dim loLines As ListObject
    Dim lr As ListRow
    Dim designs As Variant
    Dim bom As Variant
    Dim designsRowsBefore As Long
    Dim linesRowsBefore As Long

    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail
    If Not modDesignsSchema.EnsureDesignsSchema(wb, report) Then GoTo CleanExit
    Set loDesigns = FindDesignsTestTable(wb, "tblDesigns")
    Set loLines = FindDesignsTestTable(wb, "tblDesignLines")

    Set lr = loDesigns.ListRows.Add
    SetDesignsTestValue loDesigns, lr.Index, "DesignId", "TEA-BLACK"
    SetDesignsTestValue loDesigns, lr.Index, "DesignVersion", "1"
    SetDesignsTestValue loDesigns, lr.Index, "DesignType", "RECIPE"
    SetDesignsTestValue loDesigns, lr.Index, "DesignName", "Brewed Black Tea"
    SetDesignsTestValue loDesigns, lr.Index, "Status", "RELEASED"

    Set lr = loLines.ListRows.Add
    SetDesignsTestValue loLines, lr.Index, "DesignId", "TEA-BLACK"
    SetDesignsTestValue loLines, lr.Index, "DesignVersion", "1"
    SetDesignsTestValue loLines, lr.Index, "LineNo", 1
    SetDesignsTestValue loLines, lr.Index, "IOType", "USED"
    SetDesignsTestValue loLines, lr.Index, "ComponentSKU", "SKU-BLACK-TEA"
    SetDesignsTestValue loLines, lr.Index, "Qty", 32.5
    SetDesignsTestValue loLines, lr.Index, "UOM", "LB"

    designsRowsBefore = loDesigns.ListRows.Count
    linesRowsBefore = loLines.ListRows.Count
    designs = modDesignsQueries.ListDesigns(wb, "RELEASED")
    bom = modDesignsQueries.GetBOM("TEA-BLACK", "1", wb)
    If IsEmpty(designs) Or Not IsArray(designs) Then GoTo CleanExit
    If IsEmpty(bom) Or Not IsArray(bom) Then GoTo CleanExit
    If CStr(designs(1, 1)) <> "TEA-BLACK" Then GoTo CleanExit
    If CStr(bom(1, 4)) <> "SKU-BLACK-TEA" Then GoTo CleanExit
    If loDesigns.ListRows.Count <> designsRowsBefore Then GoTo CleanExit
    If loLines.ListRows.Count <> linesRowsBefore Then GoTo CleanExit
    TestDesignsQueries_ListDesignsAndGetBOMAreReadOnly = 1

CleanExit:
    CloseDesignsTestWorkbook wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestDesignsQueries_StatusConstrainedBOMRejectsDraftAndObsolete() As Long
    Dim wb As Workbook
    Dim report As String
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim evt As Object
    Dim bom As Variant

    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail
    If Not modDesignsSchema.EnsureDesignsSchema(wb, report) Then GoTo CleanExit

    Set evt = BuildDesignsTestEvent("DES-EVT-STATUS-1", "DESIGN_CREATE", "TEA-STATUS", "1", _
        "[{""DesignType"":""RECIPE"",""DesignName"":""Status Tea"",""LineNo"":1," & _
        """IOType"":""OUTPUT"",""ComponentSKU"":""SKU-STATUS-TEA"",""Qty"":1,""UOM"":""LB""}]")
    If Not modDesignsApply.ApplyDesignEvent(evt, wb, "RUN-STATUS-1", statusOut, errorCode, errorMessage) Then GoTo CleanExit

    bom = modDesignsQueries.GetBOMForStatus("TEA-STATUS", "1", "RELEASED", wb)
    If IsUsableDesignsTestArray(bom) Then GoTo CleanExit

    Set evt = BuildDesignsTestEvent("DES-EVT-STATUS-2", "DESIGN_RELEASE", "TEA-STATUS", "1", "")
    statusOut = "": errorCode = "": errorMessage = ""
    If Not modDesignsApply.ApplyDesignEvent(evt, wb, "RUN-STATUS-2", statusOut, errorCode, errorMessage) Then GoTo CleanExit
    bom = modDesignsQueries.GetBOMForStatus("TEA-STATUS", "1", "RELEASED", wb)
    If Not IsUsableDesignsTestArray(bom) Then GoTo CleanExit

    Set evt = BuildDesignsTestEvent("DES-EVT-STATUS-3", "DESIGN_OBSOLETE", "TEA-STATUS", "1", "")
    statusOut = "": errorCode = "": errorMessage = ""
    If Not modDesignsApply.ApplyDesignEvent(evt, wb, "RUN-STATUS-3", statusOut, errorCode, errorMessage) Then GoTo CleanExit
    bom = modDesignsQueries.GetBOMForStatus("TEA-STATUS", "1", "RELEASED", wb)
    If IsUsableDesignsTestArray(bom) Then GoTo CleanExit

    TestDesignsQueries_StatusConstrainedBOMRejectsDraftAndObsolete = 1

CleanExit:
    CloseDesignsTestWorkbook wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestDesignsDomain_DiagnosticDeclaresNoStartupMutation() As Long
    Dim diagnostic As String
    diagnostic = modDesignsInit.DiagnoseDesignsDomain()
    If InStr(1, diagnostic, "StartupMutation=False", vbTextCompare) > 0 _
       And InStr(1, diagnostic, "WHx.invSys.Data.Designs.xlsb", vbTextCompare) > 0 Then
        TestDesignsDomain_DiagnosticDeclaresNoStartupMutation = 1
    End If
End Function

Public Function TestDesignsRuntime_CanonicalAuthorityWindowStaysHidden() As Long
    Dim operatorWb As Workbook
    Dim designsWb As Workbook
    Dim win As Window
    Dim rootPath As String
    Dim designsPath As String
    Dim report As String
    Dim allHidden As Boolean

    rootPath = Environ$("TEMP") & "\invSys-designs-hidden-" & _
               Format$(Now, "yyyymmddhhnnss") & "-" & CStr(CLng(Timer * 100)) & "\"
    designsPath = rootPath & "WH-HIDDEN.invSys.Data.Designs.xlsb"

    On Error GoTo CleanFail
    MkDir Left$(rootPath, Len(rootPath) - 1)
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath
    Set operatorWb = Application.Workbooks.Add(xlWBATWorksheet)
    operatorWb.Activate

    Set designsWb = modDesignsRuntime.ResolveDesignsWorkbook("WH-HIDDEN", Nothing, report)
    If designsWb Is Nothing Then GoTo CleanExit
    If designsWb.Windows.Count = 0 Then GoTo CleanExit

    allHidden = True
    For Each win In designsWb.Windows
        If win.Visible Then allHidden = False
    Next win
    If Not allHidden Then GoTo CleanExit
    If Application.ActiveWorkbook Is Nothing Then GoTo CleanExit
    If Not (Application.ActiveWorkbook Is operatorWb) Then GoTo CleanExit

    TestDesignsRuntime_CanonicalAuthorityWindowStaysHidden = 1

CleanExit:
    On Error Resume Next
    If Not designsWb Is Nothing Then designsWb.Close SaveChanges:=False
    If Not operatorWb Is Nothing Then operatorWb.Close SaveChanges:=False
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    If Len(Dir$(designsPath, vbNormal)) > 0 Then Kill designsPath
    RmDir Left$(rootPath, Len(rootPath) - 1)
    On Error GoTo 0
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestDesignInboxSchema_CarriesDesignIdentity() As Long
    Dim wb As Workbook
    Dim lo As ListObject
    Dim report As String

    On Error GoTo CleanExit
    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modProcessor.EnsureProductionInboxSchema(wb, report) Then GoTo CleanExit
    Set lo = FindDesignsTestTable(wb, "tblInboxProd")
    If lo Is Nothing Then GoTo CleanExit
    If Not DesignsTestColumnExists(lo, "DesignId") Then GoTo CleanExit
    If Not DesignsTestColumnExists(lo, "DesignVersion") Then GoTo CleanExit
    TestDesignInboxSchema_CarriesDesignIdentity = 1

CleanExit:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
End Function

Public Function TestDesignsApply_LifecycleIsIdempotentAndRebuildable() As Long
    Dim wb As Workbook
    Dim report As String
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim createEvent As Object
    Dim releaseEvent As Object
    Dim loDesigns As ListObject
    Dim loLines As ListObject
    Dim loEvents As ListObject

    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail
    If Not modDesignsSchema.EnsureDesignsSchema(wb, report) Then GoTo CleanExit
    Set createEvent = BuildDesignsTestEvent("DES-EVT-1", "DESIGN_CREATE", "TEA-BLACK", "1", _
        "[{""DesignType"":""RECIPE"",""DesignName"":""Brewed Black Tea"",""Description"":""Concentrate""," & _
        """LineNo"":1,""Process"":1,""IOType"":""USED"",""ComponentSKU"":""SKU-BLACK-TEA""," & _
        """Qty"":32.5,""UOM"":""LB"",""Percent"":100}]")
    If Not modDesignsApply.ApplyDesignEvent(createEvent, wb, "RUN-DES-1", statusOut, errorCode, errorMessage) Then GoTo CleanExit
    If statusOut <> "APPLIED" Then GoTo CleanExit

    statusOut = "": errorCode = "": errorMessage = ""
    If Not modDesignsApply.ApplyDesignEvent(createEvent, wb, "RUN-DES-2", statusOut, errorCode, errorMessage) Then GoTo CleanExit
    If statusOut <> "SKIP_DUP" Then GoTo CleanExit

    Set releaseEvent = BuildDesignsTestEvent("DES-EVT-2", "DESIGN_RELEASE", "TEA-BLACK", "1", "")
    statusOut = "": errorCode = "": errorMessage = ""
    If Not modDesignsApply.ApplyDesignEvent(releaseEvent, wb, "RUN-DES-3", statusOut, errorCode, errorMessage) Then GoTo CleanExit

    Set loDesigns = FindDesignsTestTable(wb, "tblDesigns")
    Set loLines = FindDesignsTestTable(wb, "tblDesignLines")
    Set loEvents = FindDesignsTestTable(wb, "tblDesignEvents")
    If loEvents.ListRows.Count <> 2 Then GoTo CleanExit
    If loDesigns.ListRows.Count <> 1 Or loLines.ListRows.Count <> 1 Then GoTo CleanExit
    If CStr(loDesigns.DataBodyRange.Cells(1, loDesigns.ListColumns("Status").Index).Value) <> "RELEASED" Then GoTo CleanExit

    Do While loDesigns.ListRows.Count > 0: loDesigns.ListRows(1).Delete: Loop
    Do While loLines.ListRows.Count > 0: loLines.ListRows(1).Delete: Loop
    If Not modDesignsApply.RebuildDesignProjections(wb, report) Then GoTo CleanExit
    If loDesigns.ListRows.Count <> 1 Or loLines.ListRows.Count <> 1 Then GoTo CleanExit
    If CStr(loDesigns.DataBodyRange.Cells(1, loDesigns.ListColumns("Status").Index).Value) <> "RELEASED" Then GoTo CleanExit
    TestDesignsApply_LifecycleIsIdempotentAndRebuildable = 1

CleanExit:
    CloseDesignsTestWorkbook wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestDesignsApply_ObsoleteLifecycleIsRebuildable() As Long
    Dim wb As Workbook
    Dim report As String
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim evt As Object
    Dim loDesigns As ListObject
    Dim loLines As ListObject

    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail
    If Not modDesignsSchema.EnsureDesignsSchema(wb, report) Then GoTo CleanExit

    Set evt = BuildDesignsTestEvent("DES-EVT-40", "DESIGN_CREATE", "TEA-OBSOLETE", "3", _
        "[{""DesignType"":""RECIPE"",""DesignName"":""Retired Tea"",""LineNo"":1," & _
        """IOType"":""OUTPUT"",""ComponentSKU"":""SKU-RETIRED-TEA"",""Qty"":1,""UOM"":""LB""}]")
    If Not modDesignsApply.ApplyDesignEvent(evt, wb, "RUN-DES-40", statusOut, errorCode, errorMessage) Then GoTo CleanExit

    Set evt = BuildDesignsTestEvent("DES-EVT-41", "DESIGN_RELEASE", "TEA-OBSOLETE", "3", "")
    statusOut = "": errorCode = "": errorMessage = ""
    If Not modDesignsApply.ApplyDesignEvent(evt, wb, "RUN-DES-41", statusOut, errorCode, errorMessage) Then GoTo CleanExit

    Set evt = BuildDesignsTestEvent("DES-EVT-42", "DESIGN_OBSOLETE", "TEA-OBSOLETE", "3", "")
    statusOut = "": errorCode = "": errorMessage = ""
    If Not modDesignsApply.ApplyDesignEvent(evt, wb, "RUN-DES-42", statusOut, errorCode, errorMessage) Then GoTo CleanExit

    Set loDesigns = FindDesignsTestTable(wb, "tblDesigns")
    Set loLines = FindDesignsTestTable(wb, "tblDesignLines")
    If loDesigns.ListRows.Count <> 1 Or loLines.ListRows.Count <> 1 Then GoTo CleanExit
    If CStr(loDesigns.DataBodyRange.Cells(1, loDesigns.ListColumns("Status").Index).Value) <> "OBSOLETE" Then GoTo CleanExit

    Do While loDesigns.ListRows.Count > 0: loDesigns.ListRows(1).Delete: Loop
    Do While loLines.ListRows.Count > 0: loLines.ListRows(1).Delete: Loop
    If Not modDesignsApply.RebuildDesignProjections(wb, report) Then GoTo CleanExit
    If loDesigns.ListRows.Count <> 1 Or loLines.ListRows.Count <> 1 Then GoTo CleanExit
    If CStr(loDesigns.DataBodyRange.Cells(1, loDesigns.ListColumns("Status").Index).Value) <> "OBSOLETE" Then GoTo CleanExit
    TestDesignsApply_ObsoleteLifecycleIsRebuildable = 1

CleanExit:
    CloseDesignsTestWorkbook wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestDesignsApply_RejectsDuplicateImmutableVersion() As Long
    Dim wb As Workbook
    Dim report As String
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim evt As Object

    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail
    If Not modDesignsSchema.EnsureDesignsSchema(wb, report) Then GoTo CleanExit
    Set evt = BuildDesignsTestEvent("DES-EVT-10", "DESIGN_CREATE", "TEA-GREEN", "1", _
        "[{""DesignType"":""RECIPE"",""DesignName"":""Green Tea"",""LineNo"":1," & _
        """IOType"":""OUTPUT"",""ComponentSKU"":""SKU-GREEN-TEA"",""Qty"":1,""UOM"":""LB""}]")
    If Not modDesignsApply.ApplyDesignEvent(evt, wb, "RUN-DES-10", statusOut, errorCode, errorMessage) Then GoTo CleanExit
    Set evt = BuildDesignsTestEvent("DES-EVT-11", "DESIGN_CREATE", "TEA-GREEN", "1", _
        "[{""DesignType"":""RECIPE"",""DesignName"":""Changed Green Tea""}]")
    statusOut = "": errorCode = "": errorMessage = ""
    If modDesignsApply.ApplyDesignEvent(evt, wb, "RUN-DES-11", statusOut, errorCode, errorMessage) Then GoTo CleanExit
    If errorCode <> "DESIGN_VERSION_EXISTS" Then GoTo CleanExit
    TestDesignsApply_RejectsDuplicateImmutableVersion = 1

CleanExit:
    CloseDesignsTestWorkbook wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestDesignsApply_RebuildUsesAppliedSeqNotTableOrder() As Long
    Dim wb As Workbook
    Dim report As String
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim evt As Object
    Dim loEvents As ListObject
    Dim loDesigns As ListObject
    Dim rowOne As Variant

    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail
    If Not modDesignsSchema.EnsureDesignsSchema(wb, report) Then GoTo CleanExit
    Set evt = BuildDesignsTestEvent("DES-EVT-20", "DESIGN_CREATE", "TEA-ORDERED", "1", _
        "[{""DesignType"":""RECIPE"",""DesignName"":""Ordered Tea""}]")
    If Not modDesignsApply.ApplyDesignEvent(evt, wb, "RUN-DES-20", statusOut, errorCode, errorMessage) Then GoTo CleanExit
    Set evt = BuildDesignsTestEvent("DES-EVT-21", "DESIGN_RELEASE", "TEA-ORDERED", "1", "")
    statusOut = "": errorCode = "": errorMessage = ""
    If Not modDesignsApply.ApplyDesignEvent(evt, wb, "RUN-DES-21", statusOut, errorCode, errorMessage) Then GoTo CleanExit

    Set loEvents = FindDesignsTestTable(wb, "tblDesignEvents")
    Set loDesigns = FindDesignsTestTable(wb, "tblDesigns")
    If loEvents Is Nothing Or loDesigns Is Nothing Then GoTo CleanExit
    If loEvents.ListRows.Count <> 2 Then GoTo CleanExit
    rowOne = loEvents.DataBodyRange.Rows(1).Value
    loEvents.DataBodyRange.Rows(1).Value = loEvents.DataBodyRange.Rows(2).Value
    loEvents.DataBodyRange.Rows(2).Value = rowOne
    Do While loDesigns.ListRows.Count > 0: loDesigns.ListRows(1).Delete: Loop

    If Not modDesignsApply.RebuildDesignProjections(wb, report) Then GoTo CleanExit
    If loDesigns.ListRows.Count <> 1 Then GoTo CleanExit
    If CStr(loDesigns.DataBodyRange.Cells(1, loDesigns.ListColumns("Status").Index).Value) <> "RELEASED" Then GoTo CleanExit
    TestDesignsApply_RebuildUsesAppliedSeqNotTableOrder = 1

CleanExit:
    CloseDesignsTestWorkbook wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestDesignsApply_ImmutableVersionSurvivesProjectionLoss() As Long
    Dim wb As Workbook
    Dim report As String
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim evt As Object
    Dim loEvents As ListObject
    Dim loDesigns As ListObject
    Dim loLines As ListObject

    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail
    If Not modDesignsSchema.EnsureDesignsSchema(wb, report) Then GoTo CleanExit
    Set evt = BuildDesignsTestEvent("DES-EVT-30", "DESIGN_CREATE", "TEA-HISTORY", "1", _
        "[{""DesignType"":""RECIPE"",""DesignName"":""History Tea""}]")
    If Not modDesignsApply.ApplyDesignEvent(evt, wb, "RUN-DES-30", statusOut, errorCode, errorMessage) Then GoTo CleanExit

    Set loEvents = FindDesignsTestTable(wb, "tblDesignEvents")
    Set loDesigns = FindDesignsTestTable(wb, "tblDesigns")
    Set loLines = FindDesignsTestTable(wb, "tblDesignLines")
    Do While loDesigns.ListRows.Count > 0: loDesigns.ListRows(1).Delete: Loop
    Do While loLines.ListRows.Count > 0: loLines.ListRows(1).Delete: Loop

    Set evt = BuildDesignsTestEvent("DES-EVT-31", "DESIGN_CREATE", "TEA-HISTORY", "1", _
        "[{""DesignType"":""RECIPE"",""DesignName"":""Illegal Rewrite""}]")
    statusOut = "": errorCode = "": errorMessage = ""
    If modDesignsApply.ApplyDesignEvent(evt, wb, "RUN-DES-31", statusOut, errorCode, errorMessage) Then GoTo CleanExit
    If errorCode <> "DESIGN_VERSION_EXISTS" Then GoTo CleanExit
    If loEvents.ListRows.Count <> 1 Then GoTo CleanExit
    If loDesigns.ListRows.Count <> 1 Then GoTo CleanExit
    TestDesignsApply_ImmutableVersionSurvivesProjectionLoss = 1

CleanExit:
    CloseDesignsTestWorkbook wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestInventorySchema_RejectsMissingOrAddinAuthority() As Long
    Dim report As String

    If modInventorySchema.EnsureInventorySchema(Nothing, report) Then Exit Function
    If InStr(1, report, "Authoritative Inventory workbook was not supplied", vbTextCompare) = 0 Then Exit Function
    TestInventorySchema_RejectsMissingOrAddinAuthority = 1
End Function

Public Function TestInventoryDomain_DiagnosticDisablesLegacyDirectWrites() As Long
    Dim diagnostic As String

    diagnostic = modInventoryInit.DiagnoseInventoryDomain()
    If InStr(1, diagnostic, "LegacyDirectWrites=False", vbTextCompare) = 0 Then Exit Function
    If InStr(1, diagnostic, "UndoModel=CompensatingEvent", vbTextCompare) = 0 Then Exit Function
    TestInventoryDomain_DiagnosticDisablesLegacyDirectWrites = 1
End Function

Public Function TestInventoryDomain_LegacyLogDeletionBridgeIsNoOp() As Long
    Dim removedRows As Collection

    Set removedRows = modInventoryBridgeApi.RemoveLastBulkLogEntriesBridgeResult(5)
    If removedRows Is Nothing Then Exit Function
    If removedRows.Count <> 0 Then Exit Function
    TestInventoryDomain_LegacyLogDeletionBridgeIsNoOp = 1
End Function

Public Function TestInventoryQueries_ReadRebuiltEventLogProjection() As Long
    Dim wb As Workbook
    Dim report As String
    Dim loLog As ListObject
    Dim lr As ListRow
    Dim locations As Variant

    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail
    If Not modInventorySchema.EnsureInventorySchema(wb, report) Then GoTo CleanExit
    Set loLog = FindDesignsTestTable(wb, "tblInventoryLog")
    loLog.Parent.Unprotect

    Set lr = loLog.ListRows.Add
    SetDesignsTestValue loLog, lr.Index, "EventID", "INV-Q-1"
    SetDesignsTestValue loLog, lr.Index, "AppliedSeq", 1
    SetDesignsTestValue loLog, lr.Index, "AppliedAtUTC", Now
    SetDesignsTestValue loLog, lr.Index, "SKU", "SKU-QUERY-1"
    SetDesignsTestValue loLog, lr.Index, "QtyDelta", 10
    SetDesignsTestValue loLog, lr.Index, "Location", "CLEARVIEW"

    Set lr = loLog.ListRows.Add
    SetDesignsTestValue loLog, lr.Index, "EventID", "INV-Q-2"
    SetDesignsTestValue loLog, lr.Index, "AppliedSeq", 2
    SetDesignsTestValue loLog, lr.Index, "AppliedAtUTC", Now
    SetDesignsTestValue loLog, lr.Index, "SKU", "SKU-QUERY-1"
    SetDesignsTestValue loLog, lr.Index, "QtyDelta", -3
    SetDesignsTestValue loLog, lr.Index, "Location", "CLEARVIEW"

    If Not modInventoryApply.RebuildInventoryProjectionsForWorkbook(wb, report) Then GoTo CleanExit
    If Abs(modInventoryQueries.GetOnHandQty("SKU-QUERY-1", wb) - 7) > 0.0001 Then GoTo CleanExit
    locations = modInventoryQueries.GetLocationBalances("SKU-QUERY-1", wb)
    If IsEmpty(locations) Or Not IsArray(locations) Then GoTo CleanExit
    If CStr(locations(1, 1)) <> "CLEARVIEW" Then GoTo CleanExit
    If Abs(CDbl(locations(1, 2)) - 7) > 0.0001 Then GoTo CleanExit
    TestInventoryQueries_ReadRebuiltEventLogProjection = 1

CleanExit:
    CloseDesignsTestWorkbook wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestInventoryQueries_PickerPublishesEverySkuLocation() As Long
    Dim wb As Workbook
    Dim report As String
    Dim loCatalog As ListObject
    Dim loBalance As ListObject
    Dim loLocation As ListObject
    Dim lr As ListRow
    Dim items As Variant
    Dim r As Long
    Dim foundA1 As Boolean
    Dim foundClearview As Boolean

    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail
    If Not modInventorySchema.EnsureInventorySchema(wb, report) Then GoTo CleanExit
    Set loCatalog = FindDesignsTestTable(wb, "tblSkuCatalog")
    Set loBalance = FindDesignsTestTable(wb, "tblSkuBalance")
    Set loLocation = FindDesignsTestTable(wb, "tblLocationBalance")
    If loCatalog Is Nothing Or loBalance Is Nothing Or loLocation Is Nothing Then GoTo CleanExit
    loCatalog.Parent.Unprotect
    loBalance.Parent.Unprotect
    loLocation.Parent.Unprotect

    Set lr = loCatalog.ListRows.Add
    SetDesignsTestValue loCatalog, lr.Index, "SKU", "SKU-PICKER-1"
    SetDesignsTestValue loCatalog, lr.Index, "ROW", 96
    SetDesignsTestValue loCatalog, lr.Index, "ITEM_CODE", "SKU-PICKER-1"
    SetDesignsTestValue loCatalog, lr.Index, "ITEM", "Malawi Fine Cut Black Tea"
    SetDesignsTestValue loCatalog, lr.Index, "UOM", "LB"
    SetDesignsTestValue loCatalog, lr.Index, "LOCATION", "A1"

    Set lr = loBalance.ListRows.Add
    SetDesignsTestValue loBalance, lr.Index, "SKU", "SKU-PICKER-1"
    SetDesignsTestValue loBalance, lr.Index, "QtyOnHand", 3175

    Set lr = loLocation.ListRows.Add
    SetDesignsTestValue loLocation, lr.Index, "SKU", "SKU-PICKER-1"
    SetDesignsTestValue loLocation, lr.Index, "Location", "A1"
    SetDesignsTestValue loLocation, lr.Index, "QtyOnHand", 1000

    Set lr = loLocation.ListRows.Add
    SetDesignsTestValue loLocation, lr.Index, "SKU", "SKU-PICKER-1"
    SetDesignsTestValue loLocation, lr.Index, "Location", "CLEARVIEW"
    SetDesignsTestValue loLocation, lr.Index, "QtyOnHand", 2175

    items = modInventoryBridgeApi.ListInventoryPickerItemsBridgeResult("Malawi", wb)
    If IsEmpty(items) Or Not IsArray(items) Then GoTo CleanExit
    For r = LBound(items, 1) To UBound(items, 1)
        If CStr(items(r, 1)) = "96" And CDbl(items(r, 4)) = 3175 _
           And CStr(items(r, 7)) = "SKU-PICKER-1" Then
            If StrComp(CStr(items(r, 5)), "A1", vbTextCompare) = 0 Then foundA1 = True
            If StrComp(CStr(items(r, 5)), "CLEARVIEW", vbTextCompare) = 0 Then foundClearview = True
        End If
    Next r
    If foundA1 And foundClearview Then TestInventoryQueries_PickerPublishesEverySkuLocation = 1

CleanExit:
    CloseDesignsTestWorkbook wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestInventoryApply_ShipRejectsNegativeInventory() As Long
    Dim wb As Workbook
    Dim seedEvent As Object
    Dim shipEvent As Object
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim payloadJson As String
    Dim loLog As ListObject

    Set wb = TestPhase2Helpers.BuildPhase2InventoryWorkbook("WH-NEG-SHIP", Array("SKU-NEG-SHIP"), False)
    Set seedEvent = TestPhase2Helpers.CreateReceiveEvent( _
        "INV-NEG-SHIP-SEED", "WH-NEG-SHIP", "S-NEG", "U-NEG", "SKU-NEG-SHIP", 5, "CLEARVIEW", "seed")
    payloadJson = TestPhase2Helpers.BuildPayloadJson( _
        TestPhase2Helpers.CreatePayloadItem(1, "SKU-NEG-SHIP", 6, "CLEARVIEW", "overdraw"))
    Set shipEvent = TestPhase2Helpers.CreatePayloadEvent( _
        "INV-NEG-SHIP-APPLY", EVENT_TYPE_SHIP, "WH-NEG-SHIP", "S-NEG", "U-NEG", payloadJson)

    On Error GoTo CleanFail
    If Not modInventoryApply.ApplyEvent(seedEvent, wb, "RUN-NEG-SHIP-SEED", statusOut, errorCode, errorMessage) Then GoTo CleanExit
    statusOut = "": errorCode = "": errorMessage = ""
    If modInventoryApply.ApplyEvent(shipEvent, wb, "RUN-NEG-SHIP", statusOut, errorCode, errorMessage) Then GoTo CleanExit
    If errorCode <> "INSUFFICIENT_INVENTORY" Then GoTo CleanExit
    If InStr(1, errorMessage, EVENT_TYPE_SHIP, vbTextCompare) = 0 Then GoTo CleanExit
    Set loLog = FindDesignsTestTable(wb, "tblInventoryLog")
    If loLog Is Nothing Or loLog.ListRows.Count <> 1 Then GoTo CleanExit
    TestInventoryApply_ShipRejectsNegativeInventory = 1

CleanExit:
    CloseDesignsTestWorkbook wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestDesignMigration_BuildsDeterministicEventsWithoutMutatingDonor() As Long
    Dim donorWb As Workbook
    Dim designsWb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim planOne As Collection
    Dim planTwo As Collection
    Dim item As Object
    Dim itemTwo As Object
    Dim blackTeaItem As Object
    Dim evt As Object
    Dim report As String
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim r As Long
    Dim beforeRows As Long
    Dim beforeSheets As Long
    Dim beforeValue As String
    Dim loDesigns As ListObject
    Dim loLines As ListObject

    Set donorWb = Application.Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail
    Set ws = donorWb.Worksheets(1)
    ws.Name = "Recipes"
    ws.Range("A1:J1").Value = Array("RECIPE", "RECIPE_ID", "DESCRIPTION", "PROCESS", "INPUT/OUTPUT", "INGREDIENT", "INGREDIENT_ID", "AMOUNT", "UOM", "PERCENT")
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:J4"), , xlYes)
    lo.Name = "Recipes"
    lo.DataBodyRange.Rows(1).Value = Array("Brewed Black Tea", "TEA-BLACK", "Concentrate", 1, "USED", "Black Tea", "SKU-BLACK", 32.5, "LB", 100)
    lo.DataBodyRange.Rows(2).Value = Array("Brewed Black Tea", "TEA-BLACK", "Concentrate", 1, "OUTPUT", "Brew Black Tea", "OUT-BLACK", 400, "LB", 100)
    lo.DataBodyRange.Rows(3).Value = Array("Simple Syrup", "SYRUP-SIMPLE", "Syrup", 1, "OUTPUT", "Simple Syrup", "OUT-SYRUP", 100, "LB", 100)
    beforeRows = lo.ListRows.Count
    beforeSheets = donorWb.Worksheets.Count
    beforeValue = CStr(lo.DataBodyRange.Cells(1, 1).Value)

    Set planOne = modAdminDesignMigration.BuildLegacyRecipeDesignMigrationPlan(donorWb, "Recipes", report)
    If planOne Is Nothing Or planOne.Count <> 2 Then GoTo CleanExit
    Set planTwo = modAdminDesignMigration.BuildLegacyRecipeDesignMigrationPlan(donorWb, "Recipes", report)
    If planTwo Is Nothing Or planTwo.Count <> 2 Then GoTo CleanExit
    For r = 1 To planOne.Count
        Set item = planOne(r)
        Set itemTwo = planTwo(r)
        If CStr(item("EventID")) <> CStr(itemTwo("EventID")) Then GoTo CleanExit
        If InStr(1, CStr(item("MigrationSourceId")), donorWb.Name, vbTextCompare) = 0 Then GoTo CleanExit
        If StrComp(CStr(item("DesignId")), "TEA-BLACK", vbTextCompare) = 0 Then Set blackTeaItem = item
    Next r
    If blackTeaItem Is Nothing Then GoTo CleanExit
    If donorWb.Worksheets.Count <> beforeSheets Then GoTo CleanExit
    If lo.ListRows.Count <> beforeRows Then GoTo CleanExit
    If CStr(lo.DataBodyRange.Cells(1, 1).Value) <> beforeValue Then GoTo CleanExit

    Set designsWb = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modDesignsSchema.EnsureDesignsSchema(designsWb, report) Then GoTo CleanExit
    Set evt = BuildDesignsTestEvent(CStr(blackTeaItem("EventID")), _
                                    CStr(blackTeaItem("EventType")), _
                                    CStr(blackTeaItem("DesignId")), _
                                    CStr(blackTeaItem("DesignVersion")), _
                                    CStr(blackTeaItem("PayloadJson")))
    evt("MigrationSourceId") = CStr(blackTeaItem("MigrationSourceId"))
    If Not modDesignsApply.ApplyDesignEvent(evt, designsWb, "RUN-DES-MIG", statusOut, errorCode, errorMessage) Then GoTo CleanExit
    Set loDesigns = FindDesignsTestTable(designsWb, "tblDesigns")
    Set loLines = FindDesignsTestTable(designsWb, "tblDesignLines")
    If loDesigns Is Nothing Or loLines Is Nothing Then GoTo CleanExit
    If loDesigns.ListRows.Count <> 1 Or loLines.ListRows.Count <> 2 Then GoTo CleanExit
    TestDesignMigration_BuildsDeterministicEventsWithoutMutatingDonor = 1

CleanExit:
    CloseDesignsTestWorkbook designsWb
    CloseDesignsTestWorkbook donorWb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestInventoryApply_ProductionConsumeRejectsNegativeInventory() As Long
    Dim wb As Workbook
    Dim consumeEvent As Object
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim payloadJson As String
    Dim loLog As ListObject

    Set wb = TestPhase2Helpers.BuildPhase2InventoryWorkbook("WH-NEG-PROD", Array("SKU-NEG-PROD"), False)
    payloadJson = TestPhase2Helpers.BuildPayloadJson( _
        TestPhase2Helpers.CreatePayloadItem(1, "SKU-NEG-PROD", 1, "CLEARVIEW", "overdraw", "USED"))
    Set consumeEvent = TestPhase2Helpers.CreatePayloadEvent( _
        "INV-NEG-PROD-APPLY", EVENT_TYPE_PROD_CONSUME, "WH-NEG-PROD", "S-NEG", "U-NEG", payloadJson)

    On Error GoTo CleanFail
    If modInventoryApply.ApplyEvent(consumeEvent, wb, "RUN-NEG-PROD", statusOut, errorCode, errorMessage) Then GoTo CleanExit
    If errorCode <> "INSUFFICIENT_INVENTORY" Then GoTo CleanExit
    If InStr(1, errorMessage, EVENT_TYPE_PROD_CONSUME, vbTextCompare) = 0 Then GoTo CleanExit
    Set loLog = FindDesignsTestTable(wb, "tblInventoryLog")
    If loLog Is Nothing Then GoTo CleanExit
    If Not loLog.DataBodyRange Is Nothing Then
        If loLog.ListRows.Count > 0 Then GoTo CleanExit
    End If
    TestInventoryApply_ProductionConsumeRejectsNegativeInventory = 1

CleanExit:
    CloseDesignsTestWorkbook wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Private Function ReusableProcessPayloadForTest(ByVal processName As String, _
                                               ByVal requirementId As String, _
                                               ByVal acceptableItemCode As String, _
                                               ByVal requirementQty As Double, _
                                               ByVal requirementUom As String, _
                                               ByVal outputId As String, _
                                               ByVal outputItemCode As String, _
                                               ByVal outputQty As Double, _
                                               ByVal outputUom As String, _
                                               ByVal includeAlternative As Boolean) As String
    Dim payloadJson As String
    payloadJson = "[{""RecordType"":""PROCESS"",""ProcessName"":""" & processName & """}"
    If requirementId <> "" Then
        payloadJson = payloadJson & _
            ",{""RecordType"":""REQUIREMENT"",""RequirementId"":""" & _
            requirementId & """,""RequirementName"":""Input"",""Qty"":" & _
            Replace$(CStr(requirementQty), ",", ".") & ",""UOM"":""" & requirementUom & """}"
        If includeAlternative Then
            payloadJson = payloadJson & _
                ",{""RecordType"":""ALTERNATIVE"",""RequirementId"":""" & _
                requirementId & """,""ITEM_CODE"":""" & acceptableItemCode & """}"
        End If
    End If
    payloadJson = payloadJson & _
        ",{""RecordType"":""OUTPUT"",""OutputId"":""" & outputId & _
        """,""OutputName"":""Output"",""ITEM_CODE"":""" & outputItemCode & _
        """,""Qty"":" & Replace$(CStr(outputQty), ",", ".") & _
        ",""UOM"":""" & outputUom & """}]"
    ReusableProcessPayloadForTest = payloadJson
End Function

Private Function TwoNodeRecipePayloadForTest(ByVal recipeName As String, _
                                             ByVal sourceProcessId As String, _
                                             ByVal sinkProcessId As String, _
                                             ByVal connectionQty As Double, _
                                             ByVal connectionUom As String, _
                                             ByVal sourceOrdinal As Long, _
                                             ByVal sinkOrdinal As Long) As String
    TwoNodeRecipePayloadForTest = _
        "[{""RecordType"":""RECIPE"",""RecipeName"":""" & recipeName & """}," & _
        "{""RecordType"":""PROCESS_NODE"",""ProcessNodeId"":""A"",""ProcessId"":""" & sourceProcessId & _
        """,""ProcessVersion"":""1"",""ExecutionOrdinal"":" & CStr(sourceOrdinal) & "}," & _
        "{""RecordType"":""PROCESS_NODE"",""ProcessNodeId"":""B"",""ProcessId"":""" & sinkProcessId & _
        """,""ProcessVersion"":""1"",""ExecutionOrdinal"":" & CStr(sinkOrdinal) & "}," & _
        "{""RecordType"":""CONNECTION"",""FromProcessNodeId"":""A"",""FromOutputId"":""OUT-A"",""ToProcessNodeId"":""B"",""ToRequirementId"":""REQ-B"",""Qty"":" & _
        Replace$(CStr(connectionQty), ",", ".") & ",""UOM"":""" & connectionUom & """}]"
End Function

Private Function RunInvalidRecipeConnectionTest(ByVal token As String, _
                                                ByVal sourceQty As Double, _
                                                ByVal sourceUom As String, _
                                                ByVal sinkQty As Double, _
                                                ByVal sinkUom As String, _
                                                ByVal connectionQty As Double, _
                                                ByVal connectionUom As String, _
                                                ByVal sourceOrdinal As Long, _
                                                ByVal sinkOrdinal As Long, _
                                                ByVal expectedErrorCode As String) As Long
    Dim wb As Workbook
    Dim report As String
    Dim sourceJson As String
    Dim sinkJson As String
    Dim recipeJson As String
    Dim errorCode As String

    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    On Error GoTo CleanFail
    If Not modDesignsSchema.EnsureDesignsSchema(wb, report) Then GoTo CleanExit
    sourceJson = ReusableProcessPayloadForTest( _
        "Source", "", "", 0, "", "OUT-A", "SKU-A", sourceQty, sourceUom, False)
    sinkJson = ReusableProcessPayloadForTest( _
        "Sink", "REQ-B", "SKU-A", sinkQty, sinkUom, "OUT-B", "SKU-B", sinkQty, sinkUom, True)
    If Not SaveReleaseProcessForTest(wb, "PROC-" & token & "-A", _
            "PROC-" & token & "-A", "1", sourceJson, errorCode) Then GoTo CleanExit
    If Not SaveReleaseProcessForTest(wb, "PROC-" & token & "-B", _
            "PROC-" & token & "-B", "1", sinkJson, errorCode) Then GoTo CleanExit
    recipeJson = TwoNodeRecipePayloadForTest( _
        token, "PROC-" & token & "-A", "PROC-" & token & "-B", _
        connectionQty, connectionUom, sourceOrdinal, sinkOrdinal)
    If Not ApplyReusableEventForTest(wb, "RECIPE-" & token & "-SAVE", _
            "RECIPE_SAVE", "RECIPE-" & token, "1", recipeJson, errorCode) Then GoTo CleanExit
    If ApplyReusableEventForTest(wb, "RECIPE-" & token & "-RELEASE", _
            "RECIPE_RELEASE", "RECIPE-" & token, "1", "", errorCode) Then GoTo CleanExit
    If StrComp(errorCode, expectedErrorCode, vbTextCompare) <> 0 Then GoTo CleanExit
    RunInvalidRecipeConnectionTest = 1

CleanExit:
    CloseDesignsTestWorkbook wb
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Private Function SaveReleaseProcessForTest(ByVal wb As Workbook, ByVal token As String, _
                                           ByVal processId As String, _
                                           ByVal processVersion As String, _
                                           ByVal payloadJson As String, _
                                           ByRef errorCode As String) As Boolean
    If Not ApplyReusableEventForTest(wb, token & "-SAVE", "PROCESS_SAVE", _
            processId, processVersion, payloadJson, errorCode) Then Exit Function
    If Not ApplyReusableEventForTest(wb, token & "-RELEASE", "PROCESS_RELEASE", _
            processId, processVersion, "", errorCode) Then Exit Function
    SaveReleaseProcessForTest = True
End Function

Private Function ApplyReusableEventForTest(ByVal wb As Workbook, ByVal eventId As String, _
                                           ByVal eventType As String, ByVal definitionId As String, _
                                           ByVal definitionVersion As String, ByVal payloadJson As String, _
                                           ByRef errorCode As String) As Boolean
    Dim evt As Object
    Dim statusOut As String
    Dim errorMessage As String

    errorCode = ""
    Set evt = BuildDesignsTestEvent(eventId, eventType, definitionId, definitionVersion, payloadJson)
    ApplyReusableEventForTest = modDesignsApply.ApplyDesignEvent( _
        evt, wb, "RUN-" & eventId, statusOut, errorCode, errorMessage)
End Function

Private Function ReusableStatusForTest(ByVal wb As Workbook, ByVal tableName As String, _
                                       ByVal idColumn As String, ByVal versionColumn As String, _
                                       ByVal definitionId As String, _
                                       ByVal definitionVersion As String) As String
    Dim lo As ListObject
    Dim r As Long
    Set lo = FindDesignsTestTable(wb, tableName)
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    For r = 1 To lo.ListRows.Count
        If StrComp(CStr(lo.DataBodyRange.Cells(r, lo.ListColumns(idColumn).Index).Value2), _
                   definitionId, vbTextCompare) = 0 _
           And StrComp(CStr(lo.DataBodyRange.Cells(r, lo.ListColumns(versionColumn).Index).Value2), _
                       definitionVersion, vbTextCompare) = 0 Then
            ReusableStatusForTest = CStr( _
                lo.DataBodyRange.Cells(r, lo.ListColumns("Status").Index).Value2)
            Exit Function
        End If
    Next r
End Function

Private Function CountReusableRowsForTest(ByVal wb As Workbook, ByVal tableName As String, _
                                          ByVal idColumn As String, ByVal versionColumn As String, _
                                          ByVal definitionId As String, _
                                          ByVal definitionVersion As String) As Long
    Dim lo As ListObject
    Dim r As Long
    Set lo = FindDesignsTestTable(wb, tableName)
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    For r = 1 To lo.ListRows.Count
        If StrComp(CStr(lo.DataBodyRange.Cells(r, lo.ListColumns(idColumn).Index).Value2), _
                   definitionId, vbTextCompare) = 0 _
           And StrComp(CStr(lo.DataBodyRange.Cells(r, lo.ListColumns(versionColumn).Index).Value2), _
                       definitionVersion, vbTextCompare) = 0 Then
            CountReusableRowsForTest = CountReusableRowsForTest + 1
        End If
    Next r
End Function

Private Function FindDesignsTestTable(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        On Error Resume Next
        Set FindDesignsTestTable = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not FindDesignsTestTable Is Nothing Then Exit Function
    Next ws
End Function

Private Function DesignsTestColumnExists(ByVal lo As ListObject, ByVal columnName As String) As Boolean
    On Error Resume Next
    DesignsTestColumnExists = Not lo.ListColumns(columnName) Is Nothing
    On Error GoTo 0
End Function

Private Function IsUsableDesignsTestArray(ByVal values As Variant) As Boolean
    On Error GoTo CleanExit
    If IsEmpty(values) Or Not IsArray(values) Then Exit Function
    IsUsableDesignsTestArray = (UBound(values, 1) >= LBound(values, 1))
CleanExit:
End Function

Private Function BuildDesignsTestEvent(ByVal eventId As String, ByVal eventType As String, _
                                       ByVal designId As String, ByVal designVersion As String, _
                                       ByVal payloadJson As String) As Object
    Dim evt As Object
    Set evt = CreateObject("Scripting.Dictionary")
    evt.CompareMode = vbTextCompare
    evt("EventID") = eventId
    evt("EventType") = eventType
    evt("CreatedAtUTC") = Now
    evt("WarehouseId") = "WH-DES"
    evt("StationId") = "S-DES"
    evt("UserId") = "U-DES"
    evt("DesignId") = designId
    evt("DesignVersion") = designVersion
    evt("PayloadJson") = payloadJson
    evt("SourceInbox") = "TEST"
    Set BuildDesignsTestEvent = evt
End Function

Private Function CountDesignsTestTables(ByVal wb As Workbook) As Long
    Dim ws As Worksheet
    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        CountDesignsTestTables = CountDesignsTestTables + ws.ListObjects.Count
    Next ws
End Function

Private Sub SetDesignsTestValue(ByVal lo As ListObject, ByVal rowIndex As Long, _
                                ByVal columnName As String, ByVal valueOut As Variant)
    lo.DataBodyRange.Cells(rowIndex, lo.ListColumns(columnName).Index).Value = valueOut
End Sub

Private Sub CloseDesignsTestWorkbook(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub
    Application.DisplayAlerts = False
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
End Sub
