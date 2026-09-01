Attribute VB_Name = "modDesignsQueries"
Option Explicit

Public Function ListDesigns(Optional ByVal designsWb As Workbook = Nothing, _
                            Optional ByVal statusFilter As String = "") As Variant
    On Error GoTo FailQuery

    Dim wb As Workbook
    Dim lo As ListObject
    Dim src As Variant
    Dim result() As Variant
    Dim trimmed() As Variant
    Dim r As Long
    Dim c As Long
    Dim outRow As Long
    Dim report As String

    Set wb = modDesignsRuntime.ResolveDesignsWorkbook("", designsWb, report)
    If wb Is Nothing Then Exit Function
    Set lo = FindDesignsTableQuery(wb, "tblDesigns")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function

    src = lo.DataBodyRange.Value
    ReDim result(1 To UBound(src, 1), 1 To 6)
    For r = 1 To UBound(src, 1)
        If statusFilter = "" Or StrComp(ReadDesignCellQuery(lo, src, r, "Status"), statusFilter, vbTextCompare) = 0 Then
            outRow = outRow + 1
            result(outRow, 1) = ReadDesignCellQuery(lo, src, r, "DesignId")
            result(outRow, 2) = ReadDesignCellQuery(lo, src, r, "DesignVersion")
            result(outRow, 3) = ReadDesignCellQuery(lo, src, r, "DesignType")
            result(outRow, 4) = ReadDesignCellQuery(lo, src, r, "DesignName")
            result(outRow, 5) = ReadDesignCellQuery(lo, src, r, "Description")
            result(outRow, 6) = ReadDesignCellQuery(lo, src, r, "Status")
        End If
    Next r
    If outRow = 0 Then Exit Function
    ReDim trimmed(1 To outRow, 1 To 6)
    For r = 1 To outRow
        For c = 1 To 6
            trimmed(r, c) = result(r, c)
        Next c
    Next r
    ListDesigns = trimmed
    Exit Function

FailQuery:
    ListDesigns = Empty
End Function

Public Function GetBOM(ByVal designId As String, ByVal designVersion As String, _
                       Optional ByVal designsWb As Workbook = Nothing) As Variant
    On Error GoTo FailQuery

    Dim wb As Workbook
    Dim lo As ListObject
    Dim src As Variant
    Dim result() As Variant
    Dim trimmed() As Variant
    Dim r As Long
    Dim c As Long
    Dim outRow As Long
    Dim report As String

    Set wb = modDesignsRuntime.ResolveDesignsWorkbook("", designsWb, report)
    If wb Is Nothing Then Exit Function
    Set lo = FindDesignsTableQuery(wb, "tblDesignLines")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    src = lo.DataBodyRange.Value
    ReDim result(1 To UBound(src, 1), 1 To 10)
    For r = 1 To UBound(src, 1)
        If DesignIdsMatchQuery(ReadDesignCellQuery(lo, src, r, "DesignId"), designId) _
           And StrComp(ReadDesignCellQuery(lo, src, r, "DesignVersion"), Trim$(designVersion), vbTextCompare) = 0 Then
            outRow = outRow + 1
            result(outRow, 1) = ReadDesignCellQuery(lo, src, r, "LineNo")
            result(outRow, 2) = ReadDesignCellQuery(lo, src, r, "Process")
            result(outRow, 3) = ReadDesignCellQuery(lo, src, r, "IOType")
            result(outRow, 4) = ReadDesignCellQuery(lo, src, r, "ComponentSKU")
            result(outRow, 5) = ReadDesignCellQuery(lo, src, r, "ComponentDesignId")
            result(outRow, 6) = ReadDesignCellQuery(lo, src, r, "ComponentDesignVersion")
            result(outRow, 7) = ReadDesignCellQuery(lo, src, r, "Qty")
            result(outRow, 8) = ReadDesignCellQuery(lo, src, r, "UOM")
            result(outRow, 9) = ReadDesignCellQuery(lo, src, r, "Percent")
            result(outRow, 10) = ReadDesignCellQuery(lo, src, r, "Instruction")
        End If
    Next r
    If outRow = 0 Then Exit Function
    SortBomRowsByLineNo result, outRow
    ReDim trimmed(1 To outRow, 1 To 10)
    For r = 1 To outRow
        For c = 1 To 10
            trimmed(r, c) = result(r, c)
        Next c
    Next r
    GetBOM = trimmed
    Exit Function

FailQuery:
    GetBOM = Empty
End Function

Public Function ListProcesses(Optional ByVal designsWb As Workbook = Nothing, _
                              Optional ByVal statusFilter As String = "") As Variant
    Dim queryResult As Variant
    queryResult = ListReusableDefinitionsQuery(designsWb, "tblProcesses", _
        "ProcessId", "ProcessVersion", "ProcessName", statusFilter)
    ListProcesses = queryResult
End Function

Public Function GetProcessVersion(ByVal processId As String, _
                                  ByVal processVersion As String, _
                                  Optional ByVal designsWb As Workbook = Nothing) As String
    On Error GoTo FailQuery

    Dim wb As Workbook
    Dim lo As ListObject
    Dim values As Variant
    Dim report As String
    Dim jsonOut As String
    Dim recordCount As Long
    Dim r As Long

    processId = Trim$(processId)
    processVersion = Trim$(processVersion)
    If processId = "" Or processVersion = "" Then Exit Function
    Set wb = modDesignsRuntime.ResolveDesignsWorkbook("", designsWb, report)
    If wb Is Nothing Then Exit Function
    Set lo = FindDesignsTableQuery(wb, "tblProcesses")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    values = lo.DataBodyRange.Value2
    jsonOut = "["
    For r = 1 To UBound(values, 1)
        If ReusableRowMatchesQuery(lo, values, r, "ProcessId", _
                "ProcessVersion", processId, processVersion) Then
            AppendJsonRecordQuery jsonOut, recordCount, _
                "{""RecordType"":""PROCESS""," & _
                JsonTextPairQuery("ProcessId", ReadDesignCellQuery(lo, values, r, "ProcessId")) & "," & _
                JsonTextPairQuery("ProcessVersion", ReadDesignCellQuery(lo, values, r, "ProcessVersion")) & "," & _
                JsonTextPairQuery("ProcessName", ReadDesignCellQuery(lo, values, r, "ProcessName")) & "," & _
                JsonTextPairQuery("Description", ReadDesignCellQuery(lo, values, r, "Description")) & "," & _
                JsonTextPairQuery("Status", ReadDesignCellQuery(lo, values, r, "Status")) & "}"
            Exit For
        End If
    Next r
    If recordCount = 0 Then Exit Function
    AppendProcessRequirementsQuery wb, processId, processVersion, jsonOut, recordCount
    AppendProcessAlternativesQuery wb, processId, processVersion, jsonOut, recordCount
    AppendProcessOutputsQuery wb, processId, processVersion, jsonOut, recordCount
    AppendProcessInstructionsQuery wb, processId, processVersion, jsonOut, recordCount
    GetProcessVersion = jsonOut & "]"
    Exit Function

FailQuery:
    GetProcessVersion = ""
End Function

Public Function ListRecipes(Optional ByVal designsWb As Workbook = Nothing, _
                            Optional ByVal statusFilter As String = "") As Variant
    ListRecipes = ListReusableDefinitionsQuery(designsWb, "tblRecipes", _
        "RecipeId", "RecipeVersion", "RecipeName", statusFilter)
End Function

Public Function GetRecipeGraph(ByVal recipeId As String, _
                               ByVal recipeVersion As String, _
                               Optional ByVal designsWb As Workbook = Nothing) As String
    On Error GoTo FailQuery

    Dim wb As Workbook
    Dim lo As ListObject
    Dim values As Variant
    Dim report As String
    Dim jsonOut As String
    Dim recordCount As Long
    Dim r As Long

    recipeId = Trim$(recipeId)
    recipeVersion = Trim$(recipeVersion)
    If recipeId = "" Or recipeVersion = "" Then Exit Function
    Set wb = modDesignsRuntime.ResolveDesignsWorkbook("", designsWb, report)
    If wb Is Nothing Then Exit Function
    Set lo = FindDesignsTableQuery(wb, "tblRecipes")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    values = lo.DataBodyRange.Value2
    jsonOut = "["
    For r = 1 To UBound(values, 1)
        If ReusableRowMatchesQuery(lo, values, r, "RecipeId", _
                "RecipeVersion", recipeId, recipeVersion) Then
            AppendJsonRecordQuery jsonOut, recordCount, _
                "{""RecordType"":""RECIPE""," & _
                JsonTextPairQuery("RecipeId", ReadDesignCellQuery(lo, values, r, "RecipeId")) & "," & _
                JsonTextPairQuery("RecipeVersion", ReadDesignCellQuery(lo, values, r, "RecipeVersion")) & "," & _
                JsonTextPairQuery("RecipeName", ReadDesignCellQuery(lo, values, r, "RecipeName")) & "," & _
                JsonTextPairQuery("Description", ReadDesignCellQuery(lo, values, r, "Description")) & "," & _
                JsonTextPairQuery("Status", ReadDesignCellQuery(lo, values, r, "Status")) & "}"
            Exit For
        End If
    Next r
    If recordCount = 0 Then Exit Function
    AppendRecipeProcessesQuery wb, recipeId, recipeVersion, jsonOut, recordCount
    AppendRecipeConnectionsQuery wb, recipeId, recipeVersion, jsonOut, recordCount
    AppendRecipeOutputRegulationsQuery wb, recipeId, recipeVersion, jsonOut, recordCount
    GetRecipeGraph = jsonOut & "]"
    Exit Function

FailQuery:
    GetRecipeGraph = ""
End Function

Public Function ValidateReleasedRecipe(ByVal recipeId As String, _
                                       ByVal recipeVersion As String, _
                                       Optional ByVal designsWb As Workbook = Nothing, _
                                       Optional ByRef errorCode As String = "", _
                                       Optional ByRef errorMessage As String = "") As Boolean
    On Error GoTo FailQuery

    Dim wb As Workbook
    Dim report As String
    Dim statusValue As String

    recipeId = Trim$(recipeId)
    recipeVersion = Trim$(recipeVersion)
    If recipeId = "" Or recipeVersion = "" Then
        errorCode = "RECIPE_IDENTITY_REQUIRED"
        errorMessage = "Recipe ID and version are required."
        Exit Function
    End If
    Set wb = modDesignsRuntime.ResolveDesignsWorkbook("", designsWb, report)
    If wb Is Nothing Then
        errorCode = "DESIGNS_WORKBOOK_UNAVAILABLE"
        errorMessage = report
        Exit Function
    End If
    statusValue = ReusableDefinitionStatusQuery(wb, "tblRecipes", _
        "RecipeId", "RecipeVersion", recipeId, recipeVersion)
    If StrComp(statusValue, "RELEASED", vbTextCompare) <> 0 Then
        errorCode = "RECIPE_NOT_RELEASED"
        errorMessage = "The requested Recipe version is not released."
        Exit Function
    End If
    ValidateReleasedRecipe = modDesignsApply.ValidateRecipeReleaseContract( _
        wb, recipeId, recipeVersion, errorCode, errorMessage)
    Exit Function

FailQuery:
    errorCode = "RECIPE_VALIDATION_FAILED"
    errorMessage = Err.Description
End Function

Private Function ListReusableDefinitionsQuery(ByVal designsWb As Workbook, _
                                              ByVal tableName As String, _
                                              ByVal idColumn As String, _
                                              ByVal versionColumn As String, _
                                              ByVal nameColumn As String, _
                                              ByVal statusFilter As String) As Variant
    On Error GoTo FailQuery

    Dim wb As Workbook
    Dim lo As ListObject
    Dim src As Variant
    Dim result() As Variant
    Dim trimmed() As Variant
    Dim report As String
    Dim r As Long
    Dim c As Long
    Dim outRow As Long

    Set wb = modDesignsRuntime.ResolveDesignsWorkbook("", designsWb, report)
    If wb Is Nothing Then Exit Function
    Set lo = FindDesignsTableQuery(wb, tableName)
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    src = lo.DataBodyRange.Value2
    ReDim result(1 To UBound(src, 1), 1 To 6)
    For r = 1 To UBound(src, 1)
        If statusFilter = "" Or StrComp(ReadDesignCellQuery(lo, src, r, "Status"), _
                statusFilter, vbTextCompare) = 0 Then
            outRow = outRow + 1
            result(outRow, 1) = ReadDesignCellQuery(lo, src, r, idColumn)
            result(outRow, 2) = ReadDesignCellQuery(lo, src, r, versionColumn)
            result(outRow, 3) = ReadDesignCellQuery(lo, src, r, nameColumn)
            result(outRow, 4) = ReadDesignCellQuery(lo, src, r, "Description")
            result(outRow, 5) = ReadDesignCellQuery(lo, src, r, "Status")
            result(outRow, 6) = ReadDesignCellQuery(lo, src, r, "ReleasedAtUTC")
        End If
    Next r
    If outRow = 0 Then Exit Function
    ReDim trimmed(1 To outRow, 1 To 6)
    For r = 1 To outRow
        For c = 1 To 6
            trimmed(r, c) = result(r, c)
        Next c
    Next r
    ListReusableDefinitionsQuery = trimmed
    Exit Function

FailQuery:
    ListReusableDefinitionsQuery = Empty
End Function

Private Sub AppendProcessRequirementsQuery(ByVal wb As Workbook, _
                                           ByVal processId As String, _
                                           ByVal processVersion As String, _
                                           ByRef jsonOut As String, _
                                           ByRef recordCount As Long)
    Dim lo As ListObject
    Dim values As Variant
    Dim r As Long
    Set lo = FindDesignsTableQuery(wb, "tblProcessRequirements")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    values = lo.DataBodyRange.Value2
    For r = 1 To UBound(values, 1)
        If ReusableRowMatchesQuery(lo, values, r, "ProcessId", _
                "ProcessVersion", processId, processVersion) Then
            AppendJsonRecordQuery jsonOut, recordCount, _
                "{""RecordType"":""REQUIREMENT""," & _
                JsonTextPairQuery("RequirementId", ReadDesignCellQuery(lo, values, r, "RequirementId")) & "," & _
                JsonTextPairQuery("RequirementName", ReadDesignCellQuery(lo, values, r, "RequirementName")) & "," & _
                JsonValuePairQuery("Qty", ReadDesignValueQuery(lo, values, r, "Qty")) & "," & _
                JsonValuePairQuery("Percent", ReadDesignValueQuery(lo, values, r, "Percent")) & "," & _
                JsonTextPairQuery("YieldBasis", ReadDesignCellQuery(lo, values, r, "YieldBasis")) & "," & _
                JsonTextPairQuery("UOM", ReadDesignCellQuery(lo, values, r, "UOM")) & "," & _
                JsonTextPairQuery("RequirementQtyMode", ReadDesignCellQuery(lo, values, r, "RequirementQtyMode")) & "}"
        End If
    Next r
End Sub

Private Sub AppendProcessAlternativesQuery(ByVal wb As Workbook, _
                                           ByVal processId As String, _
                                           ByVal processVersion As String, _
                                           ByRef jsonOut As String, _
                                           ByRef recordCount As Long)
    Dim lo As ListObject
    Dim values As Variant
    Dim r As Long
    Set lo = FindDesignsTableQuery(wb, "tblProcessIngredientAlternatives")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    values = lo.DataBodyRange.Value2
    For r = 1 To UBound(values, 1)
        If ReusableRowMatchesQuery(lo, values, r, "ProcessId", _
                "ProcessVersion", processId, processVersion) Then
            AppendJsonRecordQuery jsonOut, recordCount, _
                "{""RecordType"":""ALTERNATIVE""," & _
                JsonTextPairQuery("RequirementId", ReadDesignCellQuery(lo, values, r, "RequirementId")) & "," & _
                JsonValuePairQuery("AlternativeOrdinal", ReadDesignValueQuery(lo, values, r, "AlternativeOrdinal")) & "," & _
                JsonTextPairQuery("ITEM_CODE", ReadDesignCellQuery(lo, values, r, "ITEM_CODE")) & "}"
        End If
    Next r
End Sub

Private Sub AppendProcessOutputsQuery(ByVal wb As Workbook, _
                                      ByVal processId As String, _
                                      ByVal processVersion As String, _
                                      ByRef jsonOut As String, _
                                      ByRef recordCount As Long)
    Dim lo As ListObject
    Dim values As Variant
    Dim r As Long
    Set lo = FindDesignsTableQuery(wb, "tblProcessOutputs")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    values = lo.DataBodyRange.Value2
    For r = 1 To UBound(values, 1)
        If ReusableRowMatchesQuery(lo, values, r, "ProcessId", _
                "ProcessVersion", processId, processVersion) Then
            AppendJsonRecordQuery jsonOut, recordCount, _
                "{""RecordType"":""OUTPUT""," & _
                JsonTextPairQuery("OutputId", ReadDesignCellQuery(lo, values, r, "OutputId")) & "," & _
                JsonTextPairQuery("OutputName", ReadDesignCellQuery(lo, values, r, "OutputName")) & "," & _
                JsonTextPairQuery("ITEM_CODE", ReadDesignCellQuery(lo, values, r, "ITEM_CODE")) & "," & _
                JsonTextPairQuery("ComponentDesignId", ReadDesignCellQuery(lo, values, r, "ComponentDesignId")) & "," & _
                JsonTextPairQuery("ComponentDesignVersion", ReadDesignCellQuery(lo, values, r, "ComponentDesignVersion")) & "," & _
                JsonValuePairQuery("Qty", ReadDesignValueQuery(lo, values, r, "Qty")) & "," & _
                JsonValuePairQuery("Percent", ReadDesignValueQuery(lo, values, r, "Percent")) & "," & _
                JsonTextPairQuery("YieldBasis", ReadDesignCellQuery(lo, values, r, "YieldBasis")) & "," & _
                JsonTextPairQuery("UOM", ReadDesignCellQuery(lo, values, r, "UOM")) & "," & _
                JsonTextPairQuery("OutputQtyMode", ReadDesignCellQuery(lo, values, r, "OutputQtyMode")) & "," & _
                JsonValuePairQuery("OutputRegulationEnabled", ReadDesignValueQuery(lo, values, r, "OutputRegulationEnabled")) & "," & _
                JsonValuePairQuery("OutputFloorQty", ReadDesignValueQuery(lo, values, r, "OutputFloorQty")) & "," & _
                JsonValuePairQuery("OutputCeilingQty", ReadDesignValueQuery(lo, values, r, "OutputCeilingQty")) & "}"
        End If
    Next r
End Sub

Private Sub AppendProcessInstructionsQuery(ByVal wb As Workbook, _
                                           ByVal processId As String, _
                                           ByVal processVersion As String, _
                                           ByRef jsonOut As String, _
                                           ByRef recordCount As Long)
    Dim lo As ListObject
    Dim values As Variant
    Dim r As Long
    Set lo = FindDesignsTableQuery(wb, "tblProcessInstructions")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    values = lo.DataBodyRange.Value2
    For r = 1 To UBound(values, 1)
        If ReusableRowMatchesQuery(lo, values, r, "ProcessId", _
                "ProcessVersion", processId, processVersion) Then
            AppendJsonRecordQuery jsonOut, recordCount, _
                "{""RecordType"":""INSTRUCTION""," & _
                JsonValuePairQuery("InstructionOrdinal", ReadDesignValueQuery(lo, values, r, "InstructionOrdinal")) & "," & _
                JsonTextPairQuery("Instruction", ReadDesignCellQuery(lo, values, r, "Instruction")) & "}"
        End If
    Next r
End Sub

Private Sub AppendRecipeProcessesQuery(ByVal wb As Workbook, _
                                       ByVal recipeId As String, _
                                       ByVal recipeVersion As String, _
                                       ByRef jsonOut As String, _
                                       ByRef recordCount As Long)
    Dim lo As ListObject
    Dim values As Variant
    Dim r As Long
    Set lo = FindDesignsTableQuery(wb, "tblRecipeProcesses")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    values = lo.DataBodyRange.Value2
    For r = 1 To UBound(values, 1)
        If ReusableRowMatchesQuery(lo, values, r, "RecipeId", _
                "RecipeVersion", recipeId, recipeVersion) Then
            AppendJsonRecordQuery jsonOut, recordCount, _
                "{""RecordType"":""PROCESS_NODE""," & _
                JsonTextPairQuery("ProcessNodeId", ReadDesignCellQuery(lo, values, r, "ProcessNodeId")) & "," & _
                JsonTextPairQuery("ProcessId", ReadDesignCellQuery(lo, values, r, "ProcessId")) & "," & _
                JsonTextPairQuery("ProcessVersion", ReadDesignCellQuery(lo, values, r, "ProcessVersion")) & "," & _
                JsonValuePairQuery("ExecutionOrdinal", ReadDesignValueQuery(lo, values, r, "ExecutionOrdinal")) & "}"
        End If
    Next r
End Sub

Private Sub AppendRecipeConnectionsQuery(ByVal wb As Workbook, _
                                         ByVal recipeId As String, _
                                         ByVal recipeVersion As String, _
                                         ByRef jsonOut As String, _
                                         ByRef recordCount As Long)
    Dim lo As ListObject
    Dim values As Variant
    Dim r As Long
    Set lo = FindDesignsTableQuery(wb, "tblRecipeConnections")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    values = lo.DataBodyRange.Value2
    For r = 1 To UBound(values, 1)
        If ReusableRowMatchesQuery(lo, values, r, "RecipeId", _
                "RecipeVersion", recipeId, recipeVersion) Then
            AppendJsonRecordQuery jsonOut, recordCount, _
                "{""RecordType"":""CONNECTION""," & _
                JsonTextPairQuery("FromProcessNodeId", ReadDesignCellQuery(lo, values, r, "FromProcessNodeId")) & "," & _
                JsonTextPairQuery("FromOutputId", ReadDesignCellQuery(lo, values, r, "FromOutputId")) & "," & _
                JsonTextPairQuery("ToProcessNodeId", ReadDesignCellQuery(lo, values, r, "ToProcessNodeId")) & "," & _
                JsonTextPairQuery("ToRequirementId", ReadDesignCellQuery(lo, values, r, "ToRequirementId")) & "," & _
                JsonValuePairQuery("Qty", ReadDesignValueQuery(lo, values, r, "Qty")) & "," & _
                JsonValuePairQuery("Percent", ReadDesignValueQuery(lo, values, r, "Percent")) & "," & _
                JsonTextPairQuery("UOM", ReadDesignCellQuery(lo, values, r, "UOM")) & "}"
        End If
    Next r
End Sub

Private Sub AppendRecipeOutputRegulationsQuery(ByVal wb As Workbook, _
                                                ByVal recipeId As String, _
                                                ByVal recipeVersion As String, _
                                                ByRef jsonOut As String, _
                                                ByRef recordCount As Long)
    Dim lo As ListObject
    Dim values As Variant
    Dim r As Long
    Set lo = FindDesignsTableQuery(wb, "tblRecipeOutputRegulations")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    values = lo.DataBodyRange.Value2
    For r = 1 To UBound(values, 1)
        If ReusableRowMatchesQuery(lo, values, r, "RecipeId", _
                "RecipeVersion", recipeId, recipeVersion) Then
            AppendJsonRecordQuery jsonOut, recordCount, _
                "{""RecordType"":""OUTPUT_REGULATION""," & _
                JsonTextPairQuery("ProcessNodeId", ReadDesignCellQuery(lo, values, r, "ProcessNodeId")) & "," & _
                JsonTextPairQuery("ProcessId", ReadDesignCellQuery(lo, values, r, "ProcessId")) & "," & _
                JsonTextPairQuery("ProcessVersion", ReadDesignCellQuery(lo, values, r, "ProcessVersion")) & "," & _
                JsonTextPairQuery("OutputId", ReadDesignCellQuery(lo, values, r, "OutputId")) & "," & _
                JsonValuePairQuery("OutputRegulationEnabled", ReadDesignValueQuery(lo, values, r, "OutputRegulationEnabled")) & "," & _
                JsonValuePairQuery("OutputFloorQty", ReadDesignValueQuery(lo, values, r, "OutputFloorQty")) & "," & _
                JsonValuePairQuery("OutputCeilingQty", ReadDesignValueQuery(lo, values, r, "OutputCeilingQty")) & "}"
        End If
    Next r
End Sub

Private Function ReusableDefinitionStatusQuery(ByVal wb As Workbook, _
                                               ByVal tableName As String, _
                                               ByVal idColumn As String, _
                                               ByVal versionColumn As String, _
                                               ByVal definitionId As String, _
                                               ByVal definitionVersion As String) As String
    Dim lo As ListObject
    Dim values As Variant
    Dim r As Long
    Set lo = FindDesignsTableQuery(wb, tableName)
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    values = lo.DataBodyRange.Value2
    For r = 1 To UBound(values, 1)
        If ReusableRowMatchesQuery(lo, values, r, idColumn, versionColumn, _
                definitionId, definitionVersion) Then
            ReusableDefinitionStatusQuery = ReadDesignCellQuery(lo, values, r, "Status")
            Exit Function
        End If
    Next r
End Function

Private Function ReusableRowMatchesQuery(ByVal lo As ListObject, _
                                         ByVal values As Variant, _
                                         ByVal rowIndex As Long, _
                                         ByVal idColumn As String, _
                                         ByVal versionColumn As String, _
                                         ByVal definitionId As String, _
                                         ByVal definitionVersion As String) As Boolean
    ReusableRowMatchesQuery = _
        (StrComp(ReadDesignCellQuery(lo, values, rowIndex, idColumn), _
                 definitionId, vbTextCompare) = 0) _
        And (StrComp(ReadDesignCellQuery(lo, values, rowIndex, versionColumn), _
                     definitionVersion, vbTextCompare) = 0)
End Function

Private Sub AppendJsonRecordQuery(ByRef jsonOut As String, _
                                  ByRef recordCount As Long, _
                                  ByVal recordJson As String)
    If recordCount > 0 Then jsonOut = jsonOut & ","
    jsonOut = jsonOut & recordJson
    recordCount = recordCount + 1
End Sub

Private Function JsonTextPairQuery(ByVal name As String, ByVal valueIn As String) As String
    JsonTextPairQuery = """" & JsonEscapeQuery(name) & """:""" & _
        JsonEscapeQuery(valueIn) & """"
End Function

Private Function JsonValuePairQuery(ByVal name As String, ByVal valueIn As Variant) As String
    Dim valueText As String
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) _
       Or Trim$(CStr(valueIn)) = "" Then
        valueText = "null"
    ElseIf IsNumeric(valueIn) Then
        valueText = Replace$(CStr(CDbl(valueIn)), ",", ".")
    Else
        valueText = """" & JsonEscapeQuery(CStr(valueIn)) & """"
    End If
    JsonValuePairQuery = """" & JsonEscapeQuery(name) & """:" & valueText
End Function

Private Function JsonEscapeQuery(ByVal valueIn As String) As String
    JsonEscapeQuery = Replace$(valueIn, "\", "\\")
    JsonEscapeQuery = Replace$(JsonEscapeQuery, Chr$(34), "\" & Chr$(34))
    JsonEscapeQuery = Replace$(JsonEscapeQuery, vbCrLf, "\n")
    JsonEscapeQuery = Replace$(JsonEscapeQuery, vbCr, "\n")
    JsonEscapeQuery = Replace$(JsonEscapeQuery, vbLf, "\n")
    JsonEscapeQuery = Replace$(JsonEscapeQuery, vbTab, "\t")
End Function

Private Function ReadDesignValueQuery(ByVal lo As ListObject, ByVal values As Variant, _
                                      ByVal rowIndex As Long, ByVal columnName As String) As Variant
    Dim columnIndex As Long
    On Error Resume Next
    columnIndex = lo.ListColumns(columnName).Index
    On Error GoTo 0
    If columnIndex = 0 Then Exit Function
    If IsError(values(rowIndex, columnIndex)) Or IsNull(values(rowIndex, columnIndex)) Then Exit Function
    ReadDesignValueQuery = values(rowIndex, columnIndex)
End Function

Private Sub SortBomRowsByLineNo(ByRef values As Variant, ByVal rowCount As Long)
    Dim i As Long
    Dim j As Long
    Dim c As Long
    Dim leftLine As Double
    Dim rightLine As Double
    Dim swapValue As Variant

    For i = 1 To rowCount - 1
        For j = i + 1 To rowCount
            leftLine = DesignLineSortValueQuery(values(i, 1), i)
            rightLine = DesignLineSortValueQuery(values(j, 1), j)
            If rightLine < leftLine Then
                For c = 1 To 10
                    swapValue = values(i, c)
                    values(i, c) = values(j, c)
                    values(j, c) = swapValue
                Next c
            End If
        Next j
    Next i
End Sub

Private Function DesignLineSortValueQuery(ByVal lineValue As Variant, ByVal fallbackRow As Long) As Double
    If IsNumeric(lineValue) Then
        DesignLineSortValueQuery = CDbl(lineValue)
    Else
        DesignLineSortValueQuery = 1000000000# + fallbackRow
    End If
End Function

Public Function GetBOMForStatus(ByVal designId As String, _
                                ByVal designVersion As String, _
                                ByVal requiredStatus As String, _
                                Optional ByVal designsWb As Workbook = Nothing) As Variant
    On Error GoTo FailQuery

    Dim wb As Workbook
    Dim lo As ListObject
    Dim src As Variant
    Dim r As Long
    Dim report As String

    designId = Trim$(designId)
    designVersion = Trim$(designVersion)
    requiredStatus = UCase$(Trim$(requiredStatus))
    If designId = "" Or designVersion = "" Or requiredStatus = "" Then Exit Function

    Set wb = modDesignsRuntime.ResolveDesignsWorkbook("", designsWb, report)
    If wb Is Nothing Then Exit Function
    Set lo = FindDesignsTableQuery(wb, "tblDesigns")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function

    src = lo.DataBodyRange.Value
    For r = 1 To UBound(src, 1)
        If DesignIdsMatchQuery(ReadDesignCellQuery(lo, src, r, "DesignId"), designId) _
           And StrComp(ReadDesignCellQuery(lo, src, r, "DesignVersion"), designVersion, vbTextCompare) = 0 Then
            If StrComp(ReadDesignCellQuery(lo, src, r, "Status"), requiredStatus, vbTextCompare) <> 0 Then Exit Function
            GetBOMForStatus = GetBOM(designId, designVersion, wb)
            Exit Function
        End If
    Next r
    Exit Function

FailQuery:
    GetBOMForStatus = Empty
End Function

Private Function DesignIdsMatchQuery(ByVal leftId As String, ByVal rightId As String) As Boolean
    DesignIdsMatchQuery = (StrComp(CanonicalDesignIdQuery(leftId), _
                                   CanonicalDesignIdQuery(rightId), vbTextCompare) = 0)
End Function

Private Function CanonicalDesignIdQuery(ByVal valueIn As String) As String
    Dim textValue As String
    Dim numericValue As Long

    textValue = UCase$(Trim$(valueIn))
    If textValue = "" Then Exit Function
    If Len(textValue) <= 3 And IsNumeric(textValue) Then
        numericValue = CLng(CDbl(textValue))
        If numericValue >= 0 And numericValue <= 999 Then
            CanonicalDesignIdQuery = Right$("000" & CStr(numericValue), 3)
            Exit Function
        End If
    End If
    CanonicalDesignIdQuery = textValue
End Function

Private Function FindDesignsTableQuery(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        On Error Resume Next
        Set FindDesignsTableQuery = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not FindDesignsTableQuery Is Nothing Then Exit Function
    Next ws
End Function

Private Function ReadDesignCellQuery(ByVal lo As ListObject, ByVal values As Variant, _
                                     ByVal rowIndex As Long, ByVal columnName As String) As String
    Dim columnIndex As Long
    On Error Resume Next
    columnIndex = lo.ListColumns(columnName).Index
    On Error GoTo 0
    If columnIndex = 0 Then Exit Function
    If IsError(values(rowIndex, columnIndex)) Or IsNull(values(rowIndex, columnIndex)) Then Exit Function
    ReadDesignCellQuery = CStr(values(rowIndex, columnIndex))
End Function
