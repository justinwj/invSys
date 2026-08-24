Attribute VB_Name = "modDesignsSchema"
Option Explicit

Private Const SHEET_DESIGNS As String = "Designs"
Private Const SHEET_LINES As String = "DesignLines"
Private Const SHEET_EVENTS As String = "DesignEvents"
Private Const SHEET_APPLIED As String = "AppliedDesignEvents"
Private Const SHEET_LOCKS As String = "Locks"
Private Const SHEET_PROCESSES As String = "Processes"
Private Const SHEET_PROCESS_REQUIREMENTS As String = "ProcessRequirements"
Private Const SHEET_PROCESS_ALTERNATIVES As String = "ProcessAlternatives"
Private Const SHEET_PROCESS_OUTPUTS As String = "ProcessOutputs"
Private Const SHEET_PROCESS_INSTRUCTIONS As String = "ProcessInstructions"
Private Const SHEET_RECIPES As String = "Recipes"
Private Const SHEET_RECIPE_PROCESSES As String = "RecipeProcesses"
Private Const SHEET_RECIPE_CONNECTIONS As String = "RecipeConnections"

Private Const TABLE_DESIGNS As String = "tblDesigns"
Private Const TABLE_LINES As String = "tblDesignLines"
Private Const TABLE_EVENTS As String = "tblDesignEvents"
Private Const TABLE_APPLIED As String = "tblAppliedDesignEvents"
Private Const TABLE_LOCKS As String = "tblLocks"
Private Const TABLE_PROCESSES As String = "tblProcesses"
Private Const TABLE_PROCESS_REQUIREMENTS As String = "tblProcessRequirements"
Private Const TABLE_PROCESS_ALTERNATIVES As String = "tblProcessIngredientAlternatives"
Private Const TABLE_PROCESS_OUTPUTS As String = "tblProcessOutputs"
Private Const TABLE_PROCESS_INSTRUCTIONS As String = "tblProcessInstructions"
Private Const TABLE_RECIPES As String = "tblRecipes"
Private Const TABLE_RECIPE_PROCESSES As String = "tblRecipeProcesses"
Private Const TABLE_RECIPE_CONNECTIONS As String = "tblRecipeConnections"

Public Function EnsureDesignsSchema(Optional ByVal targetWb As Workbook = Nothing, _
                                    Optional ByRef report As String = "") As Boolean
    On Error GoTo FailEnsure

    Dim wb As Workbook
    Set wb = targetWb
    If wb Is Nothing Then
        report = "Authoritative Designs workbook was not supplied."
        Exit Function
    End If
    If wb.IsAddin Then
        report = "Designs schema cannot be created inside an XLAM."
        Exit Function
    End If

    EnsureDesignsTable wb
    EnsureDesignLinesTable wb
    EnsureDesignEventsTable wb
    EnsureAppliedDesignEventsTable wb
    EnsureDesignLocksTable wb
    EnsureProcessesTable wb
    EnsureProcessRequirementsTable wb
    EnsureProcessAlternativesTable wb
    EnsureProcessOutputsTable wb
    EnsureProcessInstructionsTable wb
    EnsureRecipesTable wb
    EnsureRecipeProcessesTable wb
    EnsureRecipeConnectionsTable wb
    report = "OK"
    EnsureDesignsSchema = True
    Exit Function

FailEnsure:
    report = "EnsureDesignsSchema failed: " & Err.Description
End Function

Public Function ValidateDesignsSchema(ByVal targetWb As Workbook) As String
    If targetWb Is Nothing Then
        ValidateDesignsSchema = "Authoritative Designs workbook was not supplied."
        Exit Function
    End If

    ValidateDesignsSchema = ValidateRequiredTable(targetWb, TABLE_DESIGNS, _
        Array("DesignId", "DesignVersion", "DesignType", "DesignName", "Status", "CreatedAtUTC", "CreatedByUserId"))
    If ValidateDesignsSchema <> "" Then Exit Function
    ValidateDesignsSchema = ValidateRequiredTable(targetWb, TABLE_LINES, _
        Array("DesignId", "DesignVersion", "LineNo", "IOType", "ComponentSKU", "Qty", "UOM"))
    If ValidateDesignsSchema <> "" Then Exit Function
    ValidateDesignsSchema = ValidateRequiredTable(targetWb, TABLE_EVENTS, _
        Array("EventID", "AppliedSeq", "EventType", "WarehouseId", "DesignId", "DesignVersion", "PayloadJson"))
    If ValidateDesignsSchema <> "" Then Exit Function
    ValidateDesignsSchema = ValidateRequiredTable(targetWb, TABLE_APPLIED, _
        Array("EventID", "AppliedSeq", "AppliedAtUTC", "RunId", "SourceInbox", "Status"))
    If ValidateDesignsSchema <> "" Then Exit Function
    ValidateDesignsSchema = ValidateRequiredTable(targetWb, TABLE_LOCKS, _
        Array("LockName", "OwnerStationId", "OwnerUserId", "RunId", "AcquiredAtUTC", "ExpiresAtUTC", "HeartbeatAtUTC", "Status"))
    If ValidateDesignsSchema <> "" Then Exit Function
    ValidateDesignsSchema = ValidateRequiredTable(targetWb, TABLE_PROCESSES, _
        Array("ProcessId", "ProcessVersion", "ProcessName", "Status", "SourceEventID"))
    If ValidateDesignsSchema <> "" Then Exit Function
    ValidateDesignsSchema = ValidateRequiredTable(targetWb, TABLE_PROCESS_REQUIREMENTS, _
        Array("ProcessId", "ProcessVersion", "RequirementId", "RequirementName", "Qty", "Percent", "YieldBasis", "UOM"))
    If ValidateDesignsSchema <> "" Then Exit Function
    ValidateDesignsSchema = ValidateRequiredTable(targetWb, TABLE_PROCESS_ALTERNATIVES, _
        Array("ProcessId", "ProcessVersion", "RequirementId", "AlternativeOrdinal", "ITEM_CODE"))
    If ValidateDesignsSchema <> "" Then Exit Function
    ValidateDesignsSchema = ValidateRequiredTable(targetWb, TABLE_PROCESS_OUTPUTS, _
        Array("ProcessId", "ProcessVersion", "OutputId", "OutputName", "ITEM_CODE", "Qty", "Percent", "YieldBasis", "UOM"))
    If ValidateDesignsSchema <> "" Then Exit Function
    ValidateDesignsSchema = ValidateRequiredTable(targetWb, TABLE_PROCESS_INSTRUCTIONS, _
        Array("ProcessId", "ProcessVersion", "InstructionOrdinal", "Instruction"))
    If ValidateDesignsSchema <> "" Then Exit Function
    ValidateDesignsSchema = ValidateRequiredTable(targetWb, TABLE_RECIPES, _
        Array("RecipeId", "RecipeVersion", "RecipeName", "Status", "SourceEventID"))
    If ValidateDesignsSchema <> "" Then Exit Function
    ValidateDesignsSchema = ValidateRequiredTable(targetWb, TABLE_RECIPE_PROCESSES, _
        Array("RecipeId", "RecipeVersion", "ProcessNodeId", "ProcessId", "ProcessVersion", "ExecutionOrdinal"))
    If ValidateDesignsSchema <> "" Then Exit Function
    ValidateDesignsSchema = ValidateRequiredTable(targetWb, TABLE_RECIPE_CONNECTIONS, _
        Array("RecipeId", "RecipeVersion", "FromProcessNodeId", "FromOutputId", "ToProcessNodeId", "ToRequirementId", "Qty", "Percent", "UOM"))
End Function

Private Sub EnsureDesignsTable(ByVal wb As Workbook)
    EnsureTable wb, SHEET_DESIGNS, TABLE_DESIGNS, Array( _
        "DesignId", "DesignVersion", "DesignType", "DesignName", "Description", "Status", _
        "EffectiveFromUTC", "EffectiveToUTC", "CreatedAtUTC", "CreatedByUserId", _
        "ReleasedAtUTC", "ReleasedByUserId", "ObsoletedAtUTC", "ObsoletedByUserId", "SourceEventID")
End Sub

Private Sub EnsureDesignLinesTable(ByVal wb As Workbook)
    EnsureTable wb, SHEET_LINES, TABLE_LINES, Array( _
        "DesignId", "DesignVersion", "LineNo", "Process", "IOType", "ComponentSKU", _
        "ComponentDesignId", "ComponentDesignVersion", "Qty", "UOM", "Percent", "Instruction")
End Sub

Private Sub EnsureDesignEventsTable(ByVal wb As Workbook)
    EnsureTable wb, SHEET_EVENTS, TABLE_EVENTS, Array( _
        "EventID", "UndoOfEventId", "AppliedSeq", "EventType", "OccurredAtUTC", "AppliedAtUTC", _
        "WarehouseId", "StationId", "UserId", "DefinitionType", "DefinitionId", _
        "DefinitionVersion", "DesignId", "DesignVersion", "PayloadJson", "Note")
End Sub

Private Sub EnsureAppliedDesignEventsTable(ByVal wb As Workbook)
    EnsureTable wb, SHEET_APPLIED, TABLE_APPLIED, Array( _
        "EventID", "UndoOfEventId", "AppliedSeq", "AppliedAtUTC", "RunId", "SourceInbox", "Status")
End Sub

Private Sub EnsureDesignLocksTable(ByVal wb As Workbook)
    EnsureTable wb, SHEET_LOCKS, TABLE_LOCKS, Array( _
        "LockName", "OwnerStationId", "OwnerUserId", "RunId", "AcquiredAtUTC", _
        "ExpiresAtUTC", "HeartbeatAtUTC", "Status")
End Sub

Private Sub EnsureProcessesTable(ByVal wb As Workbook)
    EnsureTable wb, SHEET_PROCESSES, TABLE_PROCESSES, Array( _
        "ProcessId", "ProcessVersion", "ProcessName", "Description", "Status", _
        "CreatedAtUTC", "CreatedByUserId", "ReleasedAtUTC", "ReleasedByUserId", _
        "ObsoletedAtUTC", "ObsoletedByUserId", "SourceEventID")
End Sub

Private Sub EnsureProcessRequirementsTable(ByVal wb As Workbook)
    EnsureTable wb, SHEET_PROCESS_REQUIREMENTS, TABLE_PROCESS_REQUIREMENTS, Array( _
        "ProcessId", "ProcessVersion", "RequirementId", "RequirementName", _
        "Qty", "Percent", "YieldBasis", "UOM")
End Sub

Private Sub EnsureProcessAlternativesTable(ByVal wb As Workbook)
    EnsureTable wb, SHEET_PROCESS_ALTERNATIVES, TABLE_PROCESS_ALTERNATIVES, Array( _
        "ProcessId", "ProcessVersion", "RequirementId", "AlternativeOrdinal", "ITEM_CODE")
End Sub

Private Sub EnsureProcessOutputsTable(ByVal wb As Workbook)
    EnsureTable wb, SHEET_PROCESS_OUTPUTS, TABLE_PROCESS_OUTPUTS, Array( _
        "ProcessId", "ProcessVersion", "OutputId", "OutputName", "ITEM_CODE", _
        "ComponentDesignId", "ComponentDesignVersion", "Qty", "Percent", _
        "YieldBasis", "UOM")
End Sub

Private Sub EnsureProcessInstructionsTable(ByVal wb As Workbook)
    EnsureTable wb, SHEET_PROCESS_INSTRUCTIONS, TABLE_PROCESS_INSTRUCTIONS, Array( _
        "ProcessId", "ProcessVersion", "InstructionOrdinal", "Instruction")
End Sub

Private Sub EnsureRecipesTable(ByVal wb As Workbook)
    EnsureTable wb, SHEET_RECIPES, TABLE_RECIPES, Array( _
        "RecipeId", "RecipeVersion", "RecipeName", "Description", "Status", _
        "CreatedAtUTC", "CreatedByUserId", "ReleasedAtUTC", "ReleasedByUserId", _
        "ObsoletedAtUTC", "ObsoletedByUserId", "SourceEventID")
End Sub

Private Sub EnsureRecipeProcessesTable(ByVal wb As Workbook)
    EnsureTable wb, SHEET_RECIPE_PROCESSES, TABLE_RECIPE_PROCESSES, Array( _
        "RecipeId", "RecipeVersion", "ProcessNodeId", "ProcessId", _
        "ProcessVersion", "ExecutionOrdinal")
End Sub

Private Sub EnsureRecipeConnectionsTable(ByVal wb As Workbook)
    EnsureTable wb, SHEET_RECIPE_CONNECTIONS, TABLE_RECIPE_CONNECTIONS, Array( _
        "RecipeId", "RecipeVersion", "FromProcessNodeId", "FromOutputId", _
        "ToProcessNodeId", "ToRequirementId", "Qty", "Percent", "UOM")
End Sub

Private Sub EnsureTable(ByVal wb As Workbook, ByVal sheetName As String, _
                        ByVal tableName As String, ByVal headers As Variant)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim i As Long
    Dim nextColumn As Long
    Dim dataRange As Range

    Set ws = EnsureWorksheet(wb, sheetName)
    Set lo = FindTable(wb, tableName)
    If lo Is Nothing Then
        For i = LBound(headers) To UBound(headers)
            ws.Cells(1, i - LBound(headers) + 1).Value = CStr(headers(i))
        Next i
        Set dataRange = ws.Range(ws.Cells(1, 1), ws.Cells(2, UBound(headers) - LBound(headers) + 1))
        Set lo = ws.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
        lo.Name = tableName
        If Not lo.DataBodyRange Is Nothing Then lo.ListRows(1).Delete
    Else
        For i = LBound(headers) To UBound(headers)
            If ColumnIndex(lo, CStr(headers(i))) = 0 Then
                nextColumn = lo.ListColumns.Count + 1
                lo.ListColumns.Add nextColumn
                lo.ListColumns(nextColumn).Name = CStr(headers(i))
            End If
        Next i
    End If
    FormatDesignIdentityColumns lo
End Sub

Private Sub FormatDesignIdentityColumns(ByVal lo As ListObject)
    Dim columnName As String
    Dim currentFormat As Variant
    Dim lc As ListColumn

    If lo Is Nothing Then Exit Sub
    For Each lc In lo.ListColumns
        columnName = UCase$(Trim$(lc.Name))
        Select Case columnName
            Case "EVENTID", "UNDOOFEVENTID", "WAREHOUSEID", "STATIONID", "USERID", _
                 "DESIGNID", "DESIGNVERSION", "COMPONENTSKU", "COMPONENTDESIGNID", _
                 "COMPONENTDESIGNVERSION", "SOURCEEVENTID", "RUNID", "OWNERSTATIONID", _
                 "OWNERUSERID", "DEFINITIONID", "DEFINITIONVERSION", "PROCESSID", _
                 "PROCESSVERSION", "RECIPEID", "RECIPEVERSION", "PROCESSNODEID", _
                 "REQUIREMENTID", "OUTPUTID", "FROMPROCESSNODEID", "FROMOUTPUTID", _
                 "TOPROCESSNODEID", "TOREQUIREMENTID", "ITEM_CODE"
                currentFormat = lc.Range.NumberFormat
                If IsNull(currentFormat) Then
                    lc.Range.NumberFormat = "@"
                ElseIf StrComp(CStr(currentFormat), "@", vbBinaryCompare) <> 0 Then
                    lc.Range.NumberFormat = "@"
                End If
        End Select
    Next lc
End Sub

Private Function EnsureWorksheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureWorksheet = wb.Worksheets(sheetName)
    On Error GoTo 0
    If EnsureWorksheet Is Nothing Then
        Set EnsureWorksheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        EnsureWorksheet.Name = sheetName
    End If
End Function

Private Function FindTable(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        On Error Resume Next
        Set FindTable = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not FindTable Is Nothing Then Exit Function
    Next ws
End Function

Private Function ColumnIndex(ByVal lo As ListObject, ByVal columnName As String) As Long
    Dim i As Long
    If lo Is Nothing Then Exit Function
    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).Name, columnName, vbTextCompare) = 0 Then
            ColumnIndex = i
            Exit Function
        End If
    Next i
End Function

Private Function ValidateRequiredTable(ByVal wb As Workbook, ByVal tableName As String, _
                                       ByVal headers As Variant) As String
    Dim lo As ListObject
    Dim i As Long

    Set lo = FindTable(wb, tableName)
    If lo Is Nothing Then
        ValidateRequiredTable = tableName & " not found."
        Exit Function
    End If
    For i = LBound(headers) To UBound(headers)
        If ColumnIndex(lo, CStr(headers(i))) = 0 Then
            ValidateRequiredTable = tableName & " missing column " & CStr(headers(i)) & "."
            Exit Function
        End If
    Next i
End Function
