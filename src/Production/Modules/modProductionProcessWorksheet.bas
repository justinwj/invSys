Attribute VB_Name = "modProductionProcessWorksheet"
Option Explicit

Private Const EDITOR_SHEET As String = "invSys Process Editor"
Private Const TABLE_PREFIX As String = "invSys_Process_"
Private Const HEADER_ROW As Long = 6
Private Const FIRST_DATA_ROW As Long = 7

Private Const COL_RECORD_TYPE As Long = 1
Private Const COL_ID As Long = 2
Private Const COL_NAME As Long = 3
Private Const COL_ITEM_CODE As Long = 4
Private Const COL_QTY As Long = 5
Private Const COL_PERCENT As Long = 6
Private Const COL_BASIS_QTY As Long = 7
Private Const COL_UOM As Long = 8
Private Const COL_DESIGN_ID As Long = 9
Private Const COL_DESIGN_VERSION As Long = 10
Private Const COL_INSTRUCTION As Long = 11

Public Function SendProcessDraftToWorksheet(ByVal wb As Workbook, _
                                            ByVal processId As String, _
                                            ByVal processVersion As String, _
                                            ByVal processName As String, _
                                            ByVal description As String, _
                                            ByVal payloadJson As String, _
                                            ByRef tableName As String, _
                                            ByRef report As String) As Boolean
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim existingName As String
    Dim existingReport As String
    Dim rows As Collection
    Dim rowCount As Long
    Dim tableRange As Range
    Dim rowIndex As Long
    Dim record As Object

    On Error GoTo Failed
    If wb Is Nothing Then
        report = "The captured Production operator workbook is unavailable."
        Exit Function
    End If
    processId = UCase$(Trim$(processId))
    processVersion = Trim$(processVersion)
    processName = Trim$(processName)
    If Not mProduction.IsBase36Identifier(processId) Then
        report = "The Process needs a generated three-character Base-36 ID before worksheet editing."
        Exit Function
    End If
    If processVersion = "" Or Not IsNumeric(processVersion) Or CDbl(processVersion) <= 0 Then
        report = "The Process needs a positive generated version before worksheet editing."
        Exit Function
    End If
    If FindOutstandingProcessWorksheetTable(wb, existingName, existingReport) Then
        tableName = existingName
        report = "Retrieve or discard the outstanding Process worksheet table first: " & existingName & "."
        Exit Function
    End If

    Set rows = BuildWorksheetRows(payloadJson, report)
    If rows Is Nothing Then Exit Function
    AddWorksheetTemplateRows rows
    rowCount = rows.Count
    If rowCount = 0 Then
        report = "The Process worksheet could not create an editable row set."
        Exit Function
    End If

    Set ws = EnsureProcessEditorSheet(wb)
    ClearOwnedEditorSurface ws
    ws.Range("A1").Value2 = "invSys Process worksheet"
    ws.Range("A2").Value2 = "Process Name"
    ws.Range("B2").Value2 = processName
    ws.Range("D2").Value2 = "Process ID"
    ws.Range("E2").Value2 = processId
    ws.Range("G2").Value2 = "Version"
    ws.Range("H2").Value2 = processVersion
    ws.Range("A3").Value2 = "Description"
    ws.Range("B3").Value2 = description
    ws.Range("A4").Value2 = _
        "Enter INPUT quantities in one compatible UOM. Batch basis and percentages calculate automatically."

    WriteWorksheetHeaders ws
    Set tableRange = ws.Range(ws.Cells(HEADER_ROW, 1), _
                              ws.Cells(HEADER_ROW + rowCount, COL_INSTRUCTION))
    Set lo = ws.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
    tableName = BuildUniqueProcessTableName(wb, processId)
    lo.Name = tableName

    rowIndex = 1
    For Each record In rows
        WriteWorksheetRecord lo, rowIndex, record
        rowIndex = rowIndex + 1
    Next record
    ApplyInputFormulas lo
    FormatProcessWorksheet ws, lo
    wb.Save
    ws.Visible = xlSheetVisible
    ws.Activate
    lo.Range.Cells(1, 1).Select
    report = "Process draft sent to " & tableName & ". Edit the table, then select Retrieve Process from Sheet."
    SendProcessDraftToWorksheet = True
    Exit Function

Failed:
    report = "Process worksheet creation failed: " & Err.Description
End Function

Public Function ReadProcessDraftFromWorksheet(ByVal wb As Workbook, _
                                              ByVal tableName As String, _
                                              ByRef processId As String, _
                                              ByRef processVersion As String, _
                                              ByRef processName As String, _
                                              ByRef description As String, _
                                              ByRef payloadJson As String, _
                                              ByRef report As String) As Boolean
    Dim lo As ListObject
    Dim records As New Collection
    Dim record As Object
    Dim requirementIds As Object
    Dim outputIds As Object
    Dim inputUoms As Object
    Dim inputRowCount As Long
    Dim derivedInputCount As Long
    Dim outputRowCount As Long
    Dim instructionOrdinal As Long
    Dim inputPercentTotal As Double
    Dim rowIndex As Long
    Dim recordType As String
    Dim rowId As String
    Dim rowName As String
    Dim itemCode As String
    Dim uom As String
    Dim instructionText As String
    Dim qty As Double
    Dim percentValue As Double
    Dim basisQty As Double
    Dim hasQty As Boolean
    Dim hasPercent As Boolean
    Dim expectedProcessId As String

    On Error GoTo Failed
    If wb Is Nothing Then
        report = "The captured Production operator workbook is unavailable."
        Exit Function
    End If
    Set lo = FindProcessTableByName(wb, tableName)
    If lo Is Nothing Then
        report = "The bound Process worksheet table is missing: " & tableName & "."
        Exit Function
    End If
    expectedProcessId = ProcessIdFromTableName(lo.Name)
    If Not mProduction.IsBase36Identifier(expectedProcessId) Then
        report = "The Process worksheet table identity is invalid."
        Exit Function
    End If
    processId = expectedProcessId
    processVersion = Trim$(CellText(lo.Parent.Range("H2").Value2))
    processName = Trim$(CellText(lo.Parent.Range("B2").Value2))
    description = Trim$(CellText(lo.Parent.Range("B3").Value2))
    If processName = "" Then
        report = "Process Name is required on the worksheet."
        Exit Function
    End If
    If processVersion = "" Or Not IsNumeric(processVersion) Or CDbl(processVersion) <= 0 Then
        report = "The Process worksheet version must be a positive number."
        Exit Function
    End If

    Set requirementIds = CreateObject("Scripting.Dictionary")
    requirementIds.CompareMode = vbTextCompare
    Set outputIds = CreateObject("Scripting.Dictionary")
    outputIds.CompareMode = vbTextCompare
    Set inputUoms = CreateObject("Scripting.Dictionary")
    inputUoms.CompareMode = vbTextCompare

    Set record = NewWorksheetRecord("PROCESS")
    record("ProcessName") = processName
    record("Description") = description
    records.Add record

    If lo.DataBodyRange Is Nothing Then
        report = "The Process worksheet table contains no rows."
        Exit Function
    End If
    For rowIndex = 1 To lo.ListRows.Count
        recordType = UCase$(Trim$(WorksheetValue(lo, rowIndex, COL_RECORD_TYPE)))
        If Not WorksheetRowHasBusinessData(lo, rowIndex) Then GoTo ContinueRow
        rowId = UCase$(Trim$(WorksheetValue(lo, rowIndex, COL_ID)))
        rowName = Trim$(WorksheetValue(lo, rowIndex, COL_NAME))
        itemCode = Trim$(WorksheetValue(lo, rowIndex, COL_ITEM_CODE))
        uom = UCase$(Trim$(WorksheetValue(lo, rowIndex, COL_UOM)))
        instructionText = Trim$(WorksheetValue(lo, rowIndex, COL_INSTRUCTION))
        hasQty = TryPositiveWorksheetNumber(WorksheetValue(lo, rowIndex, COL_QTY), qty)
        hasPercent = TryPositiveWorksheetNumber(WorksheetValue(lo, rowIndex, COL_PERCENT), percentValue)
        Call TryPositiveWorksheetNumber(WorksheetValue(lo, rowIndex, COL_BASIS_QTY), basisQty)

        Select Case recordType
            Case "INPUT", "REQUIREMENT"
                rowId = ResolveWorksheetRowId(rowId, requirementIds)
                If rowName = "" Or uom = "" Then
                    report = "Each INPUT row needs a name and UOM."
                    Exit Function
                End If
                If Not hasQty And Not hasPercent Then
                    report = "Each INPUT row needs a positive quantity or percentage."
                    Exit Function
                End If
                If hasPercent And basisQty <= 0 Then
                    report = "Each percentage INPUT needs a positive Batch basis quantity."
                    Exit Function
                End If
                If requirementIds.Exists(rowId) Then
                    report = "Duplicate INPUT ID " & rowId & "."
                    Exit Function
                End If
                requirementIds.Add rowId, True
                inputUoms(uom) = True
                inputRowCount = inputRowCount + 1
                If hasQty Then
                    derivedInputCount = derivedInputCount + 1
                    inputPercentTotal = inputPercentTotal + percentValue
                End If
                Set record = NewWorksheetRecord("REQUIREMENT")
                record("RequirementId") = rowId
                record("RequirementName") = rowName
                If hasQty Then record("Qty") = qty
                If hasPercent Then record("Percent") = percentValue
                If basisQty > 0 Then record("YieldBasis") = basisQty
                record("UOM") = uom
                records.Add record
            Case "OUTPUT"
                rowId = ResolveWorksheetRowId(rowId, outputIds)
                If rowName = "" Or itemCode = "" Or uom = "" Then
                    report = "Each OUTPUT row needs a name, Item Code, and UOM."
                    Exit Function
                End If
                If Not hasQty And Not hasPercent Then
                    report = "Each OUTPUT row needs a positive quantity or percentage."
                    Exit Function
                End If
                If hasPercent And basisQty <= 0 Then
                    report = "Each percentage OUTPUT needs a positive Yield basis quantity."
                    Exit Function
                End If
                If outputIds.Exists(rowId) Then
                    report = "Duplicate OUTPUT ID " & rowId & "."
                    Exit Function
                End If
                outputIds.Add rowId, True
                outputRowCount = outputRowCount + 1
                Set record = NewWorksheetRecord("OUTPUT")
                record("OutputId") = rowId
                record("OutputName") = rowName
                record("ITEM_CODE") = itemCode
                record("ComponentDesignId") = Trim$(WorksheetValue(lo, rowIndex, COL_DESIGN_ID))
                record("ComponentDesignVersion") = Trim$(WorksheetValue(lo, rowIndex, COL_DESIGN_VERSION))
                If hasQty Then record("Qty") = qty
                If hasPercent Then record("Percent") = percentValue
                If basisQty > 0 Then record("YieldBasis") = basisQty
                record("UOM") = uom
                records.Add record
            Case "INSTRUCTION"
                If instructionText = "" Then instructionText = rowName
                If instructionText = "" Then
                    report = "Each INSTRUCTION row needs instruction text."
                    Exit Function
                End If
                instructionOrdinal = instructionOrdinal + 1
                Set record = NewWorksheetRecord("INSTRUCTION")
                record("InstructionOrdinal") = instructionOrdinal
                record("Instruction") = instructionText
                records.Add record
            Case "ALTERNATIVE"
                If rowId = "" Or itemCode = "" Then
                    report = "Each ALTERNATIVE row needs its Requirement ID and Item Code."
                    Exit Function
                End If
                Set record = NewWorksheetRecord("ALTERNATIVE")
                record("RequirementId") = rowId
                record("ITEM_CODE") = itemCode
                records.Add record
            Case Else
                report = "Unknown Process worksheet row type: " & recordType & "."
                Exit Function
        End Select
ContinueRow:
    Next rowIndex

    If outputRowCount = 0 Then
        report = "Every Process must declare at least one OUTPUT."
        Exit Function
    End If
    If inputUoms.Count > 1 Then
        report = "INPUT percentage formulas require one compatible UOM. Add an explicit conversion Process before retrieval."
        Exit Function
    End If
    If derivedInputCount > 0 And Abs(inputPercentTotal - 100#) > 0.05 Then
        report = "INPUT formula percentages must total 100.0%; current total is " & Format$(inputPercentTotal, "0.0") & "%."
        Exit Function
    End If

    payloadJson = modProductionJson.BuildJsonArray(records)
    report = "Process worksheet validated: " & CStr(inputRowCount) & _
        " input(s), " & CStr(outputRowCount) & " output(s), and " & _
        CStr(instructionOrdinal) & " instruction(s)."
    ReadProcessDraftFromWorksheet = True
    Exit Function

Failed:
    report = "Process worksheet retrieval failed: " & Err.Description
End Function

Public Function DeleteProcessWorksheetTable(ByVal wb As Workbook, _
                                            ByVal tableName As String, _
                                            ByRef report As String) As Boolean
    Dim lo As ListObject
    Dim ws As Worksheet

    On Error GoTo Failed
    Set lo = FindProcessTableByName(wb, tableName)
    If lo Is Nothing Then
        report = "The Process worksheet table is already absent."
        DeleteProcessWorksheetTable = True
        Exit Function
    End If
    Set ws = lo.Parent
    lo.Delete
    ClearOwnedEditorSurface ws
    wb.Save
    report = "Retrieved Process and removed temporary table " & tableName & "."
    DeleteProcessWorksheetTable = True
    Exit Function
Failed:
    report = "The Process was retrieved, but its temporary table could not be removed: " & Err.Description
End Function

Public Function FindOutstandingProcessWorksheetTable(ByVal wb As Workbook, _
                                                     ByRef tableName As String, _
                                                     ByRef report As String) As Boolean
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim foundCount As Long

    tableName = ""
    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If Left$(lo.Name, Len(TABLE_PREFIX)) = TABLE_PREFIX Then
                foundCount = foundCount + 1
                tableName = lo.Name
            End If
        Next lo
    Next ws
    If foundCount = 1 Then
        report = "Outstanding Process worksheet table found: " & tableName & "."
        FindOutstandingProcessWorksheetTable = True
    ElseIf foundCount > 1 Then
        report = "More than one invSys Process worksheet table exists. Remove extras before retrieval."
        tableName = ""
    Else
        report = "No outstanding Process worksheet table."
    End If
End Function

Public Function PopulateFormulationExampleForTest(ByVal wb As Workbook, _
                                                  ByVal tableName As String, _
                                                  ByVal mixedUom As Boolean, _
                                                  ByRef report As String) As Boolean
    Dim lo As ListObject
    Dim names As Variant
    Dim quantities As Variant
    Dim rowIndex As Long

    On Error GoTo Failed
    Set lo = FindProcessTableByName(wb, tableName)
    If lo Is Nothing Then
        report = "Test Process worksheet table is missing."
        Exit Function
    End If
    names = Array("Sugar", "Flour", "Baking Powder", "Filtered Water")
    quantities = Array(100#, 200#, 11.2, 300#)
    For rowIndex = 1 To 4
        lo.DataBodyRange.Cells(rowIndex, COL_RECORD_TYPE).Value2 = "INPUT"
        lo.DataBodyRange.Cells(rowIndex, COL_NAME).Value2 = names(rowIndex - 1)
        lo.DataBodyRange.Cells(rowIndex, COL_QTY).Value2 = quantities(rowIndex - 1)
        lo.DataBodyRange.Cells(rowIndex, COL_UOM).Value2 = IIf(mixedUom And rowIndex = 2, "KG", "LB")
    Next rowIndex
    lo.DataBodyRange.Cells(7, COL_RECORD_TYPE).Value2 = "OUTPUT"
    lo.DataBodyRange.Cells(7, COL_NAME).Value2 = "Finished Formula"
    lo.DataBodyRange.Cells(7, COL_ITEM_CODE).Value2 = "DEMO-FINISHED-FORMULA"
    lo.DataBodyRange.Cells(7, COL_QTY).Value2 = 611.2
    lo.DataBodyRange.Cells(7, COL_UOM).Value2 = "LB"
    ApplyInputFormulas lo
    Application.Calculate
    report = "Example populated."
    PopulateFormulationExampleForTest = True
    Exit Function
Failed:
    report = "Example population failed: " & Err.Description
End Function

Public Function ReadFormulaEvidenceForTest(ByVal wb As Workbook, _
                                           ByVal tableName As String) As String
    Dim lo As ListObject
    Dim totalPercent As Double
    Dim rowIndex As Long

    Set lo = FindProcessTableByName(wb, tableName)
    If lo Is Nothing Then
        ReadFormulaEvidenceForTest = "FAIL|TableMissing"
        Exit Function
    End If
    For rowIndex = 1 To 4
        totalPercent = totalPercent + CDbl(lo.DataBodyRange.Cells(rowIndex, COL_PERCENT).Value2)
    Next rowIndex
    ReadFormulaEvidenceForTest = "OK|Basis=" & _
        Format$(CDbl(lo.DataBodyRange.Cells(1, COL_BASIS_QTY).Value2), "0.0") & _
        "|Sugar=" & Format$(CDbl(lo.DataBodyRange.Cells(1, COL_PERCENT).Value2), "0.0") & _
        "|Flour=" & Format$(CDbl(lo.DataBodyRange.Cells(2, COL_PERCENT).Value2), "0.0") & _
        "|BakingPowder=" & Format$(CDbl(lo.DataBodyRange.Cells(3, COL_PERCENT).Value2), "0.0") & _
        "|Water=" & Format$(CDbl(lo.DataBodyRange.Cells(4, COL_PERCENT).Value2), "0.0") & _
        "|Total=" & Format$(totalPercent, "0.0")
End Function

Private Function BuildWorksheetRows(ByVal payloadJson As String, _
                                    ByRef report As String) As Collection
    Dim sourceRecords As Collection
    Dim rows As New Collection
    Dim record As Object
    Dim rowRecord As Object
    Dim parseReport As String
    Dim recordType As String

    Set sourceRecords = modProductionReusableDesigns.ParseReusableDefinitionRecords(payloadJson, parseReport)
    If sourceRecords Is Nothing Then
        report = "The Process draft could not be serialized for worksheet editing: " & parseReport
        Exit Function
    End If
    For Each record In sourceRecords
        recordType = UCase$(modProductionReusableDesigns.ReusableRecordText(record, "RecordType"))
        Select Case recordType
            Case "REQUIREMENT"
                Set rowRecord = NewWorksheetRecord("INPUT")
                rowRecord("Id") = modProductionReusableDesigns.ReusableRecordText(record, "RequirementId")
                rowRecord("Name") = modProductionReusableDesigns.ReusableRecordText(record, "RequirementName")
                rowRecord("Qty") = modProductionReusableDesigns.ReusableRecordValue(record, "Qty")
                rowRecord("Percent") = modProductionReusableDesigns.ReusableRecordValue(record, "Percent")
                rowRecord("BasisQty") = modProductionReusableDesigns.ReusableRecordValue(record, "YieldBasis")
                rowRecord("UOM") = modProductionReusableDesigns.ReusableRecordText(record, "UOM")
                rows.Add rowRecord
            Case "ALTERNATIVE"
                Set rowRecord = NewWorksheetRecord("ALTERNATIVE")
                rowRecord("Id") = modProductionReusableDesigns.ReusableRecordText(record, "RequirementId")
                rowRecord("ItemCode") = modProductionReusableDesigns.ReusableRecordText(record, "ITEM_CODE")
                rows.Add rowRecord
            Case "OUTPUT"
                Set rowRecord = NewWorksheetRecord("OUTPUT")
                rowRecord("Id") = modProductionReusableDesigns.ReusableRecordText(record, "OutputId")
                rowRecord("Name") = modProductionReusableDesigns.ReusableRecordText(record, "OutputName")
                rowRecord("ItemCode") = modProductionReusableDesigns.ReusableRecordText(record, "ITEM_CODE")
                rowRecord("Qty") = modProductionReusableDesigns.ReusableRecordValue(record, "Qty")
                rowRecord("Percent") = modProductionReusableDesigns.ReusableRecordValue(record, "Percent")
                rowRecord("BasisQty") = modProductionReusableDesigns.ReusableRecordValue(record, "YieldBasis")
                rowRecord("UOM") = modProductionReusableDesigns.ReusableRecordText(record, "UOM")
                rowRecord("DesignId") = modProductionReusableDesigns.ReusableRecordText(record, "ComponentDesignId")
                rowRecord("DesignVersion") = modProductionReusableDesigns.ReusableRecordText(record, "ComponentDesignVersion")
                rows.Add rowRecord
            Case "INSTRUCTION"
                Set rowRecord = NewWorksheetRecord("INSTRUCTION")
                rowRecord("Instruction") = modProductionReusableDesigns.ReusableRecordText(record, "Instruction")
                rows.Add rowRecord
        End Select
    Next record
    Set BuildWorksheetRows = rows
End Function

Private Sub AddWorksheetTemplateRows(ByVal rows As Collection)
    Dim inputCount As Long
    Dim outputCount As Long
    Dim instructionCount As Long
    Dim usedInputs As Object
    Dim usedOutputs As Object
    Dim record As Object
    Dim rowRecord As Object
    Dim recordType As String
    Dim addCount As Long
    Dim index As Long

    Set usedInputs = CreateObject("Scripting.Dictionary")
    usedInputs.CompareMode = vbTextCompare
    Set usedOutputs = CreateObject("Scripting.Dictionary")
    usedOutputs.CompareMode = vbTextCompare
    For Each record In rows
        recordType = UCase$(CellText(record("RecordType")))
        Select Case recordType
            Case "INPUT"
                inputCount = inputCount + 1
                If Trim$(CellText(record("Id"))) <> "" Then usedInputs(UCase$(Trim$(CellText(record("Id"))))) = True
            Case "OUTPUT"
                outputCount = outputCount + 1
                If Trim$(CellText(record("Id"))) <> "" Then usedOutputs(UCase$(Trim$(CellText(record("Id"))))) = True
            Case "INSTRUCTION": instructionCount = instructionCount + 1
        End Select
    Next record

    addCount = IIf(inputCount = 0, 6, 2)
    For index = 1 To addCount
        Set rowRecord = NewWorksheetRecord("INPUT")
        rowRecord("Id") = NextIdFromDictionary(usedInputs)
        usedInputs(rowRecord("Id")) = True
        rows.Add rowRecord
    Next index
    addCount = IIf(outputCount = 0, 2, 1)
    For index = 1 To addCount
        Set rowRecord = NewWorksheetRecord("OUTPUT")
        rowRecord("Id") = NextIdFromDictionary(usedOutputs)
        usedOutputs(rowRecord("Id")) = True
        rows.Add rowRecord
    Next index
    Set rowRecord = NewWorksheetRecord("INSTRUCTION")
    rows.Add rowRecord
End Sub

Private Function NewWorksheetRecord(ByVal recordType As String) As Object
    Dim record As Object

    Set record = CreateObject("Scripting.Dictionary")
    record.CompareMode = vbTextCompare
    record("RecordType") = recordType
    record("Id") = ""
    record("Name") = ""
    record("ItemCode") = ""
    record("Qty") = ""
    record("Percent") = ""
    record("BasisQty") = ""
    record("UOM") = ""
    record("DesignId") = ""
    record("DesignVersion") = ""
    record("Instruction") = ""
    Set NewWorksheetRecord = record
End Function

Private Sub WriteWorksheetRecord(ByVal lo As ListObject, ByVal rowIndex As Long, _
                                 ByVal record As Object)
    With lo.DataBodyRange
        .Cells(rowIndex, COL_RECORD_TYPE).Value2 = record("RecordType")
        .Cells(rowIndex, COL_ID).Value2 = record("Id")
        .Cells(rowIndex, COL_NAME).Value2 = record("Name")
        .Cells(rowIndex, COL_ITEM_CODE).Value2 = record("ItemCode")
        .Cells(rowIndex, COL_QTY).Value2 = record("Qty")
        .Cells(rowIndex, COL_PERCENT).Value2 = record("Percent")
        .Cells(rowIndex, COL_BASIS_QTY).Value2 = record("BasisQty")
        .Cells(rowIndex, COL_UOM).Value2 = record("UOM")
        .Cells(rowIndex, COL_DESIGN_ID).Value2 = record("DesignId")
        .Cells(rowIndex, COL_DESIGN_VERSION).Value2 = record("DesignVersion")
        .Cells(rowIndex, COL_INSTRUCTION).Value2 = record("Instruction")
    End With
End Sub

Private Sub ApplyInputFormulas(ByVal lo As ListObject)
    Dim rowIndex As Long
    Dim recordType As String
    Dim qty As Double
    Dim priorAutoFill As Boolean

    On Error GoTo CleanExit
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    priorAutoFill = Application.AutoCorrect.AutoFillFormulasInLists
    Application.AutoCorrect.AutoFillFormulasInLists = False
    For rowIndex = 1 To lo.ListRows.Count
        recordType = UCase$(Trim$(WorksheetValue(lo, rowIndex, COL_RECORD_TYPE)))
        If recordType = "INPUT" Or recordType = "REQUIREMENT" Then
            If TryPositiveWorksheetNumber(WorksheetValue(lo, rowIndex, COL_QTY), qty) Then
                lo.DataBodyRange.Cells(rowIndex, COL_BASIS_QTY).Formula = _
                    "=IFERROR(SUMIFS([Qty],[Record Type],""INPUT"",[UOM],[@UOM]),"""")"
                lo.DataBodyRange.Cells(rowIndex, COL_PERCENT).Formula = _
                    "=IFERROR([@Qty]/[@[Basis Qty]]*100,"""")"
            End If
        End If
    Next rowIndex
CleanExit:
    On Error Resume Next
    Application.AutoCorrect.AutoFillFormulasInLists = priorAutoFill
    On Error GoTo 0
End Sub

Private Sub WriteWorksheetHeaders(ByVal ws As Worksheet)
    Dim headers As Variant
    Dim index As Long

    headers = Array("Record Type", "ID", "Name", "Item Code", "Qty", "Percent", _
                    "Basis Qty", "UOM", "Design ID", "Design Version", "Instruction")
    For index = LBound(headers) To UBound(headers)
        ws.Cells(HEADER_ROW, index + 1).Value2 = headers(index)
    Next index
End Sub

Private Sub FormatProcessWorksheet(ByVal ws As Worksheet, ByVal lo As ListObject)
    ws.Columns("A").ColumnWidth = 14
    ws.Columns("B").ColumnWidth = 10
    ws.Columns("C").ColumnWidth = 24
    ws.Columns("D").ColumnWidth = 24
    ws.Columns("E:G").ColumnWidth = 13
    ws.Columns("H").ColumnWidth = 10
    ws.Columns("I:J").ColumnWidth = 16
    ws.Columns("K").ColumnWidth = 48
    ws.Range("A1:H1").Font.Bold = True
    ws.Range("A2:H3").WrapText = True
    lo.ListColumns("Qty").DataBodyRange.NumberFormat = "0.########"
    lo.ListColumns("Percent").DataBodyRange.NumberFormat = "0.0\%"
    lo.ListColumns("Basis Qty").DataBodyRange.NumberFormat = "0.########"
    lo.Range.VerticalAlignment = xlTop
End Sub

Private Function EnsureProcessEditorSheet(ByVal wb As Workbook) As Worksheet
    On Error Resume Next
    Set EnsureProcessEditorSheet = wb.Worksheets(EDITOR_SHEET)
    On Error GoTo 0
    If EnsureProcessEditorSheet Is Nothing Then
        Set EnsureProcessEditorSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        EnsureProcessEditorSheet.Name = EDITOR_SHEET
    End If
End Function

Private Sub ClearOwnedEditorSurface(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    ws.Range("A1:K500").Clear
End Sub

Private Function BuildUniqueProcessTableName(ByVal wb As Workbook, _
                                             ByVal processId As String) As String
    Dim token As String
    Dim candidate As String
    Dim existing As ListObject

    token = UCase$(Replace$(mProduction.CreateProductionGuid(), "-", ""))
    token = Left$(token, 8)
    candidate = TABLE_PREFIX & processId & "_" & token
    Set existing = FindProcessTableByName(wb, candidate)
    Do While Not existing Is Nothing
        token = UCase$(Replace$(mProduction.CreateProductionGuid(), "-", ""))
        candidate = TABLE_PREFIX & processId & "_" & Left$(token, 8)
        Set existing = FindProcessTableByName(wb, candidate)
    Loop
    BuildUniqueProcessTableName = candidate
End Function

Private Function FindProcessTableByName(ByVal wb As Workbook, _
                                        ByVal tableName As String) As ListObject
    Dim ws As Worksheet

    If wb Is Nothing Or Trim$(tableName) = "" Then Exit Function
    For Each ws In wb.Worksheets
        On Error Resume Next
        Set FindProcessTableByName = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not FindProcessTableByName Is Nothing Then Exit Function
    Next ws
End Function

Private Function ProcessIdFromTableName(ByVal tableName As String) As String
    Dim suffix As String

    If Left$(tableName, Len(TABLE_PREFIX)) <> TABLE_PREFIX Then Exit Function
    suffix = Mid$(tableName, Len(TABLE_PREFIX) + 1)
    ProcessIdFromTableName = UCase$(Left$(suffix, 3))
End Function

Private Function WorksheetValue(ByVal lo As ListObject, ByVal rowIndex As Long, _
                                ByVal columnIndex As Long) As String
    WorksheetValue = CellText(lo.DataBodyRange.Cells(rowIndex, columnIndex).Value2)
End Function

Private Function WorksheetRowHasBusinessData(ByVal lo As ListObject, _
                                             ByVal rowIndex As Long) As Boolean
    WorksheetRowHasBusinessData = _
        (Trim$(WorksheetValue(lo, rowIndex, COL_NAME)) <> "") Or _
        (Trim$(WorksheetValue(lo, rowIndex, COL_ITEM_CODE)) <> "") Or _
        (Trim$(WorksheetValue(lo, rowIndex, COL_QTY)) <> "") Or _
        (Trim$(WorksheetValue(lo, rowIndex, COL_UOM)) <> "") Or _
        (Trim$(WorksheetValue(lo, rowIndex, COL_DESIGN_ID)) <> "") Or _
        (Trim$(WorksheetValue(lo, rowIndex, COL_DESIGN_VERSION)) <> "") Or _
        (Trim$(WorksheetValue(lo, rowIndex, COL_INSTRUCTION)) <> "")
End Function

Private Function TryPositiveWorksheetNumber(ByVal valueText As String, _
                                            ByRef numberValue As Double) As Boolean
    numberValue = 0
    If Trim$(valueText) = "" Or Not IsNumeric(valueText) Then Exit Function
    numberValue = CDbl(valueText)
    TryPositiveWorksheetNumber = (numberValue > 0)
End Function

Private Function ResolveWorksheetRowId(ByVal proposedId As String, _
                                       ByVal used As Object) As String
    proposedId = UCase$(Trim$(proposedId))
    If mProduction.IsBase36Identifier(proposedId) And Not used.Exists(proposedId) Then
        ResolveWorksheetRowId = proposedId
    Else
        ResolveWorksheetRowId = NextIdFromDictionary(used)
    End If
End Function

Private Function NextIdFromDictionary(ByVal used As Object) As String
    Dim keys As Variant

    If used Is Nothing Or used.Count = 0 Then
        keys = Array("")
    Else
        keys = used.Keys
    End If
    NextIdFromDictionary = mProduction.NextBase36Identifier(keys)
End Function

Private Function CellText(ByVal valueIn As Variant) As String
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Then Exit Function
    CellText = CStr(valueIn)
End Function
