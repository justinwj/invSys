Attribute VB_Name = "modProductionReusableRun"
Option Explicit

Private Const EVENT_TYPE_PROD_CONSUME As String = "PROD_CONSUME"
Private Const EVENT_TYPE_PROD_COMPLETE As String = "PROD_COMPLETE"
Private Const QTY_TOLERANCE As Double = 0.0000001

Private mLoaded As Boolean
Private mCheckedIn As Boolean
Private mCompleted As Boolean
Private mRecipeId As String
Private mRecipeVersion As String
Private mRecipeName As String
Private mScalePercent As Double
Private mBatchNumber As Long
Private mNodes As Collection
Private mRequirements As Collection
Private mAlternatives As Collection
Private mOutputs As Collection
Private mConnections As Collection
Private mAllocations As Object
Private mOutputKeys As Object
Private mActualOutputQty As Object
Private mLastOutputQty As Object
Private mOutputHistory As Collection
Private mLastSummary As String

Public Sub ClearReusableRun()
    mLoaded = False
    mCheckedIn = False
    mCompleted = False
    mRecipeId = ""
    mRecipeVersion = ""
    mRecipeName = ""
    mScalePercent = 100#
    mBatchNumber = 1
    Set mNodes = New Collection
    Set mRequirements = New Collection
    Set mAlternatives = New Collection
    Set mOutputs = New Collection
    Set mConnections = New Collection
    Set mAllocations = NewTextDictionary()
    Set mOutputKeys = NewTextDictionary()
    Set mActualOutputQty = NewTextDictionary()
    Set mLastOutputQty = NewTextDictionary()
    Set mOutputHistory = New Collection
    mLastSummary = ""
End Sub

Public Function LoadReleasedReusableRecipe(ByVal recipeId As String, _
                                           ByVal recipeVersion As String, _
                                           ByVal scalePercent As Double, _
                                           Optional ByRef report As String = "") As Boolean
    On Error GoTo Failed

    Dim validation As String
    Dim jsonText As String
    Dim parseReport As String
    Dim recipeRecords As Collection
    Dim rawRecord As Variant
    Dim record As Object
    Dim node As Object

    ClearReusableRun
    recipeId = Trim$(recipeId)
    recipeVersion = Trim$(recipeVersion)
    If recipeId = "" Or recipeVersion = "" Then
        report = "Recipe ID and version are required."
        Exit Function
    End If
    If scalePercent < 0.001 Or scalePercent > 1000# Then
        report = "Batch scale must be from 0.001% through 1000%."
        Exit Function
    End If

    validation = modOperationsPrimitiveBridge.ValidateReleasedRecipe(recipeId, recipeVersion)
    If Left$(validation, 2) <> "1" & vbTab Then
        report = "Released Recipe validation failed: " & Replace$(validation, vbTab, " ")
        Exit Function
    End If
    jsonText = modOperationsPrimitiveBridge.GetRecipeGraph(recipeId, recipeVersion)
    If jsonText = "" Then
        report = "Released Recipe graph could not be read."
        Exit Function
    End If
    Set recipeRecords = modProductionReusableDesigns.ParseReusableDefinitionRecords(jsonText, parseReport)
    If recipeRecords Is Nothing Then
        report = parseReport
        Exit Function
    End If

    mRecipeId = recipeId
    mRecipeVersion = recipeVersion
    mScalePercent = scalePercent
    For Each rawRecord In recipeRecords
        Set record = rawRecord
        Select Case UCase$(modProductionReusableDesigns.ReusableRecordText(record, "RecordType"))
            Case "RECIPE"
                mRecipeName = modProductionReusableDesigns.ReusableRecordText(record, "RecipeName")
            Case "PROCESS_NODE"
                Set node = CloneRunRecord(record)
                AddNodeInExecutionOrder node
            Case "CONNECTION"
                mConnections.Add CloneRunRecord(record)
        End Select
    Next rawRecord
    If mNodes.Count = 0 Then
        report = "Released Recipe contains no Process nodes."
        Exit Function
    End If
    If Not LoadNodeProcessDefinitions(report) Then Exit Function
    If Not ValidateLoadedRunGraph(report) Then Exit Function

    mLoaded = True
    report = "Loaded released Recipe " & recipeId & " version " & recipeVersion & _
             " with " & CStr(mNodes.Count) & " Process(es) at " & _
             FormatRunNumberLocal(scalePercent) & "% scale."
    LoadReleasedReusableRecipe = True
    Exit Function

Failed:
    report = "Reusable Recipe load failed: " & Err.Description
End Function

Public Function ApplyReusableRunScale(ByVal scalePercent As Double, _
                                      Optional ByRef report As String = "") As Boolean
    If Not mLoaded Then
        report = "Load a released Recipe before applying Batch scale."
        Exit Function
    End If
    If scalePercent < 0.001 Or scalePercent > 1000# Then
        report = "Batch scale must be from 0.001% through 1000%."
        Exit Function
    End If
    mScalePercent = scalePercent
    Set mAllocations = NewTextDictionary()
    Set mOutputKeys = NewTextDictionary()
    Set mActualOutputQty = NewTextDictionary()
    mCheckedIn = False
    mCompleted = False
    mLastSummary = ""
    report = "Batch scale applied: " & FormatRunNumberLocal(scalePercent) & _
             "%. Exact-key allocations were cleared and recalculated requirements are ready."
    ApplyReusableRunScale = True
End Function

Public Function ReusableRunLoaderRows() As Variant
    Dim result() As Variant
    Dim totalRows As Long
    Dim rowIndex As Long
    Dim rawRecord As Variant
    Dim record As Object

    If Not mLoaded Then Exit Function
    totalRows = mRequirements.Count + mOutputs.Count
    If totalRows = 0 Then Exit Function
    ReDim result(1 To totalRows, 1 To 8)
    For Each rawRecord In mRequirements
        Set record = rawRecord
        rowIndex = rowIndex + 1
        result(rowIndex, 1) = RunRecordText(record, "ProcessName")
        result(rowIndex, 2) = RunRecordText(record, "ProcessNodeId")
        result(rowIndex, 3) = "INPUT"
        result(rowIndex, 4) = RunRecordText(record, "RequirementName")
        result(rowIndex, 5) = RunRecordNumber(record, "Percent")
        result(rowIndex, 6) = RunRecordText(record, "UOM")
        result(rowIndex, 7) = ScaledRecordQty(record)
        result(rowIndex, 8) = RunRecordText(record, "RequirementId")
    Next rawRecord
    For Each rawRecord In mOutputs
        Set record = rawRecord
        rowIndex = rowIndex + 1
        result(rowIndex, 1) = RunRecordText(record, "ProcessName")
        result(rowIndex, 2) = RunRecordText(record, "ProcessNodeId")
        result(rowIndex, 3) = "OUTPUT"
        result(rowIndex, 4) = RunRecordText(record, "OutputName")
        result(rowIndex, 5) = RunRecordNumber(record, "Percent")
        result(rowIndex, 6) = RunRecordText(record, "UOM")
        result(rowIndex, 7) = ScaledRecordQty(record)
        result(rowIndex, 8) = RunRecordText(record, "OutputId")
    Next rawRecord
    ReusableRunLoaderRows = result
End Function

Public Function ReusableRunPaletteRows(Optional ByVal locationFilter As String = "") As Variant
    On Error GoTo CleanFail

    Dim entities As Variant
    Dim result() As Variant
    Dim trimmed() As Variant
    Dim rawRequirement As Variant
    Dim requirement As Object
    Dim entityRow As Long
    Dim outRow As Long
    Dim rowIndex As Long
    Dim c As Long
    Dim nodeId As String
    Dim requirementId As String
    Dim itemCode As String

    If Not mLoaded Then Exit Function
    entities = modInventoryDomainBridge.ListAvailableInventoryEntitiesBridge("")
    If Not IsArray(entities) Then Exit Function
    ReDim result(1 To mRequirements.Count * UBound(entities, 1), 1 To 10)
    locationFilter = Trim$(locationFilter)
    For Each rawRequirement In mRequirements
        Set requirement = rawRequirement
        nodeId = RunRecordText(requirement, "ProcessNodeId")
        requirementId = RunRecordText(requirement, "RequirementId")
        If RequirementHasIncomingConnection(nodeId, requirementId) Then GoTo NextRequirement
        For entityRow = LBound(entities, 1) To UBound(entities, 1)
            itemCode = Trim$(CStr(entities(entityRow, 3)))
            If Not RequirementAllowsItem(nodeId, requirementId, itemCode, CStr(entities(entityRow, 2))) Then GoTo NextEntity
            If locationFilter <> "" Then
                If StrComp(locationFilter, Trim$(CStr(entities(entityRow, 7))), vbTextCompare) <> 0 Then GoTo NextEntity
            End If
            outRow = outRow + 1
            result(outRow, 1) = nodeId
            result(outRow, 2) = requirementId
            result(outRow, 3) = RunRecordText(requirement, "RequirementName")
            result(outRow, 4) = entities(entityRow, 1)
            result(outRow, 5) = entities(entityRow, 4)
            result(outRow, 6) = AllocationPercent(nodeId, requirementId, CStr(entities(entityRow, 1)))
            result(outRow, 7) = AllocationQty(nodeId, requirementId, CStr(entities(entityRow, 1)))
            result(outRow, 8) = IIf(Trim$(CStr(entities(entityRow, 5))) <> "", entities(entityRow, 5), RunRecordText(requirement, "UOM"))
            result(outRow, 9) = ExactEntityInventoryDisplay(entities, entityRow)
            result(outRow, 10) = entities(entityRow, 7)
NextEntity:
        Next entityRow
NextRequirement:
    Next rawRequirement
    If outRow = 0 Then Exit Function
    ReDim trimmed(1 To outRow, 1 To 10)
    For rowIndex = 1 To outRow
        For c = 1 To 10
            trimmed(rowIndex, c) = result(rowIndex, c)
        Next c
    Next rowIndex
    ReusableRunPaletteRows = trimmed
CleanFail:
End Function

Public Function ApplyReusableRunAllocation(ByVal processNodeId As String, _
                                         ByVal requirementId As String, _
                                         ByVal systemKey As String, _
                                         ByVal qty As Double, _
                                         Optional ByRef report As String = "") As Boolean
    Dim requirement As Object
    Dim availableQty As Double
    Dim requiredQty As Double
    Dim otherRequirementQty As Double
    Dim otherEntityQty As Double
    Dim allocationId As String
    Dim nonCounted As Boolean

    If Not mLoaded Then
        report = "Load a released Recipe first."
        Exit Function
    End If
    Set requirement = FindRequirement(processNodeId, requirementId)
    If requirement Is Nothing Or RequirementHasIncomingConnection(processNodeId, requirementId) Then
        report = "The selected external requirement could not be resolved."
        Exit Function
    End If
    If qty < 0 Then
        report = "Allocation quantity cannot be negative."
        Exit Function
    End If
    nonCounted = ExactEntityIsNonCounted(systemKey)
    availableQty = ExactEntityAvailableQty(systemKey)
    If Not nonCounted And availableQty <= 0 And qty > 0 Then
        report = "The selected System_Key is no longer available. Refresh Production Run."
        Exit Function
    End If
    requiredQty = ScaledRecordQty(requirement)
    allocationId = AllocationKey(processNodeId, requirementId, systemKey)
    otherRequirementQty = AllocationTotalForRequirement(processNodeId, requirementId, allocationId)
    If otherRequirementQty + qty > requiredQty + QTY_TOLERANCE Then
        report = "Allocation exceeds the scaled requirement of " & FormatRunNumberLocal(requiredQty) & "."
        Exit Function
    End If
    otherEntityQty = AllocationTotalForEntity(systemKey, allocationId)
    If Not nonCounted And otherEntityQty + qty > availableQty + QTY_TOLERANCE Then
        report = "Allocation exceeds exact entity availability of " & FormatRunNumberLocal(availableQty) & "."
        Exit Function
    End If
    If qty <= QTY_TOLERANCE Then
        If mAllocations.Exists(allocationId) Then mAllocations.Remove allocationId
    ElseIf mAllocations.Exists(allocationId) Then
        mAllocations(allocationId) = qty
    Else
        mAllocations.Add allocationId, qty
    End If
    mCheckedIn = False
    mCompleted = False
    report = "Exact inventory allocation saved: " & FormatRunNumberLocal(qty) & _
             " of " & FormatRunNumberLocal(requiredQty) & "."
    ApplyReusableRunAllocation = True
End Function

Public Function CheckInReusableRun(ByVal runLocation As String, _
                                   Optional ByRef report As String = "") As Boolean
    Dim rawRequirement As Variant
    Dim requirement As Object
    Dim requiredQty As Double
    Dim allocatedQty As Double
    Dim key As Variant
    Dim systemKey As String
    Dim liveQty As Double
    Dim allocatedForEntity As Double
    Dim liveLocation As String
    Dim nonCounted As Boolean

    If Not mLoaded Then
        report = "Load a released Recipe first."
        Exit Function
    End If
    If mCompleted Then
        report = "This batch is already complete. Click Next Batch."
        Exit Function
    End If
    For Each rawRequirement In mRequirements
        Set requirement = rawRequirement
        If Not RequirementHasIncomingConnection(RunRecordText(requirement, "ProcessNodeId"), _
                                                 RunRecordText(requirement, "RequirementId")) Then
            requiredQty = ScaledRecordQty(requirement)
            allocatedQty = AllocationTotalForRequirement( _
                RunRecordText(requirement, "ProcessNodeId"), _
                RunRecordText(requirement, "RequirementId"), "")
            If Abs(allocatedQty - requiredQty) > QTY_TOLERANCE Then
                report = "Inventory is insufficient or unresolved for " & _
                         RunRecordText(requirement, "RequirementName") & ". Required=" & _
                         FormatRunNumberLocal(requiredQty) & "; allocated=" & _
                         FormatRunNumberLocal(allocatedQty) & "."
                Exit Function
            End If
        End If
    Next rawRequirement
    For Each key In mAllocations.Keys
        systemKey = AllocationSystemKey(CStr(key))
        liveQty = ExactEntityAvailableQty(systemKey, liveLocation)
        nonCounted = ExactEntityIsNonCounted(systemKey)
        allocatedForEntity = AllocationTotalForEntity(systemKey, "")
        If Not nonCounted And liveQty + QTY_TOLERANCE < allocatedForEntity Then
            report = "Stale allocation rejected for System_Key " & systemKey & _
                     ". Available=" & FormatRunNumberLocal(liveQty) & "; allocated=" & _
                     FormatRunNumberLocal(allocatedForEntity) & ". Refresh Production Run."
            Exit Function
        End If
        If Trim$(runLocation) <> "" And StrComp(Trim$(runLocation), liveLocation, vbTextCompare) <> 0 Then
            report = "System_Key " & systemKey & " is at " & liveLocation & _
                     "; the Production run location is " & Trim$(runLocation) & "."
            Exit Function
        End If
    Next key
    mCheckedIn = True
    report = "Checked in " & FormatRunNumberLocal(TotalExternalAllocation()) & _
             " units using " & CStr(mAllocations.Count) & " exact System_Key allocation(s)."
    CheckInReusableRun = True
End Function

Public Function ReusableRunManagerCheckRows() As Variant
    Dim result() As Variant
    Dim key As Variant
    Dim rowIndex As Long
    Dim entity As Variant
    Dim systemKey As String
    Dim qty As Double

    If Not mCheckedIn Or mAllocations Is Nothing Or mAllocations.Count = 0 Then Exit Function
    entity = modInventoryDomainBridge.ListAvailableInventoryEntitiesBridge("")
    ReDim result(1 To mAllocations.Count, 1 To 6)
    For Each key In mAllocations.Keys
        rowIndex = rowIndex + 1
        systemKey = AllocationSystemKey(CStr(key))
        qty = CDbl(mAllocations(key))
        result(rowIndex, 1) = systemKey
        result(rowIndex, 2) = ExactEntityField(entity, systemKey, 3)
        result(rowIndex, 3) = ExactEntityField(entity, systemKey, 4)
        result(rowIndex, 4) = ExactEntityField(entity, systemKey, 5)
        result(rowIndex, 5) = qty
        result(rowIndex, 6) = ExactEntityInventoryDisplayForKey(entity, systemKey)
    Next key
    ReusableRunManagerCheckRows = result
End Function

Public Function ReusableRunOutputRows() As Variant
    Dim result() As Variant
    Dim rawHistory As Variant
    Dim history As Object
    Dim rawOutput As Variant
    Dim output As Object
    Dim totalRows As Long
    Dim rowIndex As Long
    Dim outputKey As String

    If Not mLoaded Or mOutputs.Count = 0 Then Exit Function
    totalRows = mOutputHistory.Count
    If Not mCompleted Then totalRows = totalRows + mOutputs.Count
    If totalRows = 0 Then Exit Function
    ReDim result(1 To totalRows, 1 To 9)
    For Each rawHistory In mOutputHistory
        Set history = rawHistory
        rowIndex = rowIndex + 1
        result(rowIndex, 1) = RunRecordText(history, "ProcessName")
        result(rowIndex, 2) = RunRecordText(history, "OutputName")
        result(rowIndex, 3) = RunRecordText(history, "UOM")
        result(rowIndex, 4) = RunRecordNumber(history, "ActualQty")
        result(rowIndex, 5) = CLng(RunRecordNumber(history, "BatchNumber"))
        result(rowIndex, 6) = RunRecordNumber(history, "UsedGoods")
        result(rowIndex, 7) = RunRecordNumber(history, "ProcessTotal")
        result(rowIndex, 8) = RunRecordText(history, "Recall")
        result(rowIndex, 9) = RunRecordText(history, "System_Key")
    Next rawHistory
    If Not mCompleted Then
        For Each rawOutput In mOutputs
            Set output = rawOutput
            rowIndex = rowIndex + 1
            outputKey = OutputIdentityKey(RunRecordText(output, "ProcessNodeId"), RunRecordText(output, "OutputId"))
            result(rowIndex, 1) = RunRecordText(output, "ProcessName")
            result(rowIndex, 2) = RunRecordText(output, "OutputName")
            result(rowIndex, 3) = RunRecordText(output, "UOM")
            result(rowIndex, 5) = mBatchNumber
            result(rowIndex, 6) = ProcessUsedGoodsQty( _
                RunRecordText(output, "ProcessNodeId"), RunRecordText(output, "UOM"))
            result(rowIndex, 7) = ProcessOutputTotal( _
                RunRecordText(output, "ProcessName"), RunRecordText(output, "OutputName"), _
                RunRecordText(output, "UOM"))
            result(rowIndex, 8) = mRecipeId & "-B" & Format$(mBatchNumber, "0000")
            If mOutputKeys.Exists(outputKey) Then result(rowIndex, 9) = mOutputKeys(outputKey)
        Next rawOutput
    End If
    ReusableRunOutputRows = result
End Function

Public Function ReusableRunOutputDefinitionIndex(ByVal displayRowIndex As Long) As Long
    Dim firstCurrentRow As Long

    If Not mLoaded Or mCompleted Then Exit Function
    firstCurrentRow = mOutputHistory.Count + 1
    If displayRowIndex < firstCurrentRow Then Exit Function
    ReusableRunOutputDefinitionIndex = displayRowIndex - firstCurrentRow + 1
    If ReusableRunOutputDefinitionIndex < 1 Or _
            ReusableRunOutputDefinitionIndex > mOutputs.Count Then _
        ReusableRunOutputDefinitionIndex = 0
End Function

Public Function ReusableRunPlannedOutput(ByVal outputIndex As Long) As String
    If Not mLoaded Then Exit Function
    If outputIndex < 1 Or outputIndex > mOutputs.Count Then Exit Function
    ReusableRunPlannedOutput = FormatRunNumberLocal(ScaledRecordQty(mOutputs(outputIndex)))
End Function

Public Function StageReusableRunActualOutput(ByVal outputIndex As Long, _
                                             ByVal quantityText As String, _
                                             Optional ByRef report As String = "") As Boolean
    Dim output As Object
    Dim outputKey As String
    Dim actualQty As Double

    If Not mLoaded Then
        report = "Load a released Recipe before entering Actual Output."
        Exit Function
    End If
    If mCompleted Then
        report = "This batch is already complete. Click Next Batch before entering Actual Output."
        Exit Function
    End If
    If outputIndex < 1 Or outputIndex > mOutputs.Count Then
        report = "Select the Production Output row whose Actual Output is being entered."
        Exit Function
    End If

    Set output = mOutputs(outputIndex)
    outputKey = OutputIdentityKey(RunRecordText(output, "ProcessNodeId"), _
                                  RunRecordText(output, "OutputId"))
    quantityText = Trim$(quantityText)
    If quantityText = "" Then
        If mActualOutputQty.Exists(outputKey) Then mActualOutputQty.Remove outputKey
        report = "Actual Output cleared for " & RunRecordText(output, "OutputName") & "."
        StageReusableRunActualOutput = True
        Exit Function
    End If
    If Not IsNumeric(quantityText) Then
        report = "Actual Output must be a number greater than zero."
        Exit Function
    End If
    actualQty = CDbl(quantityText)
    If actualQty <= 0 Then
        report = "Actual Output must be a number greater than zero."
        Exit Function
    End If

    mActualOutputQty(outputKey) = actualQty
    report = "Actual Output staged for " & RunRecordText(output, "OutputName") & _
             ": " & FormatRunNumberLocal(actualQty) & " " & RunRecordText(output, "UOM") & "."
    StageReusableRunActualOutput = True
End Function

Public Function ReusableRunActualOutput(ByVal outputIndex As Long) As String
    Dim output As Object
    Dim outputKey As String

    If Not mLoaded Then Exit Function
    If outputIndex < 1 Or outputIndex > mOutputs.Count Then Exit Function
    Set output = mOutputs(outputIndex)
    outputKey = OutputIdentityKey(RunRecordText(output, "ProcessNodeId"), _
                                  RunRecordText(output, "OutputId"))
    If mActualOutputQty.Exists(outputKey) Then _
        ReusableRunActualOutput = FormatRunNumberLocal(CDbl(mActualOutputQty(outputKey)))
End Function

Public Function CompleteReusableRun(ByVal runLocation As String, _
                                    Optional ByRef report As String = "") As Boolean
    On Error GoTo Failed

    Dim recheckReport As String
    Dim rawNode As Variant
    Dim node As Object
    Dim items As Collection
    Dim eventIds As String
    Dim eventId As String
    Dim queueError As String
    Dim processorReport As String
    Dim processorReports As String
    Dim appliedCount As Long
    Dim processedNow As Long
    Dim runId As String

    If Not mCheckedIn Then
        report = "Check exact inventory into Production before completing the run."
        Exit Function
    End If
    If mCompleted Then
        report = "This batch is already complete. Click Next Batch."
        Exit Function
    End If
    If Not CheckInReusableRun(runLocation, recheckReport) Then
        report = recheckReport
        Exit Function
    End If
    If Not ValidateReusableActualOutputs(report) Then Exit Function
    AssignFreshOutputKeys
    runId = "PROD-RUN-" & Replace$(modRoleEventWriter.CreateSystemKey(), "-", "")
    For Each rawNode In mNodes
        Set node = rawNode
        Set items = BuildNodeConsumeItems(node, runLocation, runId)
        If items.Count > 0 Then
            eventId = ""
            If Not modRoleEventWriter.QueuePayloadEventCurrent(EVENT_TYPE_PROD_CONSUME, "", _
                    modProductionJson.BuildJsonArray(items), RunEventNote(runId, node, "CONSUME"), _
                    eventId, queueError) Then
                report = "Production consume event was not queued: " & queueError
                Exit Function
            End If
            AppendEventId eventIds, eventId
            processedNow = modProcessor.RunBatch(Trim$(modConfig.GetWarehouseId()), 0, processorReport)
            If processedNow < 1 Then
                report = "Production consume event " & eventId & " was not applied. " & processorReport
                Exit Function
            End If
            appliedCount = appliedCount + processedNow
            AppendProcessorReport processorReports, processorReport
        End If
        Set items = BuildNodeCompleteItems(node, runLocation, runId)
        If items.Count = 0 Then
            report = "Process " & RunRecordText(node, "ProcessName") & " has no output to complete."
            Exit Function
        End If
        eventId = ""
        If Not modRoleEventWriter.QueuePayloadEventCurrent(EVENT_TYPE_PROD_COMPLETE, "", _
                modProductionJson.BuildJsonArray(items), RunEventNote(runId, node, "COMPLETE"), _
                eventId, queueError) Then
            report = "Production complete event was not queued: " & queueError
            Exit Function
        End If
        AppendEventId eventIds, eventId
        processedNow = modProcessor.RunBatch(Trim$(modConfig.GetWarehouseId()), 0, processorReport)
        If processedNow < 1 Then
            report = "Production complete event " & eventId & " was not applied. " & processorReport
            Exit Function
        End If
        appliedCount = appliedCount + processedNow
        AppendProcessorReport processorReports, processorReport
    Next rawNode

    If Not VerifyCompletedOutputBalances(report) Then
        report = report & " Processor applied=" & CStr(appliedCount) & ". " & processorReports
        Exit Function
    End If
    CaptureLastActualOutputs
    CaptureCompletedOutputHistory
    mCompleted = True
    mLastSummary = "RunId=" & runId & "; Recipe=" & mRecipeId & " v" & mRecipeVersion & _
                   "; Batch=" & CStr(mBatchNumber) & "; Scale=" & _
                   FormatRunNumberLocal(mScalePercent) & "%; ExactInputs=" & _
                   CStr(mAllocations.Count) & "; Outputs=" & CStr(mOutputs.Count) & _
                   "; Events=" & eventIds & "; ProcessorApplied=" & CStr(appliedCount) & "."
    report = "Production batch completed and persisted. " & mLastSummary
    CompleteReusableRun = True
    Exit Function

Failed:
    report = "Reusable Production completion failed: " & Err.Description
End Function

Public Function BeginNextReusableBatch(Optional ByRef report As String = "") As Boolean
    If Not mLoaded Then
        report = "Load a released Recipe first."
        Exit Function
    End If
    If Not mCompleted Then
        report = "Complete the current batch before starting the next batch."
        Exit Function
    End If
    mBatchNumber = mBatchNumber + 1
    Set mAllocations = NewTextDictionary()
    Set mOutputKeys = NewTextDictionary()
    Set mActualOutputQty = NewTextDictionary()
    mCheckedIn = False
    mCompleted = False
    mLastSummary = ""
    report = "Next Batch " & CStr(mBatchNumber) & " is ready; exact-key allocations were cleared."
    BeginNextReusableBatch = True
End Function

Public Function ReusableRunIsLoaded() As Boolean
    ReusableRunIsLoaded = mLoaded
End Function

Public Function ReusableRunIsCheckedIn() As Boolean
    ReusableRunIsCheckedIn = mCheckedIn
End Function

Public Function ReusableRunIsCompleted() As Boolean
    ReusableRunIsCompleted = mCompleted
End Function

Public Function ReusableRunBatchNumber() As Long
    ReusableRunBatchNumber = mBatchNumber
End Function

Public Function ReusableRunScalePercent() As Double
    ReusableRunScalePercent = mScalePercent
End Function

Public Function ReusableRunLastSummary() As String
    ReusableRunLastSummary = mLastSummary
End Function

Public Function ReusableRunOutputSystemKey(ByVal processNodeId As String, _
                                           ByVal outputId As String) As String
    Dim key As String
    key = OutputIdentityKey(processNodeId, outputId)
    If Not mOutputKeys Is Nothing Then
        If mOutputKeys.Exists(key) Then ReusableRunOutputSystemKey = CStr(mOutputKeys(key))
    End If
End Function

Public Function ReusableRunExactEntityQty(ByVal systemKey As String) As Double
    ReusableRunExactEntityQty = ExactEntityAvailableQty(systemKey)
End Function

Public Function ReusableRunRequirementQty(ByVal processNodeId As String, _
                                          ByVal requirementId As String) As Double
    Dim requirement As Object
    Set requirement = FindRequirement(processNodeId, requirementId)
    If Not requirement Is Nothing Then ReusableRunRequirementQty = ScaledRecordQty(requirement)
End Function

Private Function LoadNodeProcessDefinitions(ByRef report As String) As Boolean
    Dim rawNode As Variant
    Dim node As Object
    Dim jsonText As String
    Dim parseReport As String
    Dim records As Collection
    Dim rawRecord As Variant
    Dim record As Object
    Dim enriched As Object
    Dim processName As String
    Dim statusValue As String

    For Each rawNode In mNodes
        Set node = rawNode
        jsonText = modOperationsPrimitiveBridge.GetProcessVersion( _
            RunRecordText(node, "ProcessId"), RunRecordText(node, "ProcessVersion"))
        If jsonText = "" Then
            report = "Process " & RunRecordText(node, "ProcessId") & " version " & _
                     RunRecordText(node, "ProcessVersion") & " could not be read."
            Exit Function
        End If
        parseReport = ""
        Set records = modProductionReusableDesigns.ParseReusableDefinitionRecords(jsonText, parseReport)
        If records Is Nothing Then
            report = parseReport
            Exit Function
        End If
        For Each rawRecord In records
            Set record = rawRecord
            If StrComp(RunRecordText(record, "RecordType"), "PROCESS", vbTextCompare) = 0 Then
                processName = RunRecordText(record, "ProcessName")
                statusValue = RunRecordText(record, "Status")
            End If
        Next rawRecord
        If StrComp(statusValue, "RELEASED", vbTextCompare) <> 0 Then
            report = "Recipe references a Process version that is not released: " & _
                     RunRecordText(node, "ProcessId") & " v" & RunRecordText(node, "ProcessVersion") & "."
            Exit Function
        End If
        node("ProcessName") = processName
        For Each rawRecord In records
            Set record = rawRecord
            Select Case UCase$(RunRecordText(record, "RecordType"))
                Case "REQUIREMENT", "ALTERNATIVE", "OUTPUT"
                    Set enriched = CloneRunRecord(record)
                    enriched("ProcessNodeId") = RunRecordText(node, "ProcessNodeId")
                    enriched("ProcessId") = RunRecordText(node, "ProcessId")
                    enriched("ProcessVersion") = RunRecordText(node, "ProcessVersion")
                    enriched("ProcessName") = processName
                    Select Case UCase$(RunRecordText(record, "RecordType"))
                        Case "REQUIREMENT": mRequirements.Add enriched
                        Case "ALTERNATIVE": mAlternatives.Add enriched
                        Case "OUTPUT": mOutputs.Add enriched
                    End Select
            End Select
        Next rawRecord
    Next rawNode
    LoadNodeProcessDefinitions = True
End Function

Private Function ValidateLoadedRunGraph(ByRef report As String) As Boolean
    Dim rawRequirement As Variant
    Dim requirement As Object
    Dim rawConnection As Variant
    Dim connection As Object
    Dim sourceOutput As Object
    Dim sourceNode As Object
    Dim targetNode As Object

    If mOutputs.Count < mNodes.Count Then
        report = "Every Process must declare at least one output."
        Exit Function
    End If
    For Each rawConnection In mConnections
        Set connection = rawConnection
        Set sourceNode = FindNode(RunRecordText(connection, "FromProcessNodeId"))
        Set targetNode = FindNode(RunRecordText(connection, "ToProcessNodeId"))
        Set sourceOutput = FindOutput(RunRecordText(connection, "FromProcessNodeId"), _
                                      RunRecordText(connection, "FromOutputId"))
        Set requirement = FindRequirement(RunRecordText(connection, "ToProcessNodeId"), _
                                          RunRecordText(connection, "ToRequirementId"))
        If sourceNode Is Nothing Or targetNode Is Nothing Or sourceOutput Is Nothing Or requirement Is Nothing Then
            report = "Recipe contains an unresolved Process connection."
            Exit Function
        End If
        If RunRecordNumber(sourceNode, "ExecutionOrdinal") >= RunRecordNumber(targetNode, "ExecutionOrdinal") Then
            report = "Recipe execution order is invalid or circular."
            Exit Function
        End If
        If Not UomCompatible(RunRecordText(sourceOutput, "UOM"), RunRecordText(requirement, "UOM")) Then
            report = "Recipe connection UOM is incompatible for " & RunRecordText(requirement, "RequirementName") & "."
            Exit Function
        End If
    Next rawConnection
    For Each rawRequirement In mRequirements
        Set requirement = rawRequirement
        If Not RequirementHasIncomingConnection(RunRecordText(requirement, "ProcessNodeId"), _
                                                 RunRecordText(requirement, "RequirementId")) Then
            If Not RequirementHasAlternative(RunRecordText(requirement, "ProcessNodeId"), _
                                             RunRecordText(requirement, "RequirementId")) Then
                report = "Unresolved external input: " & RunRecordText(requirement, "RequirementName") & "."
                Exit Function
            End If
        End If
        If ScaledRecordQty(requirement) <= 0 Then
            report = "Requirement quantity could not be resolved for " & RunRecordText(requirement, "RequirementName") & "."
            Exit Function
        End If
    Next rawRequirement
    ValidateLoadedRunGraph = True
End Function

Private Sub AddNodeInExecutionOrder(ByVal node As Object)
    Dim i As Long
    Dim existing As Object
    For i = 1 To mNodes.Count
        Set existing = mNodes(i)
        If RunRecordNumber(node, "ExecutionOrdinal") < RunRecordNumber(existing, "ExecutionOrdinal") Then
            mNodes.Add node, Before:=i
            Exit Sub
        End If
    Next i
    mNodes.Add node
End Sub

Private Function BuildNodeConsumeItems(ByVal node As Object, ByVal runLocation As String, _
                                       ByVal runId As String) As Collection
    Dim items As New Collection
    Dim rawRequirement As Variant
    Dim requirement As Object
    Dim rawConnection As Variant
    Dim connection As Object
    Dim key As Variant
    Dim allocationParts() As String
    Dim item As Object
    Dim sourceOutput As Object
    Dim qty As Double
    Dim systemKey As String
    Dim entityRows As Variant

    entityRows = modInventoryDomainBridge.ListAvailableInventoryEntitiesBridge("")
    For Each rawRequirement In mRequirements
        Set requirement = rawRequirement
        If StrComp(RunRecordText(requirement, "ProcessNodeId"), _
                   RunRecordText(node, "ProcessNodeId"), vbTextCompare) <> 0 Then GoTo NextRequirement
        If RequirementHasIncomingConnection(RunRecordText(node, "ProcessNodeId"), _
                                            RunRecordText(requirement, "RequirementId")) Then GoTo NextRequirement
        For Each key In mAllocations.Keys
            allocationParts = Split(CStr(key), vbTab)
            If UBound(allocationParts) = 2 Then
                If StrComp(allocationParts(0), RunRecordText(node, "ProcessNodeId"), vbTextCompare) = 0 _
                   And StrComp(allocationParts(1), RunRecordText(requirement, "RequirementId"), vbTextCompare) = 0 Then
                    qty = CDbl(mAllocations(key))
                    Set item = modProductionJson.CreateProductionDeltaPayloadItem( _
                        allocationParts(2), CStr(ExactEntityField(entityRows, allocationParts(2), 2)), _
                        qty, CStr(ExactEntityField(entityRows, allocationParts(2), 7)), _
                        RunItemNote(runId, node, RunRecordText(requirement, "RequirementId")), "USED")
                    items.Add item
                End If
            End If
        Next key
NextRequirement:
    Next rawRequirement
    For Each rawConnection In mConnections
        Set connection = rawConnection
        If StrComp(RunRecordText(connection, "ToProcessNodeId"), _
                   RunRecordText(node, "ProcessNodeId"), vbTextCompare) = 0 Then
            Set sourceOutput = FindOutput(RunRecordText(connection, "FromProcessNodeId"), _
                                          RunRecordText(connection, "FromOutputId"))
            qty = ScaledConnectionQty(connection)
            systemKey = CStr(mOutputKeys(OutputIdentityKey( _
                RunRecordText(connection, "FromProcessNodeId"), _
                RunRecordText(connection, "FromOutputId"))))
            Set item = modProductionJson.CreateProductionDeltaPayloadItem( _
                systemKey, RunRecordText(sourceOutput, "ITEM_CODE"), qty, runLocation, _
                RunItemNote(runId, node, RunRecordText(connection, "ToRequirementId")), "USED")
            items.Add item
        End If
    Next rawConnection
    Set BuildNodeConsumeItems = items
End Function

Private Function BuildNodeCompleteItems(ByVal node As Object, ByVal runLocation As String, _
                                        ByVal runId As String) As Collection
    Dim items As New Collection
    Dim rawOutput As Variant
    Dim output As Object
    Dim item As Object
    Dim outputKey As String
    Dim attributesJson As String

    For Each rawOutput In mOutputs
        Set output = rawOutput
        If StrComp(RunRecordText(output, "ProcessNodeId"), _
                   RunRecordText(node, "ProcessNodeId"), vbTextCompare) = 0 Then
            outputKey = CStr(mOutputKeys(OutputIdentityKey( _
                RunRecordText(node, "ProcessNodeId"), RunRecordText(output, "OutputId"))))
            attributesJson = "{""RunId"":""" & runId & """,""RecipeId"":""" & _
                             mRecipeId & """,""RecipeVersion"":""" & mRecipeVersion & _
                             """,""ProcessId"":""" & RunRecordText(node, "ProcessId") & _
                             """,""ProcessVersion"":""" & RunRecordText(node, "ProcessVersion") & _
                             """,""OutputId"":""" & RunRecordText(output, "OutputId") & _
                             """,""PlannedQty"":" & JsonRunNumber(ScaledRecordQty(output)) & _
                             ",""ActualQty"":" & JsonRunNumber(ActualOutputQty(output)) & "}"
            Set item = modProductionJson.CreateProductionInventoryEntityPayloadItem( _
                outputKey, RunRecordText(output, "ITEM_CODE"), ActualOutputQty(output), _
                runLocation, "GOOD", attributesJson, RunItemNote(runId, node, RunRecordText(output, "OutputId")))
            item("IoType") = "MADE"
            item("ITEM_CODE") = RunRecordText(output, "ITEM_CODE")
            item("ITEM") = RunRecordText(output, "OutputName")
            item("UOM") = RunRecordText(output, "UOM")
            items.Add item
        End If
    Next rawOutput
    Set BuildNodeCompleteItems = items
End Function

Private Sub AssignFreshOutputKeys()
    Dim rawOutput As Variant
    Dim output As Object
    Dim key As String
    Set mOutputKeys = NewTextDictionary()
    For Each rawOutput In mOutputs
        Set output = rawOutput
        key = OutputIdentityKey(RunRecordText(output, "ProcessNodeId"), RunRecordText(output, "OutputId"))
        mOutputKeys.Add key, modRoleEventWriter.CreateSystemKey()
    Next rawOutput
End Sub

Private Function VerifyCompletedOutputBalances(ByRef report As String) As Boolean
    Dim rawOutput As Variant
    Dim output As Object
    Dim key As String
    Dim expectedQty As Double
    Dim actualQty As Double

    For Each rawOutput In mOutputs
        Set output = rawOutput
        key = CStr(mOutputKeys(OutputIdentityKey(RunRecordText(output, "ProcessNodeId"), _
                                                RunRecordText(output, "OutputId"))))
        expectedQty = ActualOutputQty(output) - OutgoingQtyForOutput( _
            RunRecordText(output, "ProcessNodeId"), RunRecordText(output, "OutputId"))
        actualQty = ExactEntityAvailableQty(key)
        If Abs(actualQty - expectedQty) > QTY_TOLERANCE Then
            report = "Output persistence verification failed for System_Key " & key & _
                     ". Expected=" & FormatRunNumberLocal(expectedQty) & "; actual=" & _
                     FormatRunNumberLocal(actualQty) & "."
            Exit Function
        End If
    Next rawOutput
    VerifyCompletedOutputBalances = True
End Function

Private Function ValidateReusableActualOutputs(ByRef report As String) As Boolean
    Dim rawOutput As Variant
    Dim output As Object
    Dim outputKey As String
    Dim actualQty As Double
    Dim committedQty As Double

    For Each rawOutput In mOutputs
        Set output = rawOutput
        outputKey = OutputIdentityKey(RunRecordText(output, "ProcessNodeId"), _
                                      RunRecordText(output, "OutputId"))
        If Not mActualOutputQty.Exists(outputKey) Then
            report = "Enter Actual Output for " & RunRecordText(output, "OutputName") & _
                     " before completing the run."
            Exit Function
        End If
        actualQty = ActualOutputQty(output)
        If actualQty <= 0 Then
            report = "Actual Output for " & RunRecordText(output, "OutputName") & _
                     " must be greater than zero."
            Exit Function
        End If
        committedQty = OutgoingQtyForOutput(RunRecordText(output, "ProcessNodeId"), _
                                            RunRecordText(output, "OutputId"))
        If actualQty + QTY_TOLERANCE < committedQty Then
            report = "Actual Output for " & RunRecordText(output, "OutputName") & _
                     " is below its routed downstream commitment. Actual=" & _
                     FormatRunNumberLocal(actualQty) & "; committed=" & _
                     FormatRunNumberLocal(committedQty) & "."
            Exit Function
        End If
    Next rawOutput
    ValidateReusableActualOutputs = True
End Function

Private Function ActualOutputQty(ByVal output As Object) As Double
    Dim outputKey As String

    outputKey = OutputIdentityKey(RunRecordText(output, "ProcessNodeId"), _
                                  RunRecordText(output, "OutputId"))
    If Not mActualOutputQty Is Nothing Then
        If mActualOutputQty.Exists(outputKey) Then ActualOutputQty = CDbl(mActualOutputQty(outputKey))
    End If
End Function

Private Sub CaptureLastActualOutputs()
    Dim key As Variant

    Set mLastOutputQty = NewTextDictionary()
    For Each key In mActualOutputQty.Keys
        mLastOutputQty(CStr(key)) = CDbl(mActualOutputQty(key))
    Next key
End Sub

Private Sub CaptureCompletedOutputHistory()
    Dim rawOutput As Variant
    Dim output As Object
    Dim record As Object
    Dim outputKey As String

    For Each rawOutput In mOutputs
        Set output = rawOutput
        outputKey = OutputIdentityKey(RunRecordText(output, "ProcessNodeId"), _
                                      RunRecordText(output, "OutputId"))
        Set record = NewTextDictionary()
        record("ProcessName") = RunRecordText(output, "ProcessName")
        record("OutputName") = RunRecordText(output, "OutputName")
        record("UOM") = RunRecordText(output, "UOM")
        record("ActualQty") = ActualOutputQty(output)
        record("BatchNumber") = mBatchNumber
        record("UsedGoods") = ProcessUsedGoodsQty( _
            RunRecordText(output, "ProcessNodeId"), RunRecordText(output, "UOM"))
        record("Recall") = mRecipeId & "-B" & Format$(mBatchNumber, "0000")
        record("System_Key") = CStr(mOutputKeys(outputKey))
        mOutputHistory.Add record
        record("ProcessTotal") = ProcessOutputTotal( _
            RunRecordText(output, "ProcessName"), RunRecordText(output, "OutputName"), _
            RunRecordText(output, "UOM"))
    Next rawOutput
End Sub

Private Function ProcessUsedGoodsQty(ByVal nodeId As String, ByVal outputUom As String) As Double
    Dim rawRequirement As Variant
    Dim requirement As Object
    Dim requirementUom As String

    For Each rawRequirement In mRequirements
        Set requirement = rawRequirement
        If StrComp(RunRecordText(requirement, "ProcessNodeId"), nodeId, vbTextCompare) = 0 Then
            requirementUom = RunRecordText(requirement, "UOM")
            If UomCompatible(requirementUom, outputUom) Then _
                ProcessUsedGoodsQty = ProcessUsedGoodsQty + ScaledRecordQty(requirement)
        End If
    Next rawRequirement
End Function

Private Function ProcessOutputTotal(ByVal processName As String, _
                                    ByVal outputName As String, _
                                    ByVal uom As String) As Double
    Dim rawHistory As Variant
    Dim history As Object

    For Each rawHistory In mOutputHistory
        Set history = rawHistory
        If StrComp(RunRecordText(history, "ProcessName"), processName, vbTextCompare) = 0 _
           And StrComp(RunRecordText(history, "OutputName"), outputName, vbTextCompare) = 0 _
           And StrComp(RunRecordText(history, "UOM"), uom, vbTextCompare) = 0 Then
            ProcessOutputTotal = ProcessOutputTotal + RunRecordNumber(history, "ActualQty")
        End If
    Next rawHistory
End Function

Private Function OutgoingQtyForOutput(ByVal nodeId As String, ByVal outputId As String) As Double
    Dim rawConnection As Variant
    Dim connection As Object
    For Each rawConnection In mConnections
        Set connection = rawConnection
        If StrComp(RunRecordText(connection, "FromProcessNodeId"), nodeId, vbTextCompare) = 0 _
           And StrComp(RunRecordText(connection, "FromOutputId"), outputId, vbTextCompare) = 0 Then
            OutgoingQtyForOutput = OutgoingQtyForOutput + ScaledConnectionQty(connection)
        End If
    Next rawConnection
End Function

Private Function FindNode(ByVal nodeId As String) As Object
    Dim rawRecord As Variant
    Dim record As Object
    For Each rawRecord In mNodes
        Set record = rawRecord
        If StrComp(RunRecordText(record, "ProcessNodeId"), nodeId, vbTextCompare) = 0 Then
            Set FindNode = record
            Exit Function
        End If
    Next rawRecord
End Function

Private Function FindRequirement(ByVal nodeId As String, ByVal requirementId As String) As Object
    Dim rawRecord As Variant
    Dim record As Object
    For Each rawRecord In mRequirements
        Set record = rawRecord
        If StrComp(RunRecordText(record, "ProcessNodeId"), nodeId, vbTextCompare) = 0 _
           And StrComp(RunRecordText(record, "RequirementId"), requirementId, vbTextCompare) = 0 Then
            Set FindRequirement = record
            Exit Function
        End If
    Next rawRecord
End Function

Private Function FindOutput(ByVal nodeId As String, ByVal outputId As String) As Object
    Dim rawRecord As Variant
    Dim record As Object
    For Each rawRecord In mOutputs
        Set record = rawRecord
        If StrComp(RunRecordText(record, "ProcessNodeId"), nodeId, vbTextCompare) = 0 _
           And StrComp(RunRecordText(record, "OutputId"), outputId, vbTextCompare) = 0 Then
            Set FindOutput = record
            Exit Function
        End If
    Next rawRecord
End Function

Private Function RequirementHasIncomingConnection(ByVal nodeId As String, _
                                                  ByVal requirementId As String) As Boolean
    Dim rawRecord As Variant
    Dim record As Object
    For Each rawRecord In mConnections
        Set record = rawRecord
        If StrComp(RunRecordText(record, "ToProcessNodeId"), nodeId, vbTextCompare) = 0 _
           And StrComp(RunRecordText(record, "ToRequirementId"), requirementId, vbTextCompare) = 0 Then
            RequirementHasIncomingConnection = True
            Exit Function
        End If
    Next rawRecord
End Function

Private Function RequirementHasAlternative(ByVal nodeId As String, _
                                           ByVal requirementId As String) As Boolean
    Dim rawRecord As Variant
    Dim record As Object
    For Each rawRecord In mAlternatives
        Set record = rawRecord
        If StrComp(RunRecordText(record, "ProcessNodeId"), nodeId, vbTextCompare) = 0 _
           And StrComp(RunRecordText(record, "RequirementId"), requirementId, vbTextCompare) = 0 _
           And RunRecordText(record, "ITEM_CODE") <> "" Then
            RequirementHasAlternative = True
            Exit Function
        End If
    Next rawRecord
End Function

Private Function RequirementAllowsItem(ByVal nodeId As String, ByVal requirementId As String, _
                                       ByVal itemCode As String, ByVal sku As String) As Boolean
    Dim rawRecord As Variant
    Dim record As Object
    Dim allowedCode As String
    For Each rawRecord In mAlternatives
        Set record = rawRecord
        If StrComp(RunRecordText(record, "ProcessNodeId"), nodeId, vbTextCompare) = 0 _
           And StrComp(RunRecordText(record, "RequirementId"), requirementId, vbTextCompare) = 0 Then
            allowedCode = RunRecordText(record, "ITEM_CODE")
            If StrComp(allowedCode, itemCode, vbTextCompare) = 0 _
               Or StrComp(allowedCode, sku, vbTextCompare) = 0 Then
                RequirementAllowsItem = True
                Exit Function
            End If
        End If
    Next rawRecord
End Function

Private Function ScaledRecordQty(ByVal record As Object) As Double
    Dim qty As Double
    Dim pct As Double
    Dim yieldBasis As Double
    qty = RunRecordNumber(record, "Qty")
    pct = RunRecordNumber(record, "Percent")
    yieldBasis = RunRecordNumber(record, "YieldBasis")
    If qty > 0 Then
        ScaledRecordQty = qty * mScalePercent / 100#
    ElseIf pct > 0 And yieldBasis > 0 Then
        ScaledRecordQty = yieldBasis * pct * mScalePercent / 10000#
    End If
End Function

Private Function ScaledConnectionQty(ByVal connection As Object) As Double
    Dim qty As Double
    Dim pct As Double
    Dim requirement As Object
    qty = RunRecordNumber(connection, "Qty")
    pct = RunRecordNumber(connection, "Percent")
    If qty > 0 Then
        ScaledConnectionQty = qty * mScalePercent / 100#
    ElseIf pct > 0 Then
        Set requirement = FindRequirement(RunRecordText(connection, "ToProcessNodeId"), _
                                          RunRecordText(connection, "ToRequirementId"))
        If Not requirement Is Nothing Then ScaledConnectionQty = ScaledRecordQty(requirement) * pct / 100#
    End If
End Function

Private Function AllocationKey(ByVal nodeId As String, ByVal requirementId As String, _
                               ByVal systemKey As String) As String
    AllocationKey = nodeId & vbTab & requirementId & vbTab & systemKey
End Function

Private Function AllocationSystemKey(ByVal allocationKeyValue As String) As String
    Dim parts() As String
    parts = Split(allocationKeyValue, vbTab)
    If UBound(parts) >= 2 Then AllocationSystemKey = parts(2)
End Function

Private Function AllocationQty(ByVal nodeId As String, ByVal requirementId As String, _
                               ByVal systemKey As String) As Variant
    Dim key As String
    key = AllocationKey(nodeId, requirementId, systemKey)
    If Not mAllocations Is Nothing Then
        If mAllocations.Exists(key) Then AllocationQty = mAllocations(key)
    End If
End Function

Private Function AllocationPercent(ByVal nodeId As String, ByVal requirementId As String, _
                                   ByVal systemKey As String) As Variant
    Dim requirement As Object
    Dim qty As Variant
    Dim requiredQty As Double
    qty = AllocationQty(nodeId, requirementId, systemKey)
    If IsEmpty(qty) Then Exit Function
    Set requirement = FindRequirement(nodeId, requirementId)
    If requirement Is Nothing Then Exit Function
    requiredQty = ScaledRecordQty(requirement)
    If requiredQty > 0 Then AllocationPercent = CDbl(qty) / requiredQty * 100#
End Function

Private Function AllocationTotalForRequirement(ByVal nodeId As String, ByVal requirementId As String, _
                                               ByVal excludedKey As String) As Double
    Dim key As Variant
    Dim parts() As String
    For Each key In mAllocations.Keys
        If StrComp(CStr(key), excludedKey, vbTextCompare) <> 0 Then
            parts = Split(CStr(key), vbTab)
            If UBound(parts) = 2 Then
                If StrComp(parts(0), nodeId, vbTextCompare) = 0 _
                   And StrComp(parts(1), requirementId, vbTextCompare) = 0 Then
                    AllocationTotalForRequirement = AllocationTotalForRequirement + CDbl(mAllocations(key))
                End If
            End If
        End If
    Next key
End Function

Private Function AllocationTotalForEntity(ByVal systemKey As String, ByVal excludedKey As String) As Double
    Dim key As Variant
    For Each key In mAllocations.Keys
        If StrComp(CStr(key), excludedKey, vbTextCompare) <> 0 _
           And StrComp(AllocationSystemKey(CStr(key)), systemKey, vbTextCompare) = 0 Then
            AllocationTotalForEntity = AllocationTotalForEntity + CDbl(mAllocations(key))
        End If
    Next key
End Function

Private Function TotalExternalAllocation() As Double
    Dim key As Variant
    For Each key In mAllocations.Keys
        TotalExternalAllocation = TotalExternalAllocation + CDbl(mAllocations(key))
    Next key
End Function

Private Function ExactEntityAvailableQty(ByVal systemKey As String, _
                                         Optional ByRef locationOut As String = "") As Double
    Dim entities As Variant
    Dim r As Long
    entities = modInventoryDomainBridge.ListAvailableInventoryEntitiesBridge("")
    If Not IsArray(entities) Then Exit Function
    For r = LBound(entities, 1) To UBound(entities, 1)
        If StrComp(Trim$(CStr(entities(r, 1))), Trim$(systemKey), vbTextCompare) = 0 Then
            If IsNumeric(entities(r, 6)) Then ExactEntityAvailableQty = CDbl(entities(r, 6))
            locationOut = Trim$(CStr(entities(r, 7)))
            Exit Function
        End If
    Next r
End Function

Private Function ExactEntityIsNonCounted(ByVal systemKey As String) As Boolean
    Dim entities As Variant
    Dim r As Long
    Dim trackQty As String
    Dim itemKind As String
    Dim categoryValue As String

    entities = modInventoryDomainBridge.ListAvailableInventoryEntitiesBridge("")
    If Not IsArray(entities) Then Exit Function
    For r = LBound(entities, 1) To UBound(entities, 1)
        If StrComp(Trim$(CStr(entities(r, 1))), Trim$(systemKey), vbTextCompare) = 0 Then
            If UBound(entities, 2) >= 11 Then trackQty = UCase$(Trim$(CStr(entities(r, 11))))
            If UBound(entities, 2) >= 12 Then itemKind = UCase$(Trim$(CStr(entities(r, 12))))
            If UBound(entities, 2) >= 13 Then categoryValue = UCase$(Trim$(CStr(entities(r, 13))))
            ExactEntityIsNonCounted = (trackQty = "FALSE" Or trackQty = "NO" Or trackQty = "0" _
                Or itemKind = "UTILITY" Or itemKind = "SERVICE" Or itemKind = "NON_COUNTED" _
                Or categoryValue = "UTILITY" Or categoryValue = "SERVICE")
            Exit Function
        End If
    Next r
End Function

Private Function ExactEntityField(ByVal entities As Variant, ByVal systemKey As String, _
                                  ByVal columnIndex As Long) As Variant
    Dim r As Long
    If Not IsArray(entities) Then Exit Function
    For r = LBound(entities, 1) To UBound(entities, 1)
        If StrComp(Trim$(CStr(entities(r, 1))), Trim$(systemKey), vbTextCompare) = 0 Then
            ExactEntityField = entities(r, columnIndex)
            Exit Function
        End If
    Next r
End Function

Private Function ExactEntityInventoryDisplayForKey(ByVal entities As Variant, _
                                                   ByVal systemKey As String) As Variant
    Dim r As Long

    If Not IsArray(entities) Then Exit Function
    For r = LBound(entities, 1) To UBound(entities, 1)
        If StrComp(Trim$(CStr(entities(r, 1))), Trim$(systemKey), vbTextCompare) = 0 Then
            ExactEntityInventoryDisplayForKey = ExactEntityInventoryDisplay(entities, r)
            Exit Function
        End If
    Next r
End Function

Private Function ExactEntityInventoryDisplay(ByVal entities As Variant, _
                                             ByVal entityRow As Long) As Variant
    Dim trackQty As String
    Dim itemKind As String
    Dim categoryValue As String

    If Not IsArray(entities) Then Exit Function
    If UBound(entities, 2) >= 11 Then trackQty = UCase$(Trim$(CStr(entities(entityRow, 11))))
    If UBound(entities, 2) >= 12 Then itemKind = UCase$(Trim$(CStr(entities(entityRow, 12))))
    If UBound(entities, 2) >= 13 Then categoryValue = UCase$(Trim$(CStr(entities(entityRow, 13))))
    If trackQty = "FALSE" Or trackQty = "NO" Or trackQty = "0" _
       Or itemKind = "UTILITY" Or itemKind = "SERVICE" Or itemKind = "NON_COUNTED" _
       Or categoryValue = "UTILITY" Or categoryValue = "SERVICE" Then
        Select Case itemKind
            Case "UTILITY": ExactEntityInventoryDisplay = "Utility"
            Case "SERVICE": ExactEntityInventoryDisplay = "Service"
            Case Else
                If categoryValue = "UTILITY" Then
                    ExactEntityInventoryDisplay = "Utility"
                ElseIf categoryValue = "SERVICE" Then
                    ExactEntityInventoryDisplay = "Service"
                Else
                    ExactEntityInventoryDisplay = "Not counted"
                End If
        End Select
    Else
        ExactEntityInventoryDisplay = entities(entityRow, 6)
    End If
End Function

Public Function ReusableRunInventoryDisplayForTest(ByVal trackQty As String, _
                                                   ByVal itemKind As String, _
                                                   ByVal categoryValue As String, _
                                                   ByVal qty As Double) As String
    Dim entities(1 To 1, 1 To 13) As Variant

    entities(1, 6) = qty
    entities(1, 11) = trackQty
    entities(1, 12) = itemKind
    entities(1, 13) = categoryValue
    ReusableRunInventoryDisplayForTest = CStr(ExactEntityInventoryDisplay(entities, 1))
End Function

Private Function OutputIdentityKey(ByVal nodeId As String, ByVal outputId As String) As String
    OutputIdentityKey = nodeId & vbTab & outputId
End Function

Private Function UomCompatible(ByVal sourceUom As String, ByVal targetUom As String) As Boolean
    sourceUom = Trim$(sourceUom)
    targetUom = Trim$(targetUom)
    UomCompatible = (sourceUom = "" Or targetUom = "" Or _
                     StrComp(sourceUom, targetUom, vbTextCompare) = 0)
End Function

Private Function CloneRunRecord(ByVal source As Object) As Object
    Dim result As Object
    Dim key As Variant
    Set result = NewTextDictionary()
    For Each key In source.Keys
        result(CStr(key)) = source(key)
    Next key
    Set CloneRunRecord = result
End Function

Private Function NewTextDictionary() As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    Set NewTextDictionary = result
End Function

Private Function RunRecordText(ByVal record As Object, ByVal fieldName As String) As String
    If record Is Nothing Then Exit Function
    If record.Exists(fieldName) Then
        If Not IsNull(record(fieldName)) And Not IsEmpty(record(fieldName)) Then _
            RunRecordText = Trim$(CStr(record(fieldName)))
    End If
End Function

Private Function RunRecordNumber(ByVal record As Object, ByVal fieldName As String) As Double
    If record Is Nothing Then Exit Function
    If record.Exists(fieldName) Then
        If IsNumeric(record(fieldName)) Then RunRecordNumber = CDbl(record(fieldName))
    End If
End Function

Private Function FormatRunNumberLocal(ByVal valueIn As Double) As String
    FormatRunNumberLocal = Format$(valueIn, "0.#########")
End Function

Private Function JsonRunNumber(ByVal valueIn As Double) As String
    JsonRunNumber = Replace$(CStr(valueIn), Application.International(xlDecimalSeparator), ".")
End Function

Private Function RunEventNote(ByVal runId As String, ByVal node As Object, _
                              ByVal stageName As String) As String
    RunEventNote = "PRODUCTION_REUSABLE_RUN|RunId=" & runId & "|Recipe=" & mRecipeId & _
                   "|RecipeVersion=" & mRecipeVersion & "|Batch=" & CStr(mBatchNumber) & _
                   "|ProcessNode=" & RunRecordText(node, "ProcessNodeId") & _
                   "|Process=" & RunRecordText(node, "ProcessId") & _
                   "|ProcessVersion=" & RunRecordText(node, "ProcessVersion") & _
                   "|Stage=" & stageName
End Function

Private Function RunItemNote(ByVal runId As String, ByVal node As Object, _
                             ByVal lineId As String) As String
    RunItemNote = RunEventNote(runId, node, "LINE") & "|Line=" & lineId
End Function

Private Sub AppendEventId(ByRef eventIds As String, ByVal eventId As String)
    If eventIds <> "" Then eventIds = eventIds & ","
    eventIds = eventIds & eventId
End Sub

Private Sub AppendProcessorReport(ByRef reports As String, ByVal report As String)
    If Trim$(report) = "" Then Exit Sub
    If reports <> "" Then reports = reports & " | "
    reports = reports & report
End Sub
