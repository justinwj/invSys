Attribute VB_Name = "modDesignsApply"
Option Explicit

Public Const DESIGN_STATUS_APPLIED As String = "APPLIED"
Public Const DESIGN_STATUS_SKIP_DUP As String = "SKIP_DUP"

Public Const DESIGN_EVENT_CREATE As String = "DESIGN_CREATE"
Public Const DESIGN_EVENT_RELEASE As String = "DESIGN_RELEASE"
Public Const DESIGN_EVENT_OBSOLETE As String = "DESIGN_OBSOLETE"
Public Const PROCESS_EVENT_SAVE As String = "PROCESS_SAVE"
Public Const PROCESS_EVENT_RELEASE As String = "PROCESS_RELEASE"
Public Const PROCESS_EVENT_OBSOLETE As String = "PROCESS_OBSOLETE"
Public Const RECIPE_EVENT_SAVE As String = "RECIPE_SAVE"
Public Const RECIPE_EVENT_RELEASE As String = "RECIPE_RELEASE"
Public Const RECIPE_EVENT_OBSOLETE As String = "RECIPE_OBSOLETE"

Public Function ApplyDesignEvent(ByVal evt As Object, _
                                 Optional ByVal designsWb As Workbook = Nothing, _
                                 Optional ByVal runId As String = "", _
                                 Optional ByRef statusOut As String = "", _
                                 Optional ByRef errorCode As String = "", _
                                 Optional ByRef errorMessage As String = "") As Boolean
    On Error GoTo FailApply

    Dim wb As Workbook
    Dim loEvents As ListObject
    Dim loApplied As ListObject
    Dim eventId As String
    Dim eventType As String
    Dim designId As String
    Dim designVersion As String
    Dim payloadJson As String
    Dim occurredAt As Date
    Dim appliedAt As Date
    Dim appliedSeq As Long
    Dim report As String
    Dim lr As ListRow

    Set wb = modDesignsRuntime.ResolveDesignsWorkbook(GetEventTextDesign(evt, "WarehouseId"), designsWb, report)
    If wb Is Nothing Then
        errorCode = "DESIGNS_WORKBOOK_NOT_FOUND"
        errorMessage = report
        Exit Function
    End If
    If Not modDesignsSchema.EnsureDesignsSchema(wb, report) Then
        errorCode = "DESIGNS_SCHEMA_INVALID"
        errorMessage = report
        Exit Function
    End If

    Set loEvents = FindDesignsApplyTable(wb, "tblDesignEvents")
    Set loApplied = FindDesignsApplyTable(wb, "tblAppliedDesignEvents")
    eventId = GetEventTextDesign(evt, "EventID")
    eventType = UCase$(GetEventTextDesign(evt, "EventType"))
    designId = GetEventTextDesign(evt, "DesignId")
    designVersion = GetEventTextDesign(evt, "DesignVersion")
    payloadJson = GetEventTextDesign(evt, "PayloadJson")

    If eventId = "" Or designId = "" Or designVersion = "" Then
        errorCode = "INVALID_DESIGN_EVENT"
        errorMessage = "EventID, DesignId, and DesignVersion are required."
        Exit Function
    End If
    If GetEventTextDesign(evt, "WarehouseId") = "" _
       Or GetEventTextDesign(evt, "StationId") = "" _
       Or GetEventTextDesign(evt, "UserId") = "" Then
        errorCode = "INVALID_DESIGN_EVENT"
        errorMessage = "WarehouseId, StationId, and UserId are required."
        Exit Function
    End If
    If Not TryGetEventDateDesign(evt, "CreatedAtUTC", occurredAt) Then
        errorCode = "INVALID_DESIGN_EVENT"
        errorMessage = "CreatedAtUTC is required and must be a valid date."
        Exit Function
    End If
    If AppliedDesignEventExists(loApplied, eventId) Then
        statusOut = DESIGN_STATUS_SKIP_DUP
        ApplyDesignEvent = True
        Exit Function
    End If
    ' Lifecycle decisions are derived from authoritative event history, never
    ' from whatever projection rows happen to be present in the workbook.
    If Not RebuildDesignProjections(wb, report) Then
        errorCode = "DESIGN_PROJECTION_FAILED"
        errorMessage = report
        Exit Function
    End If
    If Not ValidateLifecycleTransition(wb, eventType, designId, designVersion, payloadJson, errorCode, errorMessage) Then
        Exit Function
    End If

    appliedAt = Now
    appliedSeq = NextDesignAppliedSeq(loApplied)
    If runId = "" Then runId = "DESIGN-RUN-" & Format$(appliedAt, "yyyymmddhhnnss")

    Set lr = loEvents.ListRows.Add
    SetDesignTableValue loEvents, lr.Index, "EventID", eventId
    SetDesignTableValue loEvents, lr.Index, "UndoOfEventId", GetEventTextDesign(evt, "UndoOfEventId")
    SetDesignTableValue loEvents, lr.Index, "AppliedSeq", appliedSeq
    SetDesignTableValue loEvents, lr.Index, "EventType", eventType
    SetDesignTableValue loEvents, lr.Index, "OccurredAtUTC", occurredAt
    SetDesignTableValue loEvents, lr.Index, "AppliedAtUTC", appliedAt
    SetDesignTableValue loEvents, lr.Index, "WarehouseId", GetEventTextDesign(evt, "WarehouseId")
    SetDesignTableValue loEvents, lr.Index, "StationId", GetEventTextDesign(evt, "StationId")
    SetDesignTableValue loEvents, lr.Index, "UserId", GetEventTextDesign(evt, "UserId")
    SetDesignTableValue loEvents, lr.Index, "DefinitionType", DefinitionTypeForEvent(eventType)
    SetDesignTableValue loEvents, lr.Index, "DefinitionId", designId
    SetDesignTableValue loEvents, lr.Index, "DefinitionVersion", designVersion
    SetDesignTableValue loEvents, lr.Index, "DesignId", designId
    SetDesignTableValue loEvents, lr.Index, "DesignVersion", designVersion
    SetDesignTableValue loEvents, lr.Index, "PayloadJson", payloadJson
    SetDesignTableValue loEvents, lr.Index, "Note", GetEventTextDesign(evt, "Note")

    Set lr = loApplied.ListRows.Add
    SetDesignTableValue loApplied, lr.Index, "EventID", eventId
    SetDesignTableValue loApplied, lr.Index, "UndoOfEventId", GetEventTextDesign(evt, "UndoOfEventId")
    SetDesignTableValue loApplied, lr.Index, "AppliedSeq", appliedSeq
    SetDesignTableValue loApplied, lr.Index, "AppliedAtUTC", appliedAt
    SetDesignTableValue loApplied, lr.Index, "RunId", runId
    SetDesignTableValue loApplied, lr.Index, "SourceInbox", GetEventTextDesign(evt, "SourceInbox")
    SetDesignTableValue loApplied, lr.Index, "Status", DESIGN_STATUS_APPLIED

    If Not RebuildDesignProjections(wb, report) Then
        errorCode = "DESIGN_PROJECTION_FAILED"
        errorMessage = report
        Exit Function
    End If
    SaveDesignsWorkbook wb
    statusOut = DESIGN_STATUS_APPLIED
    ApplyDesignEvent = True
    Exit Function

FailApply:
    If errorCode = "" Then errorCode = "DESIGN_APPLY_EXCEPTION"
    If errorMessage = "" Then errorMessage = CStr(Err.Number) & ": " & Err.Description
End Function

Public Function RebuildDesignProjections(ByVal designsWb As Workbook, _
                                         Optional ByRef report As String = "") As Boolean
    On Error GoTo FailRebuild

    Dim loEvents As ListObject
    Dim loDesigns As ListObject
    Dim loLines As ListObject
    Dim reusableTables As Variant
    Dim reusableTableName As Variant
    Dim values As Variant
    Dim r As Long
    Dim replayOrder As Variant
    Dim orderIndex As Long
    Dim errorMessage As String

    If designsWb Is Nothing Then
        report = "An explicit authoritative Designs workbook is required."
        Exit Function
    End If
    If designsWb.IsAddin Then
        report = "An explicit authoritative Designs workbook is required."
        Exit Function
    End If
    If Not modDesignsSchema.EnsureDesignsSchema(designsWb, report) Then Exit Function
    Set loEvents = FindDesignsApplyTable(designsWb, "tblDesignEvents")
    Set loDesigns = FindDesignsApplyTable(designsWb, "tblDesigns")
    Set loLines = FindDesignsApplyTable(designsWb, "tblDesignLines")
    ClearDesignTableRows loDesigns
    ClearDesignTableRows loLines
    reusableTables = Array("tblProcesses", "tblProcessRequirements", _
        "tblProcessIngredientAlternatives", "tblProcessOutputs", _
        "tblProcessInstructions", "tblRecipes", "tblRecipeProcesses", _
        "tblRecipeConnections")
    For Each reusableTableName In reusableTables
        ClearDesignTableRows FindDesignsApplyTable(designsWb, CStr(reusableTableName))
    Next reusableTableName

    If Not loEvents.DataBodyRange Is Nothing Then
        values = loEvents.DataBodyRange.Value
        replayOrder = DesignEventReplayOrder(loEvents, values, errorMessage)
        If IsEmpty(replayOrder) Then
            report = "Projection replay order is invalid: " & errorMessage
            Exit Function
        End If
        For orderIndex = LBound(replayOrder) To UBound(replayOrder)
            r = CLng(replayOrder(orderIndex))
            If IsReusableProductionEvent( _
                    ReadDesignTableText(loEvents, values, r, "EventType")) Then
                If Not ReplayReusableProductionEvent( _
                        loEvents, values, r, designsWb, errorMessage) Then
                    report = "Projection replay failed at event row " & CStr(r) & ": " & errorMessage
                    Exit Function
                End If
            ElseIf Not ReplayDesignEvent(loEvents, values, r, loDesigns, loLines, errorMessage) Then
                report = "Projection replay failed at event row " & CStr(r) & ": " & errorMessage
                Exit Function
            End If
        Next orderIndex
    End If
    report = "OK"
    RebuildDesignProjections = True
    Exit Function

FailRebuild:
    report = "RebuildDesignProjections failed: " & Err.Description
End Function

Private Function DesignEventReplayOrder(ByVal loEvents As ListObject, ByVal values As Variant, _
                                        ByRef errorMessage As String) As Variant
    Dim rowCount As Long
    Dim order() As Long
    Dim seqs() As Long
    Dim seen As Object
    Dim i As Long
    Dim j As Long
    Dim seqValue As Variant
    Dim swapValue As Long

    rowCount = UBound(values, 1)
    If rowCount <= 0 Then Exit Function
    ReDim order(1 To rowCount)
    ReDim seqs(1 To rowCount)
    Set seen = CreateObject("Scripting.Dictionary")

    For i = 1 To rowCount
        seqValue = ReadDesignTableValue(loEvents, values, i, "AppliedSeq")
        If Not IsNumeric(seqValue) Then
            errorMessage = "AppliedSeq is required at event row " & CStr(i) & "."
            Exit Function
        End If
        seqs(i) = CLng(seqValue)
        If seqs(i) <= 0 Then
            errorMessage = "AppliedSeq must be positive at event row " & CStr(i) & "."
            Exit Function
        End If
        If seen.Exists(CStr(seqs(i))) Then
            errorMessage = "Duplicate AppliedSeq " & CStr(seqs(i)) & "."
            Exit Function
        End If
        seen.Add CStr(seqs(i)), True
        order(i) = i
    Next i

    For i = 1 To rowCount - 1
        For j = i + 1 To rowCount
            If seqs(order(j)) < seqs(order(i)) Then
                swapValue = order(i)
                order(i) = order(j)
                order(j) = swapValue
            End If
        Next j
    Next i
    DesignEventReplayOrder = order
End Function

Private Function ValidateLifecycleTransition(ByVal wb As Workbook, ByVal eventType As String, _
                                             ByVal designId As String, ByVal designVersion As String, _
                                             ByVal payloadJson As String, ByRef errorCode As String, _
                                             ByRef errorMessage As String) As Boolean
    Dim currentStatus As String
    Dim payload As Collection

    currentStatus = CurrentDesignStatus(wb, designId, designVersion)
    Select Case eventType
        Case DESIGN_EVENT_CREATE
            If currentStatus <> "" Then
                errorCode = "DESIGN_VERSION_EXISTS"
                errorMessage = "DesignId and DesignVersion are immutable once created."
                Exit Function
            End If
            Set payload = ParseDesignPayload(payloadJson, errorMessage)
            If payload Is Nothing Then
                errorCode = "INVALID_DESIGN_PAYLOAD"
                If errorMessage = "" Then errorMessage = "DESIGN_CREATE requires at least one payload row."
                Exit Function
            End If
            If payload.Count = 0 Then
                errorCode = "INVALID_DESIGN_PAYLOAD"
                errorMessage = "DESIGN_CREATE requires at least one payload row."
                Exit Function
            End If
            If PayloadText(payload(1), "DesignType") = "" Or PayloadText(payload(1), "DesignName") = "" Then
                errorCode = "INVALID_DESIGN_PAYLOAD"
                errorMessage = "DESIGN_CREATE payload requires DesignType and DesignName."
                Exit Function
            End If
        Case DESIGN_EVENT_RELEASE
            If StrComp(currentStatus, "DRAFT", vbTextCompare) <> 0 Then
                errorCode = "INVALID_DESIGN_TRANSITION"
                errorMessage = "Only a DRAFT design version can be released."
                Exit Function
            End If
        Case DESIGN_EVENT_OBSOLETE
            If currentStatus = "" Or StrComp(currentStatus, "OBSOLETE", vbTextCompare) = 0 Then
                errorCode = "INVALID_DESIGN_TRANSITION"
                errorMessage = "Only an existing, non-obsolete design version can be obsoleted."
                Exit Function
            End If
        Case PROCESS_EVENT_SAVE
            If CurrentReusableDefinitionStatus(wb, "tblProcesses", "ProcessId", _
                    "ProcessVersion", designId, designVersion) <> "" Then
                errorCode = "PROCESS_VERSION_EXISTS"
                errorMessage = "ProcessId and ProcessVersion are immutable once saved."
                Exit Function
            End If
            Set payload = ParseDesignPayload(payloadJson, errorMessage)
            If Not ValidateProcessSavePayload(payload, errorCode, errorMessage) Then Exit Function
        Case PROCESS_EVENT_RELEASE
            If StrComp(CurrentReusableDefinitionStatus(wb, "tblProcesses", _
                    "ProcessId", "ProcessVersion", designId, designVersion), _
                    "DRAFT", vbTextCompare) <> 0 Then
                errorCode = "INVALID_PROCESS_TRANSITION"
                errorMessage = "Only a DRAFT Process version can be released."
                Exit Function
            End If
        Case PROCESS_EVENT_OBSOLETE
            currentStatus = CurrentReusableDefinitionStatus(wb, "tblProcesses", _
                "ProcessId", "ProcessVersion", designId, designVersion)
            If currentStatus = "" Or StrComp(currentStatus, "OBSOLETE", vbTextCompare) = 0 Then
                errorCode = "INVALID_PROCESS_TRANSITION"
                errorMessage = "Only an existing, non-obsolete Process version can be obsoleted."
                Exit Function
            End If
            If HasReleasedRecipeDependency(wb, designId, designVersion) Then
                errorCode = "PROCESS_HAS_RELEASED_RECIPE_DEPENDENCY"
                errorMessage = "A Process version referenced by a released Recipe cannot be obsoleted."
                Exit Function
            End If
        Case RECIPE_EVENT_SAVE
            If CurrentReusableDefinitionStatus(wb, "tblRecipes", "RecipeId", _
                    "RecipeVersion", designId, designVersion) <> "" Then
                errorCode = "RECIPE_VERSION_EXISTS"
                errorMessage = "RecipeId and RecipeVersion are immutable once saved."
                Exit Function
            End If
            Set payload = ParseDesignPayload(payloadJson, errorMessage)
            If payload Is Nothing Then
                errorCode = "INVALID_RECIPE_PAYLOAD"
                If errorMessage = "" Then errorMessage = "RECIPE_SAVE requires a payload."
                Exit Function
            End If
            If RecipePayloadHasCycle(payload) Then
                errorCode = "RECIPE_CYCLE"
                errorMessage = "Recipe Process connections contain a circular dependency."
                Exit Function
            End If
            If Not ValidateRecipeSavePayload(payload, errorCode, errorMessage) Then Exit Function
        Case RECIPE_EVENT_RELEASE
            If StrComp(CurrentReusableDefinitionStatus(wb, "tblRecipes", _
                    "RecipeId", "RecipeVersion", designId, designVersion), _
                    "DRAFT", vbTextCompare) <> 0 Then
                errorCode = "INVALID_RECIPE_TRANSITION"
                errorMessage = "Only a DRAFT Recipe version can be released."
                Exit Function
            End If
            If Not ValidateRecipeReleaseContract(wb, designId, designVersion, _
                    errorCode, errorMessage) Then Exit Function
        Case RECIPE_EVENT_OBSOLETE
            currentStatus = CurrentReusableDefinitionStatus(wb, "tblRecipes", _
                "RecipeId", "RecipeVersion", designId, designVersion)
            If currentStatus = "" Or StrComp(currentStatus, "OBSOLETE", vbTextCompare) = 0 Then
                errorCode = "INVALID_RECIPE_TRANSITION"
                errorMessage = "Only an existing, non-obsolete Recipe version can be obsoleted."
                Exit Function
            End If
        Case Else
            errorCode = "INVALID_DESIGN_EVENT_TYPE"
            errorMessage = "Unsupported design event type: " & eventType
            Exit Function
    End Select
    ValidateLifecycleTransition = True
End Function

Private Function ValidateProcessSavePayload(ByVal payload As Collection, _
                                            ByRef errorCode As String, _
                                            ByRef errorMessage As String) As Boolean
    Dim item As Variant
    Dim outputCount As Long

    If payload Is Nothing Or payload.Count = 0 Then
        errorCode = "INVALID_PROCESS_PAYLOAD"
        errorMessage = "PROCESS_SAVE requires a Process payload."
        Exit Function
    End If
    If StrComp(PayloadText(payload(1), "RecordType"), "PROCESS", vbTextCompare) <> 0 _
       Or PayloadText(payload(1), "ProcessName") = "" Then
        errorCode = "INVALID_PROCESS_PAYLOAD"
        errorMessage = "PROCESS_SAVE requires a PROCESS header with ProcessName."
        Exit Function
    End If
    For Each item In payload
        If StrComp(PayloadText(item, "RecordType"), "OUTPUT", vbTextCompare) = 0 Then
            outputCount = outputCount + 1
            If PayloadText(item, "OutputId") = "" _
               Or PayloadText(item, "OutputName") = "" _
               Or PayloadText(item, "ITEM_CODE") = "" _
               Or PayloadText(item, "UOM") = "" Then
                errorCode = "INVALID_PROCESS_OUTPUT"
                errorMessage = "Each Process output requires identity, item code, and UOM."
                Exit Function
            End If
            If Not PayloadPositiveNumber(item, "Qty") _
               And Not PayloadPositiveNumber(item, "Percent") Then
                errorCode = "INVALID_PROCESS_OUTPUT"
                errorMessage = "Each Process output requires a positive quantity or percentage."
                Exit Function
            End If
        End If
    Next item
    If outputCount = 0 Then
        errorCode = "PROCESS_OUTPUT_REQUIRED"
        errorMessage = "Every Process must declare at least one output."
        Exit Function
    End If
    ValidateProcessSavePayload = True
End Function

Private Function ValidateRecipeSavePayload(ByVal payload As Collection, _
                                           ByRef errorCode As String, _
                                           ByRef errorMessage As String) As Boolean
    Dim item As Variant
    Dim nodeCount As Long

    If payload.Count = 0 _
       Or StrComp(PayloadText(payload(1), "RecordType"), "RECIPE", vbTextCompare) <> 0 _
       Or PayloadText(payload(1), "RecipeName") = "" Then
        errorCode = "INVALID_RECIPE_PAYLOAD"
        errorMessage = "RECIPE_SAVE requires a RECIPE header with RecipeName."
        Exit Function
    End If
    For Each item In payload
        If StrComp(PayloadText(item, "RecordType"), "PROCESS_NODE", vbTextCompare) = 0 Then
            nodeCount = nodeCount + 1
        End If
    Next item
    If nodeCount = 0 Then
        errorCode = "RECIPE_PROCESS_REQUIRED"
        errorMessage = "Every Recipe must select at least one Process."
        Exit Function
    End If
    ValidateRecipeSavePayload = True
End Function

Private Function PayloadPositiveNumber(ByVal item As Object, ByVal key As String) As Boolean
    Dim valueIn As Variant
    valueIn = PayloadValue(item, key)
    If IsNumeric(valueIn) Then PayloadPositiveNumber = (CDbl(valueIn) > 0)
End Function

Private Function ReplayDesignEvent(ByVal loEvents As ListObject, ByVal values As Variant, _
                                   ByVal rowIndex As Long, ByVal loDesigns As ListObject, _
                                   ByVal loLines As ListObject, ByRef errorMessage As String) As Boolean
    Dim eventType As String
    Dim designId As String
    Dim designVersion As String
    Dim userId As String
    Dim eventId As String
    Dim appliedAt As Variant
    Dim payload As Collection
    Dim item As Variant
    Dim lr As ListRow
    Dim targetRow As Long

    eventType = UCase$(ReadDesignTableText(loEvents, values, rowIndex, "EventType"))
    designId = ReadDesignTableText(loEvents, values, rowIndex, "DesignId")
    designVersion = ReadDesignTableText(loEvents, values, rowIndex, "DesignVersion")
    userId = ReadDesignTableText(loEvents, values, rowIndex, "UserId")
    eventId = ReadDesignTableText(loEvents, values, rowIndex, "EventID")
    appliedAt = ReadDesignTableValue(loEvents, values, rowIndex, "AppliedAtUTC")

    Select Case eventType
        Case DESIGN_EVENT_CREATE
            Set payload = ParseDesignPayload(ReadDesignTableText(loEvents, values, rowIndex, "PayloadJson"), errorMessage)
            If payload Is Nothing Then Exit Function
            If payload.Count = 0 Then Exit Function
            Set lr = loDesigns.ListRows.Add
            SetDesignTableValue loDesigns, lr.Index, "DesignId", designId
            SetDesignTableValue loDesigns, lr.Index, "DesignVersion", designVersion
            SetDesignTableValue loDesigns, lr.Index, "DesignType", PayloadText(payload(1), "DesignType")
            SetDesignTableValue loDesigns, lr.Index, "DesignName", PayloadText(payload(1), "DesignName")
            SetDesignTableValue loDesigns, lr.Index, "Description", PayloadText(payload(1), "Description")
            SetDesignTableValue loDesigns, lr.Index, "Status", "DRAFT"
            SetDesignTableValue loDesigns, lr.Index, "EffectiveFromUTC", PayloadText(payload(1), "EffectiveFromUTC")
            SetDesignTableValue loDesigns, lr.Index, "EffectiveToUTC", PayloadText(payload(1), "EffectiveToUTC")
            SetDesignTableValue loDesigns, lr.Index, "CreatedAtUTC", ReadDesignTableValue(loEvents, values, rowIndex, "OccurredAtUTC")
            SetDesignTableValue loDesigns, lr.Index, "CreatedByUserId", userId
            SetDesignTableValue loDesigns, lr.Index, "SourceEventID", eventId
            For Each item In payload
                If PayloadHasLine(item) Then AddProjectedDesignLine loLines, designId, designVersion, item
            Next item
        Case DESIGN_EVENT_RELEASE
            targetRow = FindProjectedDesignRow(loDesigns, designId, designVersion)
            If targetRow = 0 Then
                errorMessage = "Release target not found."
                Exit Function
            End If
            SetDesignTableValue loDesigns, targetRow, "Status", "RELEASED"
            SetDesignTableValue loDesigns, targetRow, "ReleasedAtUTC", appliedAt
            SetDesignTableValue loDesigns, targetRow, "ReleasedByUserId", userId
        Case DESIGN_EVENT_OBSOLETE
            targetRow = FindProjectedDesignRow(loDesigns, designId, designVersion)
            If targetRow = 0 Then
                errorMessage = "Obsolete target not found."
                Exit Function
            End If
            SetDesignTableValue loDesigns, targetRow, "Status", "OBSOLETE"
            SetDesignTableValue loDesigns, targetRow, "ObsoletedAtUTC", appliedAt
            SetDesignTableValue loDesigns, targetRow, "ObsoletedByUserId", userId
        Case Else
            errorMessage = "Unsupported event type " & eventType
            Exit Function
    End Select
    ReplayDesignEvent = True
End Function

Private Function ReplayReusableProductionEvent(ByVal loEvents As ListObject, _
                                               ByVal values As Variant, _
                                               ByVal rowIndex As Long, _
                                               ByVal wb As Workbook, _
                                               ByRef errorMessage As String) As Boolean
    Dim eventType As String
    Dim definitionId As String
    Dim definitionVersion As String
    Dim userId As String
    Dim eventId As String
    Dim appliedAt As Variant
    Dim payload As Collection
    Dim item As Variant
    Dim loHeader As ListObject
    Dim targetRow As Long
    Dim lr As ListRow
    Dim alternativeOrdinal As Long

    eventType = UCase$(ReadDesignTableText(loEvents, values, rowIndex, "EventType"))
    definitionId = ReadDesignTableText(loEvents, values, rowIndex, "DesignId")
    definitionVersion = ReadDesignTableText(loEvents, values, rowIndex, "DesignVersion")
    userId = ReadDesignTableText(loEvents, values, rowIndex, "UserId")
    eventId = ReadDesignTableText(loEvents, values, rowIndex, "EventID")
    appliedAt = ReadDesignTableValue(loEvents, values, rowIndex, "AppliedAtUTC")

    Select Case eventType
        Case PROCESS_EVENT_SAVE
            Set payload = ParseDesignPayload( _
                ReadDesignTableText(loEvents, values, rowIndex, "PayloadJson"), errorMessage)
            If payload Is Nothing Or payload.Count = 0 Then Exit Function
            Set loHeader = FindDesignsApplyTable(wb, "tblProcesses")
            Set lr = loHeader.ListRows.Add
            SetDesignTableValue loHeader, lr.Index, "ProcessId", definitionId
            SetDesignTableValue loHeader, lr.Index, "ProcessVersion", definitionVersion
            SetDesignTableValue loHeader, lr.Index, "ProcessName", PayloadText(payload(1), "ProcessName")
            SetDesignTableValue loHeader, lr.Index, "Description", PayloadText(payload(1), "Description")
            SetDesignTableValue loHeader, lr.Index, "Status", "DRAFT"
            SetDesignTableValue loHeader, lr.Index, "CreatedAtUTC", _
                ReadDesignTableValue(loEvents, values, rowIndex, "OccurredAtUTC")
            SetDesignTableValue loHeader, lr.Index, "CreatedByUserId", userId
            SetDesignTableValue loHeader, lr.Index, "SourceEventID", eventId
            For Each item In payload
                ProjectProcessPayloadItem wb, definitionId, definitionVersion, _
                    item, alternativeOrdinal
            Next item
        Case PROCESS_EVENT_RELEASE, PROCESS_EVENT_OBSOLETE
            Set loHeader = FindDesignsApplyTable(wb, "tblProcesses")
            targetRow = FindReusableDefinitionRow(loHeader, "ProcessId", _
                "ProcessVersion", definitionId, definitionVersion)
            If targetRow = 0 Then
                errorMessage = "Process lifecycle target not found."
                Exit Function
            End If
            If eventType = PROCESS_EVENT_RELEASE Then
                SetDesignTableValue loHeader, targetRow, "Status", "RELEASED"
                SetDesignTableValue loHeader, targetRow, "ReleasedAtUTC", appliedAt
                SetDesignTableValue loHeader, targetRow, "ReleasedByUserId", userId
            Else
                SetDesignTableValue loHeader, targetRow, "Status", "OBSOLETE"
                SetDesignTableValue loHeader, targetRow, "ObsoletedAtUTC", appliedAt
                SetDesignTableValue loHeader, targetRow, "ObsoletedByUserId", userId
            End If
        Case RECIPE_EVENT_SAVE
            Set payload = ParseDesignPayload( _
                ReadDesignTableText(loEvents, values, rowIndex, "PayloadJson"), errorMessage)
            If payload Is Nothing Or payload.Count = 0 Then Exit Function
            Set loHeader = FindDesignsApplyTable(wb, "tblRecipes")
            Set lr = loHeader.ListRows.Add
            SetDesignTableValue loHeader, lr.Index, "RecipeId", definitionId
            SetDesignTableValue loHeader, lr.Index, "RecipeVersion", definitionVersion
            SetDesignTableValue loHeader, lr.Index, "RecipeName", PayloadText(payload(1), "RecipeName")
            SetDesignTableValue loHeader, lr.Index, "Description", PayloadText(payload(1), "Description")
            SetDesignTableValue loHeader, lr.Index, "Status", "DRAFT"
            SetDesignTableValue loHeader, lr.Index, "CreatedAtUTC", _
                ReadDesignTableValue(loEvents, values, rowIndex, "OccurredAtUTC")
            SetDesignTableValue loHeader, lr.Index, "CreatedByUserId", userId
            SetDesignTableValue loHeader, lr.Index, "SourceEventID", eventId
            For Each item In payload
                ProjectRecipePayloadItem wb, definitionId, definitionVersion, item
            Next item
        Case RECIPE_EVENT_RELEASE, RECIPE_EVENT_OBSOLETE
            Set loHeader = FindDesignsApplyTable(wb, "tblRecipes")
            targetRow = FindReusableDefinitionRow(loHeader, "RecipeId", _
                "RecipeVersion", definitionId, definitionVersion)
            If targetRow = 0 Then
                errorMessage = "Recipe lifecycle target not found."
                Exit Function
            End If
            If eventType = RECIPE_EVENT_RELEASE Then
                SetDesignTableValue loHeader, targetRow, "Status", "RELEASED"
                SetDesignTableValue loHeader, targetRow, "ReleasedAtUTC", appliedAt
                SetDesignTableValue loHeader, targetRow, "ReleasedByUserId", userId
            Else
                SetDesignTableValue loHeader, targetRow, "Status", "OBSOLETE"
                SetDesignTableValue loHeader, targetRow, "ObsoletedAtUTC", appliedAt
                SetDesignTableValue loHeader, targetRow, "ObsoletedByUserId", userId
            End If
        Case Else
            errorMessage = "Unsupported reusable Production event type " & eventType
            Exit Function
    End Select
    ReplayReusableProductionEvent = True
End Function

Private Sub ProjectProcessPayloadItem(ByVal wb As Workbook, ByVal processId As String, _
                                      ByVal processVersion As String, ByVal item As Object, _
                                      ByRef alternativeOrdinal As Long)
    Dim recordType As String
    Dim lo As ListObject
    Dim lr As ListRow

    recordType = UCase$(PayloadText(item, "RecordType"))
    Select Case recordType
        Case "REQUIREMENT"
            Set lo = FindDesignsApplyTable(wb, "tblProcessRequirements")
            Set lr = lo.ListRows.Add
            SetDesignTableValue lo, lr.Index, "ProcessId", processId
            SetDesignTableValue lo, lr.Index, "ProcessVersion", processVersion
            SetDesignTableValue lo, lr.Index, "RequirementId", PayloadText(item, "RequirementId")
            SetDesignTableValue lo, lr.Index, "RequirementName", PayloadText(item, "RequirementName")
            SetDesignTableValue lo, lr.Index, "Qty", PayloadValue(item, "Qty")
            SetDesignTableValue lo, lr.Index, "Percent", PayloadValue(item, "Percent")
            SetDesignTableValue lo, lr.Index, "YieldBasis", PayloadText(item, "YieldBasis")
            SetDesignTableValue lo, lr.Index, "UOM", PayloadText(item, "UOM")
        Case "ALTERNATIVE"
            alternativeOrdinal = alternativeOrdinal + 1
            Set lo = FindDesignsApplyTable(wb, "tblProcessIngredientAlternatives")
            Set lr = lo.ListRows.Add
            SetDesignTableValue lo, lr.Index, "ProcessId", processId
            SetDesignTableValue lo, lr.Index, "ProcessVersion", processVersion
            SetDesignTableValue lo, lr.Index, "RequirementId", PayloadText(item, "RequirementId")
            SetDesignTableValue lo, lr.Index, "AlternativeOrdinal", alternativeOrdinal
            SetDesignTableValue lo, lr.Index, "ITEM_CODE", PayloadText(item, "ITEM_CODE")
        Case "OUTPUT"
            Set lo = FindDesignsApplyTable(wb, "tblProcessOutputs")
            Set lr = lo.ListRows.Add
            SetDesignTableValue lo, lr.Index, "ProcessId", processId
            SetDesignTableValue lo, lr.Index, "ProcessVersion", processVersion
            SetDesignTableValue lo, lr.Index, "OutputId", PayloadText(item, "OutputId")
            SetDesignTableValue lo, lr.Index, "OutputName", PayloadText(item, "OutputName")
            SetDesignTableValue lo, lr.Index, "ITEM_CODE", PayloadText(item, "ITEM_CODE")
            SetDesignTableValue lo, lr.Index, "ComponentDesignId", PayloadText(item, "ComponentDesignId")
            SetDesignTableValue lo, lr.Index, "ComponentDesignVersion", PayloadText(item, "ComponentDesignVersion")
            SetDesignTableValue lo, lr.Index, "Qty", PayloadValue(item, "Qty")
            SetDesignTableValue lo, lr.Index, "Percent", PayloadValue(item, "Percent")
            SetDesignTableValue lo, lr.Index, "YieldBasis", PayloadText(item, "YieldBasis")
            SetDesignTableValue lo, lr.Index, "UOM", PayloadText(item, "UOM")
        Case "INSTRUCTION"
            Set lo = FindDesignsApplyTable(wb, "tblProcessInstructions")
            Set lr = lo.ListRows.Add
            SetDesignTableValue lo, lr.Index, "ProcessId", processId
            SetDesignTableValue lo, lr.Index, "ProcessVersion", processVersion
            SetDesignTableValue lo, lr.Index, "InstructionOrdinal", PayloadValue(item, "InstructionOrdinal")
            SetDesignTableValue lo, lr.Index, "Instruction", PayloadText(item, "Instruction")
    End Select
End Sub

Private Sub ProjectRecipePayloadItem(ByVal wb As Workbook, ByVal recipeId As String, _
                                     ByVal recipeVersion As String, ByVal item As Object)
    Dim recordType As String
    Dim lo As ListObject
    Dim lr As ListRow

    recordType = UCase$(PayloadText(item, "RecordType"))
    If recordType = "PROCESS_NODE" Then
        Set lo = FindDesignsApplyTable(wb, "tblRecipeProcesses")
        Set lr = lo.ListRows.Add
        SetDesignTableValue lo, lr.Index, "RecipeId", recipeId
        SetDesignTableValue lo, lr.Index, "RecipeVersion", recipeVersion
        SetDesignTableValue lo, lr.Index, "ProcessNodeId", PayloadText(item, "ProcessNodeId")
        SetDesignTableValue lo, lr.Index, "ProcessId", PayloadText(item, "ProcessId")
        SetDesignTableValue lo, lr.Index, "ProcessVersion", PayloadText(item, "ProcessVersion")
        SetDesignTableValue lo, lr.Index, "ExecutionOrdinal", PayloadValue(item, "ExecutionOrdinal")
    ElseIf recordType = "CONNECTION" Then
        Set lo = FindDesignsApplyTable(wb, "tblRecipeConnections")
        Set lr = lo.ListRows.Add
        SetDesignTableValue lo, lr.Index, "RecipeId", recipeId
        SetDesignTableValue lo, lr.Index, "RecipeVersion", recipeVersion
        SetDesignTableValue lo, lr.Index, "FromProcessNodeId", PayloadText(item, "FromProcessNodeId")
        SetDesignTableValue lo, lr.Index, "FromOutputId", PayloadText(item, "FromOutputId")
        SetDesignTableValue lo, lr.Index, "ToProcessNodeId", PayloadText(item, "ToProcessNodeId")
        SetDesignTableValue lo, lr.Index, "ToRequirementId", PayloadText(item, "ToRequirementId")
        SetDesignTableValue lo, lr.Index, "Qty", PayloadValue(item, "Qty")
        SetDesignTableValue lo, lr.Index, "Percent", PayloadValue(item, "Percent")
        SetDesignTableValue lo, lr.Index, "UOM", PayloadText(item, "UOM")
    End If
End Sub

Private Sub AddProjectedDesignLine(ByVal lo As ListObject, ByVal designId As String, _
                                   ByVal designVersion As String, ByVal item As Object)
    Dim lr As ListRow
    Set lr = lo.ListRows.Add
    SetDesignTableValue lo, lr.Index, "DesignId", designId
    SetDesignTableValue lo, lr.Index, "DesignVersion", designVersion
    SetDesignTableValue lo, lr.Index, "LineNo", PayloadValue(item, "LineNo")
    SetDesignTableValue lo, lr.Index, "Process", PayloadText(item, "Process")
    SetDesignTableValue lo, lr.Index, "IOType", PayloadText(item, "IOType")
    SetDesignTableValue lo, lr.Index, "ComponentSKU", PayloadText(item, "ComponentSKU")
    SetDesignTableValue lo, lr.Index, "ComponentDesignId", PayloadText(item, "ComponentDesignId")
    SetDesignTableValue lo, lr.Index, "ComponentDesignVersion", PayloadText(item, "ComponentDesignVersion")
    SetDesignTableValue lo, lr.Index, "Qty", PayloadValue(item, "Qty")
    SetDesignTableValue lo, lr.Index, "UOM", PayloadText(item, "UOM")
    SetDesignTableValue lo, lr.Index, "Percent", PayloadValue(item, "Percent")
    SetDesignTableValue lo, lr.Index, "Instruction", PayloadText(item, "Instruction")
End Sub

Private Function PayloadHasLine(ByVal item As Object) As Boolean
    PayloadHasLine = (PayloadText(item, "ComponentSKU") <> "" _
                      Or PayloadText(item, "ComponentDesignId") <> "" _
                      Or PayloadText(item, "IOType") <> "")
End Function

Private Function CurrentDesignStatus(ByVal wb As Workbook, ByVal designId As String, _
                                     ByVal designVersion As String) As String
    Dim lo As ListObject
    Dim values As Variant
    Dim r As Long
    Set lo = FindDesignsApplyTable(wb, "tblDesigns")
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    values = lo.DataBodyRange.Value
    For r = 1 To UBound(values, 1)
        If StrComp(ReadDesignTableText(lo, values, r, "DesignId"), designId, vbTextCompare) = 0 _
           And StrComp(ReadDesignTableText(lo, values, r, "DesignVersion"), designVersion, vbTextCompare) = 0 Then
            CurrentDesignStatus = ReadDesignTableText(lo, values, r, "Status")
            Exit Function
        End If
    Next r
End Function

Private Function CurrentReusableDefinitionStatus(ByVal wb As Workbook, _
                                                 ByVal tableName As String, _
                                                 ByVal idColumn As String, _
                                                 ByVal versionColumn As String, _
                                                 ByVal definitionId As String, _
                                                 ByVal definitionVersion As String) As String
    Dim lo As ListObject
    Dim rowIndex As Long

    Set lo = FindDesignsApplyTable(wb, tableName)
    rowIndex = FindReusableDefinitionRow(lo, idColumn, versionColumn, _
        definitionId, definitionVersion)
    If rowIndex > 0 Then
        CurrentReusableDefinitionStatus = Trim$(CStr( _
            lo.DataBodyRange.Cells(rowIndex, lo.ListColumns("Status").Index).Value2))
    End If
End Function

Private Function FindReusableDefinitionRow(ByVal lo As ListObject, _
                                           ByVal idColumn As String, _
                                           ByVal versionColumn As String, _
                                           ByVal definitionId As String, _
                                           ByVal definitionVersion As String) As Long
    Dim values As Variant
    Dim rowIndex As Long

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    values = lo.DataBodyRange.Value2
    For rowIndex = 1 To UBound(values, 1)
        If StrComp(ReadDesignTableText(lo, values, rowIndex, idColumn), _
                   definitionId, vbTextCompare) = 0 _
           And StrComp(ReadDesignTableText(lo, values, rowIndex, versionColumn), _
                       definitionVersion, vbTextCompare) = 0 Then
            FindReusableDefinitionRow = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

Private Function RecipePayloadHasCycle(ByVal payload As Collection) As Boolean
    Dim nodes As Object
    Dim inDegree As Object
    Dim adjacency As Object
    Dim processed As Object
    Dim item As Variant
    Dim nodeId As String
    Dim fromNode As String
    Dim toNode As String
    Dim edges As Collection
    Dim nodeKey As Variant
    Dim edgeTarget As Variant
    Dim foundNode As Boolean

    Set nodes = CreateObject("Scripting.Dictionary")
    Set inDegree = CreateObject("Scripting.Dictionary")
    Set adjacency = CreateObject("Scripting.Dictionary")
    Set processed = CreateObject("Scripting.Dictionary")
    For Each item In payload
        If StrComp(PayloadText(item, "RecordType"), "PROCESS_NODE", vbTextCompare) = 0 Then
            nodeId = PayloadText(item, "ProcessNodeId")
            If nodeId <> "" And Not nodes.Exists(nodeId) Then
                nodes.Add nodeId, True
                inDegree.Add nodeId, 0
                Set edges = New Collection
                adjacency.Add nodeId, edges
            End If
        End If
    Next item
    For Each item In payload
        If StrComp(PayloadText(item, "RecordType"), "CONNECTION", vbTextCompare) = 0 Then
            fromNode = PayloadText(item, "FromProcessNodeId")
            toNode = PayloadText(item, "ToProcessNodeId")
            If nodes.Exists(fromNode) And nodes.Exists(toNode) Then
                Set edges = adjacency(fromNode)
                edges.Add toNode
                inDegree(toNode) = CLng(inDegree(toNode)) + 1
            End If
        End If
    Next item

    Do
        foundNode = False
        For Each nodeKey In nodes.Keys
            If Not processed.Exists(CStr(nodeKey)) _
               And CLng(inDegree(CStr(nodeKey))) = 0 Then
                processed.Add CStr(nodeKey), True
                Set edges = adjacency(CStr(nodeKey))
                For Each edgeTarget In edges
                    inDegree(CStr(edgeTarget)) = CLng(inDegree(CStr(edgeTarget))) - 1
                Next edgeTarget
                foundNode = True
            End If
        Next nodeKey
    Loop While foundNode
    RecipePayloadHasCycle = (processed.Count < nodes.Count)
End Function

Private Function ValidateRecipeReleaseContract(ByVal wb As Workbook, _
                                               ByVal recipeId As String, _
                                               ByVal recipeVersion As String, _
                                               ByRef errorCode As String, _
                                               ByRef errorMessage As String) As Boolean
    Dim loNodes As ListObject
    Dim loConnections As ListObject
    Dim loRequirements As ListObject
    Dim nodeValues As Variant
    Dim connectionValues As Variant
    Dim requirementValues As Variant
    Dim nodes As Object
    Dim ordinals As Object
    Dim connectedRequirements As Object
    Dim routedQty As Object
    Dim routedPercent As Object
    Dim nodeInfo As Object
    Dim sourceInfo As Object
    Dim targetInfo As Object
    Dim r As Long
    Dim nodeKey As Variant
    Dim processNodeId As String
    Dim processId As String
    Dim processVersion As String
    Dim ordinalValue As Variant
    Dim ordinalKey As String
    Dim fromNodeKey As String
    Dim toNodeKey As String
    Dim outputId As String
    Dim requirementId As String
    Dim routeKey As String
    Dim targetKey As String
    Dim connectionUom As String
    Dim outputItemCode As String
    Dim outputUom As String
    Dim requirementUom As String
    Dim outputQty As Double
    Dim outputPercent As Double
    Dim requirementQty As Double
    Dim requirementPercent As Double
    Dim connectionQty As Double
    Dim connectionPercent As Double
    Dim currentRouted As Double
    Dim alternativeCount As Long
    Dim itemAccepted As Boolean

    Set loNodes = FindDesignsApplyTable(wb, "tblRecipeProcesses")
    Set loConnections = FindDesignsApplyTable(wb, "tblRecipeConnections")
    Set loRequirements = FindDesignsApplyTable(wb, "tblProcessRequirements")
    Set nodes = CreateObject("Scripting.Dictionary")
    Set ordinals = CreateObject("Scripting.Dictionary")
    Set connectedRequirements = CreateObject("Scripting.Dictionary")
    Set routedQty = CreateObject("Scripting.Dictionary")
    Set routedPercent = CreateObject("Scripting.Dictionary")
    nodes.CompareMode = vbTextCompare
    ordinals.CompareMode = vbTextCompare
    connectedRequirements.CompareMode = vbTextCompare
    routedQty.CompareMode = vbTextCompare
    routedPercent.CompareMode = vbTextCompare

    If loNodes Is Nothing Or loNodes.DataBodyRange Is Nothing Then
        errorCode = "RECIPE_PROCESS_REQUIRED"
        errorMessage = "A Recipe release requires at least one Process node."
        Exit Function
    End If
    nodeValues = loNodes.DataBodyRange.Value2
    For r = 1 To UBound(nodeValues, 1)
        If ReusableProjectionRowMatches(loNodes, nodeValues, r, _
                "RecipeId", "RecipeVersion", recipeId, recipeVersion) Then
            processNodeId = ReadDesignTableText(loNodes, nodeValues, r, "ProcessNodeId")
            processId = ReadDesignTableText(loNodes, nodeValues, r, "ProcessId")
            processVersion = ReadDesignTableText(loNodes, nodeValues, r, "ProcessVersion")
            ordinalValue = ReadDesignTableValue(loNodes, nodeValues, r, "ExecutionOrdinal")
            If processNodeId = "" Or processId = "" Or processVersion = "" _
               Or Not IsNumeric(ordinalValue) Or CLng(ordinalValue) <= 0 Then
                errorCode = "RECIPE_PROCESS_INVALID"
                errorMessage = "Each Recipe Process node requires identity, version, and positive execution order."
                Exit Function
            End If
            If nodes.Exists(processNodeId) Then
                errorCode = "RECIPE_PROCESS_INVALID"
                errorMessage = "Process node identities must be unique within a Recipe version."
                Exit Function
            End If
            ordinalKey = CStr(CLng(ordinalValue))
            If ordinals.Exists(ordinalKey) Then
                errorCode = "RECIPE_EXECUTION_ORDER"
                errorMessage = "Recipe execution order values must be unique."
                Exit Function
            End If
            If StrComp(CurrentReusableDefinitionStatus(wb, "tblProcesses", _
                    "ProcessId", "ProcessVersion", processId, processVersion), _
                    "RELEASED", vbTextCompare) <> 0 Then
                errorCode = "RECIPE_PROCESS_NOT_RELEASED"
                errorMessage = "Every Recipe node must pin an existing released Process version."
                Exit Function
            End If
            Set nodeInfo = CreateObject("Scripting.Dictionary")
            nodeInfo("ProcessId") = processId
            nodeInfo("ProcessVersion") = processVersion
            nodeInfo("ExecutionOrdinal") = CLng(ordinalValue)
            nodes.Add processNodeId, nodeInfo
            ordinals.Add ordinalKey, True
        End If
    Next r
    If nodes.Count = 0 Then
        errorCode = "RECIPE_PROCESS_REQUIRED"
        errorMessage = "A Recipe release requires at least one Process node."
        Exit Function
    End If

    If Not loConnections Is Nothing And Not loConnections.DataBodyRange Is Nothing Then
        connectionValues = loConnections.DataBodyRange.Value2
        For r = 1 To UBound(connectionValues, 1)
            If ReusableProjectionRowMatches(loConnections, connectionValues, r, _
                    "RecipeId", "RecipeVersion", recipeId, recipeVersion) Then
                fromNodeKey = ReadDesignTableText(loConnections, connectionValues, r, "FromProcessNodeId")
                toNodeKey = ReadDesignTableText(loConnections, connectionValues, r, "ToProcessNodeId")
                outputId = ReadDesignTableText(loConnections, connectionValues, r, "FromOutputId")
                requirementId = ReadDesignTableText(loConnections, connectionValues, r, "ToRequirementId")
                connectionUom = ReadDesignTableText(loConnections, connectionValues, r, "UOM")
                If Not nodes.Exists(fromNodeKey) Or Not nodes.Exists(toNodeKey) _
                   Or outputId = "" Or requirementId = "" Then
                    errorCode = "RECIPE_CONNECTION_INVALID"
                    errorMessage = "Every Recipe connection must reference selected Process nodes, an output, and a requirement."
                    Exit Function
                End If
                Set sourceInfo = nodes(fromNodeKey)
                Set targetInfo = nodes(toNodeKey)
                If CLng(sourceInfo("ExecutionOrdinal")) >= CLng(targetInfo("ExecutionOrdinal")) Then
                    errorCode = "RECIPE_EXECUTION_ORDER"
                    errorMessage = "Recipe execution order must place every source Process before its downstream Process."
                    Exit Function
                End If
                If Not TryGetProcessOutputForRecipe(wb, CStr(sourceInfo("ProcessId")), _
                        CStr(sourceInfo("ProcessVersion")), outputId, outputItemCode, _
                        outputUom, outputQty, outputPercent) Then
                    errorCode = "RECIPE_CONNECTION_INVALID"
                    errorMessage = "A Recipe connection references an output that is not declared by its Process version."
                    Exit Function
                End If
                If Not TryGetProcessRequirementForRecipe(wb, CStr(targetInfo("ProcessId")), _
                        CStr(targetInfo("ProcessVersion")), requirementId, requirementUom, _
                        requirementQty, requirementPercent) Then
                    errorCode = "RECIPE_CONNECTION_INVALID"
                    errorMessage = "A Recipe connection references a requirement that is not declared by its Process version."
                    Exit Function
                End If
                If connectionUom = "" _
                   Or StrComp(connectionUom, outputUom, vbTextCompare) <> 0 _
                   Or StrComp(connectionUom, requirementUom, vbTextCompare) <> 0 Then
                    errorCode = "RECIPE_CONNECTION_INCOMPATIBLE"
                    errorMessage = "Connection UOM must match both the output and downstream requirement."
                    Exit Function
                End If
                GetRequirementAlternativeInfo wb, CStr(targetInfo("ProcessId")), _
                    CStr(targetInfo("ProcessVersion")), requirementId, outputItemCode, _
                    alternativeCount, itemAccepted
                If alternativeCount > 0 And Not itemAccepted Then
                    errorCode = "RECIPE_CONNECTION_INCOMPATIBLE"
                    errorMessage = "The output item is not an acceptable alternative for the downstream requirement."
                    Exit Function
                End If
                targetKey = UCase$(toNodeKey & "|" & requirementId)
                If connectedRequirements.Exists(targetKey) Then
                    errorCode = "RECIPE_REQUIREMENT_MULTIPLE_SOURCES"
                    errorMessage = "A Process requirement may have only one upstream output connection."
                    Exit Function
                End If
                connectedRequirements.Add targetKey, True
                routeKey = UCase$(fromNodeKey & "|" & outputId)
                If TryPositiveDouble(ReadDesignTableValue(loConnections, connectionValues, r, "Qty"), _
                        connectionQty) Then
                    If outputQty <= 0 Or (requirementQty > 0 _
                       And Abs(connectionQty - requirementQty) > 0.0000001) Then
                        errorCode = "RECIPE_CONNECTION_QUANTITY"
                        errorMessage = "Connection quantity must use and satisfy the output and requirement quantity basis."
                        Exit Function
                    End If
                    currentRouted = connectionQty
                    If routedQty.Exists(routeKey) Then currentRouted = currentRouted + CDbl(routedQty(routeKey))
                    If currentRouted - outputQty > 0.0000001 Then
                        errorCode = "RECIPE_OUTPUT_OVERALLOCATED"
                        errorMessage = "Routed connection quantity exceeds the Process output yield."
                        Exit Function
                    End If
                    routedQty(routeKey) = currentRouted
                ElseIf TryPositiveDouble(ReadDesignTableValue(loConnections, connectionValues, r, "Percent"), _
                        connectionPercent) Then
                    If outputPercent <= 0 Or (requirementPercent > 0 _
                       And Abs(connectionPercent - requirementPercent) > 0.0000001) Then
                        errorCode = "RECIPE_CONNECTION_QUANTITY"
                        errorMessage = "Connection percentage must use and satisfy the output and requirement percentage basis."
                        Exit Function
                    End If
                    currentRouted = connectionPercent
                    If routedPercent.Exists(routeKey) Then currentRouted = currentRouted + CDbl(routedPercent(routeKey))
                    If currentRouted - outputPercent > 0.0000001 Then
                        errorCode = "RECIPE_OUTPUT_OVERALLOCATED"
                        errorMessage = "Routed connection percentage exceeds the Process output yield."
                        Exit Function
                    End If
                    routedPercent(routeKey) = currentRouted
                Else
                    errorCode = "RECIPE_CONNECTION_QUANTITY"
                    errorMessage = "Every Recipe connection requires a positive quantity or percentage."
                    Exit Function
                End If
            End If
        Next r
    End If

    If Not loRequirements Is Nothing And Not loRequirements.DataBodyRange Is Nothing Then
        requirementValues = loRequirements.DataBodyRange.Value2
        For Each nodeKey In nodes.Keys
            Set nodeInfo = nodes(CStr(nodeKey))
            For r = 1 To UBound(requirementValues, 1)
                If ReusableProjectionRowMatches(loRequirements, requirementValues, r, _
                        "ProcessId", "ProcessVersion", CStr(nodeInfo("ProcessId")), _
                        CStr(nodeInfo("ProcessVersion"))) Then
                    requirementId = ReadDesignTableText(loRequirements, requirementValues, r, "RequirementId")
                    targetKey = UCase$(CStr(nodeKey) & "|" & requirementId)
                    GetRequirementAlternativeInfo wb, CStr(nodeInfo("ProcessId")), _
                        CStr(nodeInfo("ProcessVersion")), requirementId, "", _
                        alternativeCount, itemAccepted
                    If Not connectedRequirements.Exists(targetKey) And alternativeCount = 0 Then
                        errorCode = "RECIPE_UNRESOLVED_REQUIREMENT"
                        errorMessage = "Every Process requirement needs one upstream connection or an acceptable inventory alternative."
                        Exit Function
                    End If
                End If
            Next r
        Next nodeKey
    End If
    ValidateRecipeReleaseContract = True
End Function

Private Function HasReleasedRecipeDependency(ByVal wb As Workbook, _
                                             ByVal processId As String, _
                                             ByVal processVersion As String) As Boolean
    Dim loRecipes As ListObject
    Dim loNodes As ListObject
    Dim recipeValues As Variant
    Dim nodeValues As Variant
    Dim r As Long
    Dim n As Long
    Dim recipeId As String
    Dim recipeVersion As String

    Set loRecipes = FindDesignsApplyTable(wb, "tblRecipes")
    Set loNodes = FindDesignsApplyTable(wb, "tblRecipeProcesses")
    If loRecipes Is Nothing Or loNodes Is Nothing Then Exit Function
    If loRecipes.DataBodyRange Is Nothing Or loNodes.DataBodyRange Is Nothing Then Exit Function
    recipeValues = loRecipes.DataBodyRange.Value2
    nodeValues = loNodes.DataBodyRange.Value2
    For r = 1 To UBound(recipeValues, 1)
        If StrComp(ReadDesignTableText(loRecipes, recipeValues, r, "Status"), _
                "RELEASED", vbTextCompare) = 0 Then
            recipeId = ReadDesignTableText(loRecipes, recipeValues, r, "RecipeId")
            recipeVersion = ReadDesignTableText(loRecipes, recipeValues, r, "RecipeVersion")
            For n = 1 To UBound(nodeValues, 1)
                If ReusableProjectionRowMatches(loNodes, nodeValues, n, _
                        "RecipeId", "RecipeVersion", recipeId, recipeVersion) _
                   And StrComp(ReadDesignTableText(loNodes, nodeValues, n, "ProcessId"), _
                               processId, vbTextCompare) = 0 _
                   And StrComp(ReadDesignTableText(loNodes, nodeValues, n, "ProcessVersion"), _
                               processVersion, vbTextCompare) = 0 Then
                    HasReleasedRecipeDependency = True
                    Exit Function
                End If
            Next n
        End If
    Next r
End Function

Private Function ReusableProjectionRowMatches(ByVal lo As ListObject, _
                                              ByVal values As Variant, _
                                              ByVal rowIndex As Long, _
                                              ByVal idColumn As String, _
                                              ByVal versionColumn As String, _
                                              ByVal definitionId As String, _
                                              ByVal definitionVersion As String) As Boolean
    ReusableProjectionRowMatches = _
        (StrComp(ReadDesignTableText(lo, values, rowIndex, idColumn), _
                 definitionId, vbTextCompare) = 0 _
         And StrComp(ReadDesignTableText(lo, values, rowIndex, versionColumn), _
                     definitionVersion, vbTextCompare) = 0)
End Function

Private Function TryGetProcessOutputForRecipe(ByVal wb As Workbook, _
                                              ByVal processId As String, _
                                              ByVal processVersion As String, _
                                              ByVal outputId As String, _
                                              ByRef itemCode As String, _
                                              ByRef uom As String, _
                                              ByRef qty As Double, _
                                              ByRef percent As Double) As Boolean
    Dim lo As ListObject
    Dim values As Variant
    Dim r As Long
    Set lo = FindDesignsApplyTable(wb, "tblProcessOutputs")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    values = lo.DataBodyRange.Value2
    For r = 1 To UBound(values, 1)
        If ReusableProjectionRowMatches(lo, values, r, "ProcessId", _
                "ProcessVersion", processId, processVersion) _
           And StrComp(ReadDesignTableText(lo, values, r, "OutputId"), _
                       outputId, vbTextCompare) = 0 Then
            itemCode = ReadDesignTableText(lo, values, r, "ITEM_CODE")
            uom = ReadDesignTableText(lo, values, r, "UOM")
            Call TryPositiveDouble(ReadDesignTableValue(lo, values, r, "Qty"), qty)
            Call TryPositiveDouble(ReadDesignTableValue(lo, values, r, "Percent"), percent)
            TryGetProcessOutputForRecipe = True
            Exit Function
        End If
    Next r
End Function

Private Function TryGetProcessRequirementForRecipe(ByVal wb As Workbook, _
                                                   ByVal processId As String, _
                                                   ByVal processVersion As String, _
                                                   ByVal requirementId As String, _
                                                   ByRef uom As String, _
                                                   ByRef qty As Double, _
                                                   ByRef percent As Double) As Boolean
    Dim lo As ListObject
    Dim values As Variant
    Dim r As Long
    Set lo = FindDesignsApplyTable(wb, "tblProcessRequirements")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    values = lo.DataBodyRange.Value2
    For r = 1 To UBound(values, 1)
        If ReusableProjectionRowMatches(lo, values, r, "ProcessId", _
                "ProcessVersion", processId, processVersion) _
           And StrComp(ReadDesignTableText(lo, values, r, "RequirementId"), _
                       requirementId, vbTextCompare) = 0 Then
            uom = ReadDesignTableText(lo, values, r, "UOM")
            Call TryPositiveDouble(ReadDesignTableValue(lo, values, r, "Qty"), qty)
            Call TryPositiveDouble(ReadDesignTableValue(lo, values, r, "Percent"), percent)
            TryGetProcessRequirementForRecipe = True
            Exit Function
        End If
    Next r
End Function

Private Sub GetRequirementAlternativeInfo(ByVal wb As Workbook, _
                                          ByVal processId As String, _
                                          ByVal processVersion As String, _
                                          ByVal requirementId As String, _
                                          ByVal candidateItemCode As String, _
                                          ByRef alternativeCount As Long, _
                                          ByRef itemAccepted As Boolean)
    Dim lo As ListObject
    Dim values As Variant
    Dim r As Long
    Dim itemCode As String
    alternativeCount = 0
    itemAccepted = False
    Set lo = FindDesignsApplyTable(wb, "tblProcessIngredientAlternatives")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    values = lo.DataBodyRange.Value2
    For r = 1 To UBound(values, 1)
        If ReusableProjectionRowMatches(lo, values, r, "ProcessId", _
                "ProcessVersion", processId, processVersion) _
           And StrComp(ReadDesignTableText(lo, values, r, "RequirementId"), _
                       requirementId, vbTextCompare) = 0 Then
            alternativeCount = alternativeCount + 1
            itemCode = ReadDesignTableText(lo, values, r, "ITEM_CODE")
            If candidateItemCode <> "" _
               And StrComp(itemCode, candidateItemCode, vbTextCompare) = 0 Then
                itemAccepted = True
            End If
        End If
    Next r
End Sub

Private Function TryPositiveDouble(ByVal valueIn As Variant, _
                                   ByRef valueOut As Double) As Boolean
    valueOut = 0
    If IsNumeric(valueIn) Then
        valueOut = CDbl(valueIn)
        TryPositiveDouble = (valueOut > 0)
    End If
End Function

Private Function IsReusableProductionEvent(ByVal eventType As String) As Boolean
    Select Case UCase$(Trim$(eventType))
        Case PROCESS_EVENT_SAVE, PROCESS_EVENT_RELEASE, PROCESS_EVENT_OBSOLETE, _
             RECIPE_EVENT_SAVE, RECIPE_EVENT_RELEASE, RECIPE_EVENT_OBSOLETE
            IsReusableProductionEvent = True
    End Select
End Function

Private Function DefinitionTypeForEvent(ByVal eventType As String) As String
    Select Case UCase$(Trim$(eventType))
        Case PROCESS_EVENT_SAVE, PROCESS_EVENT_RELEASE, PROCESS_EVENT_OBSOLETE
            DefinitionTypeForEvent = "PROCESS"
        Case RECIPE_EVENT_SAVE, RECIPE_EVENT_RELEASE, RECIPE_EVENT_OBSOLETE
            DefinitionTypeForEvent = "RECIPE"
        Case Else
            DefinitionTypeForEvent = "LEGACY_DESIGN"
    End Select
End Function

Private Function FindProjectedDesignRow(ByVal lo As ListObject, ByVal designId As String, _
                                        ByVal designVersion As String) As Long
    Dim values As Variant
    Dim r As Long
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    values = lo.DataBodyRange.Value
    For r = 1 To UBound(values, 1)
        If StrComp(ReadDesignTableText(lo, values, r, "DesignId"), designId, vbTextCompare) = 0 _
           And StrComp(ReadDesignTableText(lo, values, r, "DesignVersion"), designVersion, vbTextCompare) = 0 Then
            FindProjectedDesignRow = r
            Exit Function
        End If
    Next r
End Function

Private Function AppliedDesignEventExists(ByVal lo As ListObject, ByVal eventId As String) As Boolean
    Dim values As Variant
    Dim r As Long
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    values = lo.DataBodyRange.Value
    For r = 1 To UBound(values, 1)
        If StrComp(ReadDesignTableText(lo, values, r, "EventID"), eventId, vbTextCompare) = 0 Then
            AppliedDesignEventExists = True
            Exit Function
        End If
    Next r
End Function

Private Function NextDesignAppliedSeq(ByVal lo As ListObject) As Long
    Dim values As Variant
    Dim r As Long
    Dim seq As Long
    If lo Is Nothing Then
        NextDesignAppliedSeq = 1
        Exit Function
    End If
    If lo.DataBodyRange Is Nothing Then
        NextDesignAppliedSeq = 1
        Exit Function
    End If
    values = lo.DataBodyRange.Value
    For r = 1 To UBound(values, 1)
        If IsNumeric(ReadDesignTableValue(lo, values, r, "AppliedSeq")) Then
            If CLng(ReadDesignTableValue(lo, values, r, "AppliedSeq")) > seq Then _
                seq = CLng(ReadDesignTableValue(lo, values, r, "AppliedSeq"))
        End If
    Next r
    NextDesignAppliedSeq = seq + 1
End Function

Private Sub ClearDesignTableRows(ByVal lo As ListObject)
    If lo Is Nothing Then Exit Sub
    Do While lo.ListRows.Count > 0
        lo.ListRows(lo.ListRows.Count).Delete
    Loop
End Sub

Private Sub SaveDesignsWorkbook(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub
    If wb.ReadOnly Then Exit Sub
    If Trim$(wb.Path) = "" Then Exit Sub
    wb.Save
End Sub

Private Function FindDesignsApplyTable(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        On Error Resume Next
        Set FindDesignsApplyTable = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not FindDesignsApplyTable Is Nothing Then Exit Function
    Next ws
End Function

Private Sub SetDesignTableValue(ByVal lo As ListObject, ByVal rowIndex As Long, _
                                ByVal columnName As String, ByVal valueOut As Variant)
    lo.DataBodyRange.Cells(rowIndex, lo.ListColumns(columnName).Index).Value = valueOut
End Sub

Private Function ReadDesignTableValue(ByVal lo As ListObject, ByVal values As Variant, _
                                      ByVal rowIndex As Long, ByVal columnName As String) As Variant
    ReadDesignTableValue = values(rowIndex, lo.ListColumns(columnName).Index)
End Function

Private Function ReadDesignTableText(ByVal lo As ListObject, ByVal values As Variant, _
                                     ByVal rowIndex As Long, ByVal columnName As String) As String
    Dim valueIn As Variant
    valueIn = ReadDesignTableValue(lo, values, rowIndex, columnName)
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Then Exit Function
    ReadDesignTableText = Trim$(CStr(valueIn))
End Function

Private Function GetEventTextDesign(ByVal evt As Object, ByVal key As String) As String
    On Error GoTo CleanExit
    If evt Is Nothing Then Exit Function
    If Not evt.Exists(key) Then Exit Function
    If IsNull(evt(key)) Or IsEmpty(evt(key)) Or IsError(evt(key)) Then Exit Function
    GetEventTextDesign = Trim$(CStr(evt(key)))
CleanExit:
End Function

Private Function TryGetEventDateDesign(ByVal evt As Object, ByVal key As String, _
                                       ByRef dateOut As Date) As Boolean
    Dim textIn As String
    On Error GoTo CleanExit
    textIn = GetEventTextDesign(evt, key)
    If textIn = "" Or Not IsDate(textIn) Then Exit Function
    dateOut = CDate(textIn)
    TryGetEventDateDesign = True
CleanExit:
End Function

Private Function PayloadText(ByVal item As Object, ByVal key As String) As String
    Dim valueIn As Variant
    On Error GoTo CleanExit
    If item Is Nothing Then Exit Function
    If Not item.Exists(key) Then Exit Function
    valueIn = item(key)
    If IsNull(valueIn) Or IsEmpty(valueIn) Or IsError(valueIn) Then Exit Function
    PayloadText = Trim$(CStr(valueIn))
CleanExit:
End Function

Private Function PayloadValue(ByVal item As Object, ByVal key As String) As Variant
    On Error GoTo CleanExit
    If item Is Nothing Then Exit Function
    If Not item.Exists(key) Then Exit Function
    PayloadValue = item(key)
CleanExit:
End Function

Private Function ParseDesignPayload(ByVal jsonText As String, ByRef errorMessage As String) As Collection
    Dim pos As Long
    Dim item As Object
    Set ParseDesignPayload = New Collection
    pos = 1
    SkipJsonWhitespaceDesign jsonText, pos
    If Mid$(jsonText, pos, 1) <> "[" Then
        errorMessage = "PayloadJson must start with '['."
        Set ParseDesignPayload = Nothing
        Exit Function
    End If
    pos = pos + 1
    Do
        SkipJsonWhitespaceDesign jsonText, pos
        If Mid$(jsonText, pos, 1) = "]" Then
            pos = pos + 1
            Exit Do
        End If
        Set item = ParseJsonObjectDesign(jsonText, pos, errorMessage)
        If item Is Nothing Then
            Set ParseDesignPayload = Nothing
            Exit Function
        End If
        ParseDesignPayload.Add item
        SkipJsonWhitespaceDesign jsonText, pos
        If Mid$(jsonText, pos, 1) = "," Then
            pos = pos + 1
        ElseIf Mid$(jsonText, pos, 1) <> "]" Then
            errorMessage = "PayloadJson array is missing a comma separator."
            Set ParseDesignPayload = Nothing
            Exit Function
        End If
    Loop
    SkipJsonWhitespaceDesign jsonText, pos
    If pos <= Len(jsonText) Then
        errorMessage = "PayloadJson contains unexpected trailing characters."
        Set ParseDesignPayload = Nothing
    End If
End Function

Private Function ParseJsonObjectDesign(ByVal jsonText As String, ByRef pos As Long, _
                                       ByRef errorMessage As String) As Object
    Dim item As Object
    Dim key As String
    Set item = CreateObject("Scripting.Dictionary")
    item.CompareMode = vbTextCompare
    SkipJsonWhitespaceDesign jsonText, pos
    If Mid$(jsonText, pos, 1) <> "{" Then
        errorMessage = "PayloadJson object must start with '{'."
        Exit Function
    End If
    pos = pos + 1
    Do
        SkipJsonWhitespaceDesign jsonText, pos
        If Mid$(jsonText, pos, 1) = "}" Then
            pos = pos + 1
            Set ParseJsonObjectDesign = item
            Exit Function
        End If
        key = ParseJsonStringDesign(jsonText, pos, errorMessage)
        If errorMessage <> "" Then Exit Function
        SkipJsonWhitespaceDesign jsonText, pos
        If Mid$(jsonText, pos, 1) <> ":" Then
            errorMessage = "PayloadJson object is missing ':' after key '" & key & "'."
            Exit Function
        End If
        pos = pos + 1
        item(key) = ParseJsonValueDesign(jsonText, pos, errorMessage)
        If errorMessage <> "" Then Exit Function
        SkipJsonWhitespaceDesign jsonText, pos
        If Mid$(jsonText, pos, 1) = "," Then
            pos = pos + 1
        ElseIf Mid$(jsonText, pos, 1) = "}" Then
            pos = pos + 1
            Set ParseJsonObjectDesign = item
            Exit Function
        Else
            errorMessage = "PayloadJson object is missing a comma separator."
            Exit Function
        End If
    Loop
End Function

Private Function ParseJsonValueDesign(ByVal jsonText As String, ByRef pos As Long, _
                                      ByRef errorMessage As String) As Variant
    Dim token As String
    Dim startAt As Long
    SkipJsonWhitespaceDesign jsonText, pos
    If Mid$(jsonText, pos, 1) = Chr$(34) Then
        ParseJsonValueDesign = ParseJsonStringDesign(jsonText, pos, errorMessage)
        Exit Function
    End If
    startAt = pos
    Do While pos <= Len(jsonText) And InStr(1, ",}]" & vbCr & vbLf & vbTab & " ", Mid$(jsonText, pos, 1), vbBinaryCompare) = 0
        pos = pos + 1
    Loop
    token = Trim$(Mid$(jsonText, startAt, pos - startAt))
    Select Case LCase$(token)
        Case "true": ParseJsonValueDesign = True
        Case "false": ParseJsonValueDesign = False
        Case "null": ParseJsonValueDesign = vbNullString
        Case Else
            If IsNumeric(token) Then
                ParseJsonValueDesign = CDbl(token)
            Else
                errorMessage = "Unsupported value in PayloadJson at position " & CStr(startAt) & "."
            End If
    End Select
End Function

Private Function ParseJsonStringDesign(ByVal jsonText As String, ByRef pos As Long, _
                                       ByRef errorMessage As String) As String
    Dim ch As String
    Dim esc As String
    If Mid$(jsonText, pos, 1) <> Chr$(34) Then
        errorMessage = "Expected string value in PayloadJson."
        Exit Function
    End If
    pos = pos + 1
    Do While pos <= Len(jsonText)
        ch = Mid$(jsonText, pos, 1)
        pos = pos + 1
        If ch = Chr$(34) Then Exit Function
        If ch = "\" Then
            If pos > Len(jsonText) Then
                errorMessage = "Incomplete escape sequence in PayloadJson."
                Exit Function
            End If
            esc = Mid$(jsonText, pos, 1)
            pos = pos + 1
            Select Case esc
                Case Chr$(34), "\", "/": ParseJsonStringDesign = ParseJsonStringDesign & esc
                Case "n": ParseJsonStringDesign = ParseJsonStringDesign & vbLf
                Case "r": ParseJsonStringDesign = ParseJsonStringDesign & vbCr
                Case "t": ParseJsonStringDesign = ParseJsonStringDesign & vbTab
                Case Else
                    errorMessage = "Unsupported escape sequence in PayloadJson."
                    Exit Function
            End Select
        Else
            ParseJsonStringDesign = ParseJsonStringDesign & ch
        End If
    Loop
    errorMessage = "Unterminated string in PayloadJson."
End Function

Private Sub SkipJsonWhitespaceDesign(ByVal jsonText As String, ByRef pos As Long)
    Do While pos <= Len(jsonText)
        If InStr(1, " " & vbTab & vbCr & vbLf, Mid$(jsonText, pos, 1), vbBinaryCompare) = 0 Then Exit Do
        pos = pos + 1
    Loop
End Sub
