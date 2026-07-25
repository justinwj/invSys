Attribute VB_Name = "modAdminDesignMigration"
Option Explicit

Private Const LEGACY_DESIGN_VERSION As String = "1"
Private Const DESIGN_EVENT_CREATE As String = "DESIGN_CREATE"

Public Function BuildLegacyRecipeDesignMigrationPlan(ByVal donorWb As Workbook, _
                                                     Optional ByVal tableName As String = "Recipes", _
                                                     Optional ByRef report As String = "") As Collection
    On Error GoTo FailBuild

    Dim plan As New Collection
    Dim lo As ListObject
    Dim groups As Object
    Dim names As Object
    Dim descriptions As Object
    Dim groupOrder As New Collection
    Dim payloadRows As Collection
    Dim payloadItem As Object
    Dim planItem As Object
    Dim payloadJson As String
    Dim designId As String
    Dim designName As String
    Dim componentSku As String
    Dim componentDesignId As String
    Dim groupKey As String
    Dim key As Variant
    Dim r As Long
    Dim lineNo As Long
    Dim cRecipe As Long
    Dim cRecipeId As Long
    Dim cDescription As Long
    Dim cProcess As Long
    Dim cIo As Long
    Dim cIngredient As Long
    Dim cIngredientId As Long
    Dim cItemCode As Long
    Dim cAmount As Long
    Dim cUom As Long
    Dim cPercent As Long
    Dim cInstruction As Long
    Dim cLineNo As Long

    Set BuildLegacyRecipeDesignMigrationPlan = plan
    If donorWb Is Nothing Then
        report = "An explicit legacy recipe donor workbook is required."
        Exit Function
    End If
    If donorWb.IsAddin Then
        report = "An operator/data workbook must be supplied; an XLAM cannot be a migration donor."
        Exit Function
    End If

    Set lo = FindLegacyRecipeTable(donorWb, tableName)
    If lo Is Nothing Then
        report = "Legacy recipe table '" & tableName & "' was not found in " & donorWb.Name & "."
        Exit Function
    End If
    If lo.DataBodyRange Is Nothing Then
        report = "Legacy recipe table '" & tableName & "' has no rows."
        Exit Function
    End If

    cRecipe = LegacyRecipeColumn(lo, "RECIPE")
    cRecipeId = LegacyRecipeColumn(lo, "RECIPE_ID")
    cDescription = LegacyRecipeColumn(lo, "DESCRIPTION")
    cProcess = LegacyRecipeColumn(lo, "PROCESS")
    cIo = LegacyRecipeColumn(lo, "INPUT/OUTPUT")
    cIngredient = LegacyRecipeColumn(lo, "INGREDIENT")
    cIngredientId = LegacyRecipeColumn(lo, "INGREDIENT_ID")
    cItemCode = LegacyRecipeColumn(lo, "ITEM_CODE")
    cAmount = LegacyRecipeColumn(lo, "AMOUNT")
    cUom = LegacyRecipeColumn(lo, "UOM")
    cPercent = LegacyRecipeColumn(lo, "PERCENT")
    cInstruction = LegacyRecipeColumn(lo, "INSTRUCTION")
    cLineNo = LegacyRecipeColumn(lo, "RECIPE_LIST_ROW")
    If cRecipe = 0 Then
        report = "Legacy recipe table is missing RECIPE."
        Exit Function
    End If

    Set groups = CreateObject("Scripting.Dictionary")
    groups.CompareMode = vbTextCompare
    Set names = CreateObject("Scripting.Dictionary")
    names.CompareMode = vbTextCompare
    Set descriptions = CreateObject("Scripting.Dictionary")
    descriptions.CompareMode = vbTextCompare

    For r = 1 To lo.ListRows.Count
        designName = LegacyRecipeText(lo, r, cRecipe)
        If designName = "" Then GoTo NextRow
        designId = LegacyRecipeText(lo, r, cRecipeId)
        If designId = "" Then designId = "LEGACY-" & LegacyDesignToken(designName)
        groupKey = UCase$(designId) & "|" & LEGACY_DESIGN_VERSION

        If Not groups.Exists(groupKey) Then
            Set payloadRows = New Collection
            groups.Add groupKey, payloadRows
            names.Add groupKey, designName
            descriptions.Add groupKey, LegacyRecipeText(lo, r, cDescription)
            groupOrder.Add groupKey
        Else
            Set payloadRows = groups(groupKey)
        End If

        Set payloadItem = CreateObject("Scripting.Dictionary")
        payloadItem.CompareMode = vbTextCompare
        If payloadRows.Count = 0 Then
            payloadItem("DesignType") = "RECIPE"
            payloadItem("DesignName") = names(groupKey)
            payloadItem("Description") = descriptions(groupKey)
        End If
        lineNo = payloadRows.Count + 1
        If cLineNo > 0 Then
            If IsNumeric(lo.DataBodyRange.Cells(r, cLineNo).Value) Then lineNo = CLng(lo.DataBodyRange.Cells(r, cLineNo).Value)
        End If
        payloadItem("LineNo") = lineNo
        If cProcess > 0 Then payloadItem("Process") = LegacyRecipeText(lo, r, cProcess)
        If cIo > 0 Then payloadItem("IOType") = LegacyRecipeText(lo, r, cIo)
        componentSku = LegacyRecipeText(lo, r, cItemCode)
        componentDesignId = LegacyRecipeText(lo, r, cIngredientId)
        If componentSku <> "" Then payloadItem("ComponentSKU") = componentSku
        If componentDesignId <> "" Then payloadItem("ComponentDesignId") = componentDesignId
        If componentSku = "" And componentDesignId = "" Then
            componentSku = LegacyRecipeText(lo, r, cIngredient)
            If componentSku <> "" Then payloadItem("ComponentSKU") = componentSku
        End If
        If cAmount > 0 Then payloadItem("Qty") = LegacyRecipeNumber(lo.DataBodyRange.Cells(r, cAmount).Value)
        If cUom > 0 Then payloadItem("UOM") = LegacyRecipeText(lo, r, cUom)
        If cPercent > 0 Then payloadItem("Percent") = LegacyRecipeNumber(lo.DataBodyRange.Cells(r, cPercent).Value)
        If cInstruction > 0 Then payloadItem("Instruction") = LegacyRecipeText(lo, r, cInstruction)
        payloadRows.Add payloadItem
NextRow:
    Next r

    For Each key In groupOrder
        Set payloadRows = groups(CStr(key))
        payloadJson = modRoleEventWriter.BuildPayloadJsonFromCollection(payloadRows)
        designId = Left$(CStr(key), InStrRev(CStr(key), "|") - 1)
        Set planItem = CreateObject("Scripting.Dictionary")
        planItem.CompareMode = vbTextCompare
        planItem("EventType") = DESIGN_EVENT_CREATE
        planItem("DesignId") = designId
        planItem("DesignVersion") = LEGACY_DESIGN_VERSION
        planItem("PayloadJson") = payloadJson
        planItem("MigrationSourceId") = LegacyMigrationSourceId(donorWb, lo, designId)
        planItem("EventID") = LegacyDesignMigrationEventId(designId, LEGACY_DESIGN_VERSION, payloadJson)
        plan.Add planItem
    Next key

    report = "Prepared " & CStr(plan.Count) & " legacy recipe design event(s) from " & donorWb.Name & "."
    Exit Function

FailBuild:
    report = "BuildLegacyRecipeDesignMigrationPlan failed: " & Err.Description
End Function

Public Function QueueLegacyRecipeDesignMigration(ByVal donorWb As Workbook, _
                                                 Optional ByVal tableName As String = "Recipes", _
                                                 Optional ByRef report As String = "") As Boolean
    Dim plan As Collection
    Dim planItem As Object
    Dim eventIdOut As String
    Dim queueError As String
    Dim queued As Long

    Set plan = BuildLegacyRecipeDesignMigrationPlan(donorWb, tableName, report)
    If plan Is Nothing Then Exit Function
    If plan.Count = 0 Then Exit Function

    For Each planItem In plan
        eventIdOut = CStr(planItem("EventID"))
        queueError = ""
        If Not modRoleEventWriter.QueueDesignEventCurrent( _
            CStr(planItem("EventType")), _
            CStr(planItem("DesignId")), _
            CStr(planItem("DesignVersion")), _
            CStr(planItem("PayloadJson")), _
            "Explicit legacy recipe migration", _
            "", _
            eventIdOut, _
            queueError, _
            "", _
            CStr(planItem("MigrationSourceId")), _
            CStr(planItem("EventID"))) Then
            report = "Queued " & CStr(queued) & " event(s), then failed for " & _
                     CStr(planItem("DesignId")) & ": " & queueError
            Exit Function
        End If
        queued = queued + 1
    Next planItem

    report = "Queued " & CStr(queued) & " legacy recipe DESIGN_CREATE event(s)."
    QueueLegacyRecipeDesignMigration = True
End Function

Private Function FindLegacyRecipeTable(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    On Error Resume Next
    For Each ws In wb.Worksheets
        Set FindLegacyRecipeTable = ws.ListObjects(tableName)
        If Not FindLegacyRecipeTable Is Nothing Then Exit Function
    Next ws
    On Error GoTo 0
End Function

Private Function LegacyRecipeColumn(ByVal lo As ListObject, ByVal columnName As String) As Long
    Dim lc As ListColumn
    If lo Is Nothing Then Exit Function
    For Each lc In lo.ListColumns
        If StrComp(Trim$(lc.Name), Trim$(columnName), vbTextCompare) = 0 Then
            LegacyRecipeColumn = lc.Index
            Exit Function
        End If
    Next lc
End Function

Private Function LegacyRecipeText(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnIndex As Long) As String
    If lo Is Nothing Or columnIndex <= 0 Then Exit Function
    If rowIndex <= 0 Or rowIndex > lo.ListRows.Count Then Exit Function
    If IsError(lo.DataBodyRange.Cells(rowIndex, columnIndex).Value) Then Exit Function
    LegacyRecipeText = Trim$(CStr(lo.DataBodyRange.Cells(rowIndex, columnIndex).Value))
End Function

Private Function LegacyRecipeNumber(ByVal valueIn As Variant) As Double
    If IsError(valueIn) Or IsEmpty(valueIn) Then Exit Function
    If IsNumeric(valueIn) Then LegacyRecipeNumber = CDbl(valueIn)
End Function

Private Function LegacyDesignToken(ByVal valueIn As String) As String
    Dim result As String
    Dim ch As String
    Dim i As Long

    valueIn = UCase$(Trim$(valueIn))
    For i = 1 To Len(valueIn)
        ch = Mid$(valueIn, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then
            result = result & ch
        ElseIf result <> "" Then
            If Right$(result, 1) <> "-" Then result = result & "-"
        End If
    Next i
    Do While Right$(result, 1) = "-"
        result = Left$(result, Len(result) - 1)
    Loop
    If result = "" Then result = "UNNAMED"
    LegacyDesignToken = result
End Function

Private Function LegacyMigrationSourceId(ByVal donorWb As Workbook, ByVal lo As ListObject, _
                                         ByVal designId As String) As String
    LegacyMigrationSourceId = "LEGACY_RECIPE|" & UCase$(donorWb.Name) & "|" & _
                              UCase$(lo.Name) & "|" & UCase$(designId) & "|" & LEGACY_DESIGN_VERSION
End Function

Private Function LegacyDesignMigrationEventId(ByVal designId As String, ByVal designVersion As String, _
                                              ByVal payloadJson As String) As String
    LegacyDesignMigrationEventId = "MIG-DESIGN-" & LegacyDesignToken(designId) & "-V" & _
                                   LegacyDesignToken(designVersion) & "-" & StableLegacyPayloadHash(payloadJson)
End Function

Private Function StableLegacyPayloadHash(ByVal textIn As String) As String
    Dim hashValue As Double
    Dim i As Long

    For i = 1 To Len(textIn)
        hashValue = hashValue * 131# + AscW(Mid$(textIn, i, 1))
        hashValue = hashValue - Int(hashValue / 2147483647#) * 2147483647#
    Next i
    StableLegacyPayloadHash = Right$("00000000" & Hex$(CLng(hashValue)), 8)
End Function
