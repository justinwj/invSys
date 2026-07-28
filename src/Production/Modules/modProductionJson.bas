Attribute VB_Name = "modProductionJson"
Option Explicit

Public Function BuildJsonArray(ByVal items As Collection) As String
    Dim encodedItems() As String
    Dim i As Long

    If items Is Nothing Then
        BuildJsonArray = "[]"
        Exit Function
    End If
    If items.Count = 0 Then
        BuildJsonArray = "[]"
        Exit Function
    End If

    ReDim encodedItems(0 To items.Count - 1)
    For i = 1 To items.Count
        encodedItems(i - 1) = DictionaryToJson(items(i))
    Next i
    BuildJsonArray = "[" & Join(encodedItems, ",") & "]"
End Function

Public Function CreateProductionDeltaPayloadItem(ByVal systemKey As String, _
                                                 ByVal sku As String, _
                                                 ByVal qty As Double, _
                                                 Optional ByVal location As String = "", _
                                                 Optional ByVal noteText As String = "", _
                                                 Optional ByVal ioType As String = "") As Object
    Dim item As Object

    Set item = CreateObject("Scripting.Dictionary")
    item.CompareMode = vbTextCompare
    item("System_Key") = Trim$(systemKey)
    item("SKU") = Trim$(sku)
    item("ITEM_CODE") = Trim$(sku)
    item("Qty") = qty
    item("Location") = Trim$(location)
    item("Note") = noteText
    If Trim$(ioType) <> "" Then item("IoType") = Trim$(ioType)
    Set CreateProductionDeltaPayloadItem = item
End Function

Public Function CreateProductionInventoryEntityPayloadItem(ByVal systemKey As String, _
                                                           ByVal sku As String, _
                                                           ByVal qty As Double, _
                                                           Optional ByVal location As String = "", _
                                                           Optional ByVal conditionValue As String = "GOOD", _
                                                           Optional ByVal attributesJson As String = "", _
                                                           Optional ByVal noteText As String = "") As Object
    Dim item As Object
    Dim normalizedCondition As String

    systemKey = Trim$(systemKey)
    If systemKey = "" Then Err.Raise vbObjectError + 7701, _
        "modProductionJson.CreateProductionInventoryEntityPayloadItem", _
        "System_Key is required."
    normalizedCondition = UCase$(Trim$(conditionValue))
    If normalizedCondition = "" Then normalizedCondition = "GOOD"

    Set item = CreateObject("Scripting.Dictionary")
    item.CompareMode = vbTextCompare
    item("System_Key") = systemKey
    item("SKU") = Trim$(sku)
    item("Qty") = qty
    item("Location") = Trim$(location)
    item("Condition") = normalizedCondition
    item("AttributesJson") = attributesJson
    item("Note") = noteText
    item("IoType") = "CREATE"
    Set CreateProductionInventoryEntityPayloadItem = item
End Function

Private Function DictionaryToJson(ByVal values As Object) As String
    Dim keys As Variant
    Dim i As Long
    Dim key As String

    DictionaryToJson = "{"
    keys = values.Keys
    For i = LBound(keys) To UBound(keys)
        key = CStr(keys(i))
        If i > LBound(keys) Then DictionaryToJson = DictionaryToJson & ","
        DictionaryToJson = DictionaryToJson & """" & EscapeJson(key) & """:" & JsonValue(values(key))
    Next i
    DictionaryToJson = DictionaryToJson & "}"
End Function

Private Function JsonValue(ByVal valueIn As Variant) As String
    Select Case True
        Case IsObject(valueIn)
            JsonValue = "null"
        Case IsNull(valueIn), IsEmpty(valueIn)
            JsonValue = "null"
        Case VarType(valueIn) = vbBoolean
            JsonValue = IIf(CBool(valueIn), "true", "false")
        Case IsNumeric(valueIn)
            JsonValue = Replace$(CStr(valueIn), Application.International(xlDecimalSeparator), ".")
        Case Else
            JsonValue = """" & EscapeJson(CStr(valueIn)) & """"
    End Select
End Function

Private Function EscapeJson(ByVal textIn As String) As String
    Dim escaped As String

    escaped = Replace$(textIn, "\", "\\")
    escaped = Replace$(escaped, Chr$(34), "\" & Chr$(34))
    escaped = Replace$(escaped, vbCrLf, "\n")
    escaped = Replace$(escaped, vbCr, "\n")
    escaped = Replace$(escaped, vbLf, "\n")
    EscapeJson = Replace$(escaped, vbTab, "\t")
End Function
