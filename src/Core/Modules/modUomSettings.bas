Attribute VB_Name = "modUomSettings"
Option Explicit

Private Const CONFIG_KEY_UOM_CATALOG As String = "UomCatalog"
Private Const UOM_DELIMITER As String = "|"
Private Const DEFAULT_UOMS As String = "EA|LB|LBS|OZ|KG|G|GAL|QT|PT|L|ML|CS"

Public Function GetConfiguredUoms() As Variant
    Dim uoms As Collection
    Dim result() As Variant
    Dim idx As Long

    Set uoms = ConfiguredUomCollection()
    If uoms Is Nothing Or uoms.Count = 0 Then Exit Function

    ReDim result(1 To uoms.Count)
    For idx = 1 To uoms.Count
        result(idx) = CStr(uoms(idx))
    Next idx
    GetConfiguredUoms = result
End Function

Public Function AddConfiguredUom(ByVal uomName As String, _
                                 Optional ByRef report As String = "") As Boolean
    Dim uoms As Collection

    uomName = NormalizeUomName(uomName)
    If uomName = "" Then
        report = "Enter a UOM containing letters or numbers."
        Exit Function
    End If

    Set uoms = ConfiguredUomCollection()
    If UomCollectionContains(uoms, uomName) Then
        report = uomName & " is already in the warehouse UOM catalog."
        AddConfiguredUom = True
        Exit Function
    End If

    uoms.Add uomName
    AddConfiguredUom = SaveUomCollection(uoms, report)
End Function

Public Function RemoveConfiguredUom(ByVal uomName As String, _
                                    Optional ByRef report As String = "") As Boolean
    Dim uoms As Collection
    Dim idx As Long

    uomName = NormalizeUomName(uomName)
    If uomName = "" Then
        report = "Select a UOM to remove."
        Exit Function
    End If

    Set uoms = ConfiguredUomCollection()
    For idx = uoms.Count To 1 Step -1
        If StrComp(CStr(uoms(idx)), uomName, vbTextCompare) = 0 Then uoms.Remove idx
    Next idx
    If uoms.Count = 0 Then
        report = "The warehouse UOM catalog must contain at least one value."
        Exit Function
    End If
    RemoveConfiguredUom = SaveUomCollection(uoms, report)
End Function

Public Function ResetConfiguredUoms(Optional ByRef report As String = "") As Boolean
    ResetConfiguredUoms = SaveUomCollection(ParseUomPackedText(DEFAULT_UOMS), report)
End Function

Public Function GetConfiguredUomsPackedText() As String
    GetConfiguredUomsPackedText = PackUomCollection(ConfiguredUomCollection())
End Function

Public Function NormalizeConfiguredUomName(ByVal uomName As String) As String
    NormalizeConfiguredUomName = NormalizeUomName(uomName)
End Function

Private Function ConfiguredUomCollection() As Collection
    Dim packed As String

    packed = Trim$(modConfig.GetString(CONFIG_KEY_UOM_CATALOG, DEFAULT_UOMS))
    If packed = "" Then packed = DEFAULT_UOMS
    Set ConfiguredUomCollection = ParseUomPackedText(packed)
    If ConfiguredUomCollection.Count = 0 Then _
        Set ConfiguredUomCollection = ParseUomPackedText(DEFAULT_UOMS)
End Function

Private Function ParseUomPackedText(ByVal packedText As String) As Collection
    Dim normalized As String
    Dim parts As Variant
    Dim idx As Long
    Dim uoms As New Collection
    Dim uomName As String

    normalized = Replace$(packedText, vbCrLf, UOM_DELIMITER)
    normalized = Replace$(normalized, vbCr, UOM_DELIMITER)
    normalized = Replace$(normalized, vbLf, UOM_DELIMITER)
    normalized = Replace$(normalized, ",", UOM_DELIMITER)
    normalized = Replace$(normalized, ";", UOM_DELIMITER)
    parts = Split(normalized, UOM_DELIMITER)
    For idx = LBound(parts) To UBound(parts)
        uomName = NormalizeUomName(CStr(parts(idx)))
        If uomName <> "" Then
            If Not UomCollectionContains(uoms, uomName) Then uoms.Add uomName
        End If
    Next idx
    Set ParseUomPackedText = uoms
End Function

Private Function SaveUomCollection(ByVal uoms As Collection, ByRef report As String) As Boolean
    Dim packed As String

    packed = PackUomCollection(uoms)
    If packed = "" Then
        report = "The warehouse UOM catalog must contain at least one value."
        Exit Function
    End If
    SaveUomCollection = modConfig.UpdateConfigValue(CONFIG_KEY_UOM_CATALOG, packed, report)
End Function

Private Function PackUomCollection(ByVal uoms As Collection) As String
    Dim idx As Long
    Dim uomName As String

    If uoms Is Nothing Then Exit Function
    For idx = 1 To uoms.Count
        uomName = NormalizeUomName(CStr(uoms(idx)))
        If uomName <> "" Then
            If PackUomCollection <> "" Then PackUomCollection = PackUomCollection & UOM_DELIMITER
            PackUomCollection = PackUomCollection & uomName
        End If
    Next idx
End Function

Private Function UomCollectionContains(ByVal uoms As Collection, ByVal uomName As String) As Boolean
    Dim idx As Long

    If uoms Is Nothing Then Exit Function
    For idx = 1 To uoms.Count
        If StrComp(CStr(uoms(idx)), uomName, vbTextCompare) = 0 Then
            UomCollectionContains = True
            Exit Function
        End If
    Next idx
End Function

Private Function NormalizeUomName(ByVal uomName As String) As String
    Dim idx As Long
    Dim ch As String
    Dim normalized As String

    uomName = UCase$(Trim$(uomName))
    For idx = 1 To Len(uomName)
        ch = Mid$(uomName, idx, 1)
        If ch Like "[A-Z0-9]" Or ch = "/" Or ch = "-" Then
            normalized = normalized & ch
        ElseIf ch = " " Then
            If normalized <> "" Then
                If Right$(normalized, 1) <> " " Then normalized = normalized & " "
            End If
        End If
    Next idx
    NormalizeUomName = Trim$(normalized)
End Function
