Attribute VB_Name = "modInventoryQueries"
Option Explicit

Public Function GetOnHandQty(ByVal sku As String, _
                             Optional ByVal inventoryWb As Workbook = Nothing) As Double
    On Error GoTo CleanFail

    Dim wb As Workbook
    Dim lo As ListObject
    Dim rowIndex As Long
    Dim report As String

    sku = Trim$(sku)
    If sku = "" Then Exit Function
    Set wb = modInventoryApply.ResolveInventoryWorkbook("", inventoryWb)
    If wb Is Nothing Then Exit Function
    Set lo = FindInventoryQueryTable(wb, "tblSkuBalance")
    rowIndex = FindInventoryQueryRow(lo, "SKU", sku)
    If rowIndex = 0 Then Exit Function
    GetOnHandQty = NzInventoryQueryNumber(ReadInventoryQueryValue(lo, rowIndex, "QtyOnHand"))
CleanFail:
End Function

Public Function GetLocationBalances(ByVal sku As String, _
                                    Optional ByVal inventoryWb As Workbook = Nothing) As Variant
    On Error GoTo CleanFail

    Dim wb As Workbook
    Dim lo As ListObject
    Dim src As Variant
    Dim result() As Variant
    Dim trimmed() As Variant
    Dim r As Long
    Dim c As Long
    Dim outRow As Long

    sku = Trim$(sku)
    If sku = "" Then Exit Function
    Set wb = modInventoryApply.ResolveInventoryWorkbook("", inventoryWb)
    If wb Is Nothing Then Exit Function
    Set lo = FindInventoryQueryTable(wb, "tblLocationBalance")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    src = lo.DataBodyRange.Value
    ReDim result(1 To UBound(src, 1), 1 To 3)
    For r = 1 To UBound(src, 1)
        If StrComp(CStr(src(r, lo.ListColumns("SKU").Index)), sku, vbTextCompare) = 0 Then
            outRow = outRow + 1
            result(outRow, 1) = src(r, lo.ListColumns("Location").Index)
            result(outRow, 2) = src(r, lo.ListColumns("QtyOnHand").Index)
            result(outRow, 3) = src(r, lo.ListColumns("LastAppliedUTC").Index)
        End If
    Next r
    If outRow = 0 Then Exit Function
    ReDim trimmed(1 To outRow, 1 To 3)
    For r = 1 To outRow
        For c = 1 To 3
            trimmed(r, c) = result(r, c)
        Next c
    Next r
    GetLocationBalances = trimmed
CleanFail:
End Function

Public Function ListInventoryPickerItems(Optional ByVal filterText As String = "", _
                                         Optional ByVal inventoryWb As Workbook = Nothing) As Variant
    On Error GoTo CleanFail

    Dim wb As Workbook
    Dim loCatalog As ListObject
    Dim loBalance As ListObject
    Dim loLocation As ListObject
    Dim catalogRows As Variant
    Dim locationRows As Variant
    Dim result() As Variant
    Dim trimmed() As Variant
    Dim r As Long
    Dim lr As Long
    Dim c As Long
    Dim outRow As Long
    Dim maxRows As Long
    Dim sku As String
    Dim systemKey As String
    Dim itemCode As String
    Dim itemName As String
    Dim uom As String
    Dim catalogLocation As String
    Dim description As String
    Dim category As String
    Dim locationValue As String
    Dim qtyOnHand As Variant
    Dim matchedLocation As Boolean
    Dim haystack As String

    Set wb = modInventoryApply.ResolveInventoryWorkbook("", inventoryWb)
    If wb Is Nothing Then Exit Function
    Set loCatalog = FindInventoryQueryTable(wb, "tblSkuCatalog")
    If loCatalog Is Nothing Or loCatalog.DataBodyRange Is Nothing Then Exit Function
    Set loBalance = FindInventoryQueryTable(wb, "tblSkuBalance")
    Set loLocation = FindInventoryQueryTable(wb, "tblLocationBalance")

    catalogRows = loCatalog.DataBodyRange.Value
    maxRows = UBound(catalogRows, 1)
    If Not loLocation Is Nothing Then
        If Not loLocation.DataBodyRange Is Nothing Then
            locationRows = loLocation.DataBodyRange.Value
            maxRows = maxRows + UBound(locationRows, 1)
        End If
    End If
    ReDim result(1 To maxRows, 1 To 7)
    filterText = LCase$(Trim$(filterText))

    For r = 1 To UBound(catalogRows, 1)
        sku = InventoryQueryText(catalogRows, r, InventoryQueryColumn(loCatalog, "SKU"))
        systemKey = InventoryQueryText(catalogRows, r, InventoryQueryColumn(loCatalog, "System_Key"))
        itemCode = InventoryQueryText(catalogRows, r, InventoryQueryColumn(loCatalog, "ITEM_CODE"))
        itemName = InventoryQueryText(catalogRows, r, InventoryQueryColumn(loCatalog, "ITEM"))
        uom = InventoryQueryText(catalogRows, r, InventoryQueryColumn(loCatalog, "UOM"))
        catalogLocation = InventoryQueryText(catalogRows, r, InventoryQueryColumn(loCatalog, "LOCATION"))
        description = InventoryQueryText(catalogRows, r, InventoryQueryColumn(loCatalog, "DESCRIPTION"))
        category = InventoryQueryText(catalogRows, r, InventoryQueryColumn(loCatalog, "CATEGORY"))
        If itemCode = "" Then itemCode = sku
        If sku = "" Then sku = itemCode
        If itemName = "" And itemCode = "" Then GoTo NextCatalogRow

        haystack = LCase$(systemKey & " " & itemCode & " " & itemName & " " & uom & " " & _
                              catalogLocation & " " & description & " " & category)
        If filterText <> "" Then
            If InStr(1, haystack, filterText, vbTextCompare) = 0 Then GoTo NextCatalogRow
        End If

        If InventoryQueryIsNonCounted(category) Then
            qtyOnHand = "utility"
        Else
            qtyOnHand = InventoryQuerySkuBalance(loBalance, sku)
        End If

        matchedLocation = False
        If IsArray(locationRows) Then
            For lr = LBound(locationRows, 1) To UBound(locationRows, 1)
                If StrComp(InventoryQueryText(locationRows, lr, InventoryQueryColumn(loLocation, "SKU")), _
                           sku, vbTextCompare) = 0 Then
                    locationValue = InventoryQueryText(locationRows, lr, InventoryQueryColumn(loLocation, "Location"))
                    If locationValue <> "" Then
                        AddInventoryPickerResultRow result, outRow, systemKey, itemName, uom, qtyOnHand, _
                                                    locationValue, description, itemCode
                        matchedLocation = True
                    End If
                End If
            Next lr
        End If
        If Not matchedLocation Then
            AddInventoryPickerResultRow result, outRow, systemKey, itemName, uom, qtyOnHand, _
                                        catalogLocation, description, itemCode
        End If
NextCatalogRow:
    Next r

    If outRow = 0 Then Exit Function
    ReDim trimmed(1 To outRow, 1 To 7)
    For r = 1 To outRow
        For c = 1 To 7
            trimmed(r, c) = result(r, c)
        Next c
    Next r
    ListInventoryPickerItems = trimmed
CleanFail:
End Function

Public Function ListAvailableInventoryEntities(Optional ByVal filterText As String = "", _
                                               Optional ByVal inventoryWb As Workbook = Nothing) As Variant
    On Error GoTo CleanFail

    Dim wb As Workbook
    Dim loEntities As ListObject
    Dim loCatalog As ListObject
    Dim entityRows As Variant
    Dim result() As Variant
    Dim trimmed() As Variant
    Dim r As Long
    Dim c As Long
    Dim catalogRow As Long
    Dim outRow As Long
    Dim systemKey As String
    Dim sku As String
    Dim itemCode As String
    Dim itemName As String
    Dim uom As String
    Dim description As String
    Dim qtyOnHand As Double
    Dim locationValue As String
    Dim conditionValue As String
    Dim inventoryState As String
    Dim attributesJson As String
    Dim haystack As String

    Set wb = modInventoryApply.ResolveInventoryWorkbook("", inventoryWb)
    If wb Is Nothing Then Exit Function
    Set loEntities = FindInventoryQueryTable(wb, "tblInventoryEntities")
    If loEntities Is Nothing Or loEntities.DataBodyRange Is Nothing Then Exit Function
    Set loCatalog = FindInventoryQueryTable(wb, "tblSkuCatalog")
    entityRows = loEntities.DataBodyRange.Value2
    ReDim result(1 To UBound(entityRows, 1), 1 To 10)
    filterText = LCase$(Trim$(filterText))

    For r = 1 To UBound(entityRows, 1)
        systemKey = InventoryQueryText(entityRows, r, InventoryQueryColumn(loEntities, "System_Key"))
        sku = InventoryQueryText(entityRows, r, InventoryQueryColumn(loEntities, "SKU"))
        qtyOnHand = NzInventoryQueryNumber(entityRows(r, InventoryQueryColumn(loEntities, "QtyOnHand")))
        locationValue = InventoryQueryText(entityRows, r, InventoryQueryColumn(loEntities, "Location"))
        conditionValue = InventoryQueryText(entityRows, r, InventoryQueryColumn(loEntities, "Condition"))
        inventoryState = InventoryQueryText(entityRows, r, InventoryQueryColumn(loEntities, "InventoryState"))
        attributesJson = InventoryQueryText(entityRows, r, InventoryQueryColumn(loEntities, "AttributesJson"))
        If systemKey = "" Or sku = "" Or qtyOnHand <= 0 Then GoTo NextEntity
        If inventoryState <> "" And StrComp(inventoryState, "ACTIVE", vbTextCompare) <> 0 Then GoTo NextEntity

        catalogRow = FindInventoryQueryRow(loCatalog, "SKU", sku)
        If catalogRow > 0 Then
            itemCode = Trim$(CStr(ReadInventoryQueryValue(loCatalog, catalogRow, "ITEM_CODE")))
            itemName = Trim$(CStr(ReadInventoryQueryValue(loCatalog, catalogRow, "ITEM")))
            uom = Trim$(CStr(ReadInventoryQueryValue(loCatalog, catalogRow, "UOM")))
            description = Trim$(CStr(ReadInventoryQueryValue(loCatalog, catalogRow, "DESCRIPTION")))
        Else
            itemCode = sku
            itemName = sku
            uom = ""
            description = ""
        End If
        If itemCode = "" Then itemCode = sku
        If itemName = "" Then itemName = itemCode
        haystack = LCase$(systemKey & " " & sku & " " & itemCode & " " & itemName & " " & _
                          uom & " " & locationValue & " " & conditionValue & " " & description)
        If filterText <> "" Then
            If InStr(1, haystack, filterText, vbTextCompare) = 0 Then GoTo NextEntity
        End If

        outRow = outRow + 1
        result(outRow, 1) = systemKey
        result(outRow, 2) = sku
        result(outRow, 3) = itemCode
        result(outRow, 4) = itemName
        result(outRow, 5) = uom
        result(outRow, 6) = qtyOnHand
        result(outRow, 7) = locationValue
        result(outRow, 8) = conditionValue
        result(outRow, 9) = inventoryState
        result(outRow, 10) = attributesJson
NextEntity:
    Next r

    If outRow = 0 Then Exit Function
    ReDim trimmed(1 To outRow, 1 To 10)
    For r = 1 To outRow
        For c = 1 To 10
            trimmed(r, c) = result(r, c)
        Next c
    Next r
    ListAvailableInventoryEntities = trimmed
CleanFail:
End Function

Private Sub AddInventoryPickerResultRow(ByRef result() As Variant, ByRef outRow As Long, _
                                        ByVal systemKey As String, ByVal itemName As String, _
                                        ByVal uom As String, ByVal qtyOnHand As Variant, _
                                        ByVal locationValue As String, ByVal description As String, _
                                        ByVal itemCode As String)
    outRow = outRow + 1
    result(outRow, 1) = systemKey
    result(outRow, 2) = itemName
    result(outRow, 3) = uom
    result(outRow, 4) = qtyOnHand
    result(outRow, 5) = locationValue
    result(outRow, 6) = description
    result(outRow, 7) = itemCode
End Sub

Private Function InventoryQuerySkuBalance(ByVal loBalance As ListObject, ByVal sku As String) As Variant
    Dim rowIndex As Long
    If loBalance Is Nothing Then Exit Function
    rowIndex = FindInventoryQueryRow(loBalance, "SKU", sku)
    If rowIndex > 0 Then InventoryQuerySkuBalance = ReadInventoryQueryValue(loBalance, rowIndex, "QtyOnHand")
End Function

Private Function InventoryQueryColumn(ByVal lo As ListObject, ByVal columnName As String) As Long
    If lo Is Nothing Then Exit Function
    On Error Resume Next
    InventoryQueryColumn = lo.ListColumns(columnName).Index
    On Error GoTo 0
End Function

Private Function InventoryQueryText(ByVal rows As Variant, ByVal rowIndex As Long, _
                                    ByVal columnIndex As Long) As String
    If columnIndex = 0 Then Exit Function
    If IsError(rows(rowIndex, columnIndex)) Or IsNull(rows(rowIndex, columnIndex)) _
       Or IsEmpty(rows(rowIndex, columnIndex)) Then Exit Function
    InventoryQueryText = Trim$(CStr(rows(rowIndex, columnIndex)))
End Function

Private Function InventoryQueryIsNonCounted(ByVal category As String) As Boolean
    category = UCase$(Trim$(category))
    InventoryQueryIsNonCounted = (category = "UTILITY" Or category = "SERVICE" Or category = "NON_COUNTED")
End Function

Private Function FindInventoryQueryTable(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        On Error Resume Next
        Set FindInventoryQueryTable = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not FindInventoryQueryTable Is Nothing Then Exit Function
    Next ws
End Function

Private Function FindInventoryQueryRow(ByVal lo As ListObject, ByVal columnName As String, _
                                       ByVal matchValue As String) As Long
    Dim r As Long
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    For r = 1 To lo.ListRows.Count
        If StrComp(CStr(ReadInventoryQueryValue(lo, r, columnName)), matchValue, vbTextCompare) = 0 Then
            FindInventoryQueryRow = r
            Exit Function
        End If
    Next r
End Function

Private Function ReadInventoryQueryValue(ByVal lo As ListObject, ByVal rowIndex As Long, _
                                         ByVal columnName As String) As Variant
    ReadInventoryQueryValue = lo.DataBodyRange.Cells(rowIndex, lo.ListColumns(columnName).Index).Value
End Function

Private Function NzInventoryQueryNumber(ByVal valueIn As Variant) As Double
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Then Exit Function
    If Trim$(CStr(valueIn)) = "" Then Exit Function
    NzInventoryQueryNumber = CDbl(valueIn)
End Function
