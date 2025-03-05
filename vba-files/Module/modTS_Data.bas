Attribute VB_Name = "modTS_Data"
' ========================
' Module: modTS_Data
' ========================
Option Explicit

Public Function LoadItemList() As Variant
    Dim ws As Worksheet, tbl As ListObject
    Dim items As Variant
    Dim arr() As Variant, i As Long
    
    Set ws = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
    Set tbl = ws.ListObjects("invSys")
    
    ' Get the ITEM column data (assumes header is "ITEM")
    items = tbl.ListColumns("ITEM").DataBodyRange.Value
    ' Convert the 2D array (n x 1) into a 1D array.
    ReDim arr(LBound(items, 1) To UBound(items, 1))
    For i = LBound(items, 1) To UBound(items, 1)
        arr(i) = items(i, 1)
    Next i
    LoadItemList = arr
End Function
Public Function GetColumnIndexByHeader(headerName As String) As Long
    Dim ws As Worksheet, tbl As ListObject, headers As Variant, i As Long
    Set ws = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
    Set tbl = ws.ListObjects("invSys")
    headers = tbl.HeaderRowRange.Value
    For i = LBound(headers, 2) To UBound(headers, 2)
        If CStr(headers(1, i)) = headerName Then
            GetColumnIndexByHeader = i
            Exit Function
        End If
    Next i
    GetColumnIndexByHeader = 0 ' if not found
End Function







