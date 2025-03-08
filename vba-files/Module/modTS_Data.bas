Attribute VB_Name = "modTS_Data"
' ========================
' Module: modTS_Data
' ========================
Option Explicit

Public Function LoadItemList() As Variant
    Dim ws As Worksheet, tbl As ListObject
    Dim data As Variant
    Dim result() As Variant, i As Long
    
    Set ws = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
    Set tbl = ws.ListObjects("invSys")
    
    ' Get multiple columns: ITEM_CODE, VENDOR(s), ITEM, and DESCRIPTION
    Dim itemCodeCol As Range, vendorCol As Range, itemCol As Range, descCol As Range
    
    Set itemCodeCol = tbl.ListColumns("ITEM_CODE").DataBodyRange
    Set vendorCol = tbl.ListColumns("VENDOR(s)").DataBodyRange
    Set itemCol = tbl.ListColumns("ITEM").DataBodyRange
    Set descCol = tbl.ListColumns("DESCRIPTION").DataBodyRange
    
    ' Prepare a 2D array to hold the data - now with 4 columns
    ReDim result(1 To itemCol.Rows.count, 1 To 4)
    
    ' Fill the array with the four columns of data
    For i = 1 To itemCol.Rows.count
        result(i, 1) = itemCodeCol.Cells(i, 1).Value
        result(i, 2) = vendorCol.Cells(i, 1).Value
        result(i, 3) = itemCol.Cells(i, 1).Value
        result(i, 4) = descCol.Cells(i, 1).Value
    Next i
    
    LoadItemList = result
End Function







