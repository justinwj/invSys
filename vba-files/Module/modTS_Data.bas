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

' Add this function to lookup UOM by item name
Public Function GetItemUOM(itemName As String) As String
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim itemCol As Range, uomCol As Range
    Dim foundCell As Range
    Dim foundRow As Long
    
    ' Default return value if not found
    GetItemUOM = "each"
    
    Set ws = ThisWorkbook.Sheets("INVENTORY MANAGEMENT")
    Set tbl = ws.ListObjects("invSys")
    
    Set itemCol = tbl.ListColumns("ITEM").DataBodyRange
    Set uomCol = tbl.ListColumns("UOM").DataBodyRange
    
    ' Find the item in the invSys table
    Set foundCell = itemCol.Find(What:=itemName, _
                                 LookIn:=xlValues, _
                                 LookAt:=xlWhole, _
                                 SearchOrder:=xlByRows, _
                                 MatchCase:=False)
    
    ' If found, return its UOM
    If Not foundCell Is Nothing Then
        foundRow = foundCell.Row - itemCol.Row + 1
        GetItemUOM = uomCol.Cells(foundRow, 1).Value
        
        ' If UOM is empty, return default
        If Trim(GetItemUOM) = "" Then
            GetItemUOM = "each"
        End If
    End If
End Function







