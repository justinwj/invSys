Attribute VB_Name = "modInventoryBridgeApi"
Option Explicit

Public Function ResolveInventoryWorkbookBridgeResult(Optional ByVal warehouseId As String = "", _
                                                    Optional ByVal inventoryWb As Workbook = Nothing) As Workbook
    Set ResolveInventoryWorkbookBridgeResult = modInventoryApply.ResolveInventoryWorkbook(warehouseId, inventoryWb)
End Function

Public Function EnsureInventorySchemaBridgeResult(Optional ByVal targetWb As Workbook = Nothing) As Object
    Dim result As Object
    Dim report As String

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    result("Success") = modInventorySchema.EnsureInventorySchema(targetWb, report)
    result("Report") = report
    Set EnsureInventorySchemaBridgeResult = result
End Function

Public Function ApplyEventBridgeResult(ByVal evt As Object, _
                                      Optional ByVal inventoryWb As Workbook = Nothing, _
                                      Optional ByVal runId As String = "") As Object
    Dim result As Object
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    result("Success") = modInventoryApply.ApplyEvent(evt, inventoryWb, runId, statusOut, errorCode, errorMessage)
    result("StatusOut") = statusOut
    result("ErrorCode") = errorCode
    result("ErrorMessage") = errorMessage
    Set ApplyEventBridgeResult = result
End Function

Public Function RemoveLastBulkLogEntriesBridgeResult(ByVal countToRemove As Long) As Collection
    Set RemoveLastBulkLogEntriesBridgeResult = modInvMan.RemoveLastBulkLogEntries(countToRemove)
End Function

Public Sub ReAddBulkLogEntriesBridgeResult(ByVal logDataCollection As Collection)
    modInvMan.ReAddBulkLogEntries logDataCollection
End Sub
