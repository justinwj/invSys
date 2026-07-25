Attribute VB_Name = "modDesignsBridgeApi"
Option Explicit

Public Function ResolveDesignsWorkbookBridgeResult(Optional ByVal warehouseId As String = "") As Workbook
    Dim report As String
    Set ResolveDesignsWorkbookBridgeResult = modDesignsRuntime.ResolveDesignsWorkbook(warehouseId, Nothing, report)
End Function

Public Function EnsureDesignsSchemaBridgeEncoded(ByVal targetWb As Workbook) As String
    Dim report As String
    Dim succeeded As Boolean

    succeeded = modDesignsSchema.EnsureDesignsSchema(targetWb, report)
    EnsureDesignsSchemaBridgeEncoded = IIf(succeeded, "OK", "FAIL") & vbTab & report
End Function

Public Function ValidateDesignsSchemaBridgeResult(ByVal targetWb As Workbook) As String
    ValidateDesignsSchemaBridgeResult = modDesignsSchema.ValidateDesignsSchema(targetWb)
End Function

Public Function ApplyDesignEventBridgeEncoded(ByVal evt As Object, ByVal designsWb As Workbook, _
                                              Optional ByVal runId As String = "") As String
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim succeeded As Boolean

    succeeded = modDesignsApply.ApplyDesignEvent(evt, designsWb, runId, statusOut, errorCode, errorMessage)
    ApplyDesignEventBridgeEncoded = CStr(Abs(CLng(succeeded))) & vbTab & statusOut & vbTab & errorCode & vbTab & errorMessage
End Function

Public Function RebuildDesignProjectionsBridgeEncoded(ByVal designsWb As Workbook) As String
    Dim report As String
    Dim succeeded As Boolean

    succeeded = modDesignsApply.RebuildDesignProjections(designsWb, report)
    RebuildDesignProjectionsBridgeEncoded = CStr(Abs(CLng(succeeded))) & vbTab & report
End Function

Public Function ListDesignsBridgeResult(Optional ByVal designsWb As Workbook = Nothing, _
                                        Optional ByVal statusFilter As String = "") As Variant
    ListDesignsBridgeResult = modDesignsQueries.ListDesigns(designsWb, statusFilter)
End Function

Public Function GetBOMBridgeResult(ByVal designId As String, ByVal designVersion As String, _
                                   Optional ByVal designsWb As Workbook = Nothing) As Variant
    GetBOMBridgeResult = modDesignsQueries.GetBOM(designId, designVersion, designsWb)
End Function

Public Function DiagnoseDesignsDomainBridgeResult() As String
    DiagnoseDesignsDomainBridgeResult = modDesignsInit.DiagnoseDesignsDomain()
End Function
