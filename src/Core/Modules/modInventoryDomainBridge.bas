Attribute VB_Name = "modInventoryDomainBridge"
Option Explicit

Public Const APPLY_STATUS_APPLIED As String = "APPLIED"
Public Const APPLY_STATUS_SKIP_DUP As String = "SKIP_DUP"

Public Const EVENT_TYPE_RECEIVE As String = "RECEIVE"
Public Const EVENT_TYPE_SHIP As String = "SHIP"
Public Const EVENT_TYPE_PROD_CONSUME As String = "PROD_CONSUME"
Public Const EVENT_TYPE_PROD_COMPLETE As String = "PROD_COMPLETE"

Private Const INVENTORY_DOMAIN_ADDIN_NAME As String = "invSys.Inventory.Domain.xlam"

Public Function ResolveInventoryWorkbookBridge(Optional ByVal warehouseId As String = "", _
                                              Optional ByVal inventoryWb As Workbook = Nothing) As Workbook
    Dim result As Variant

    If Not inventoryWb Is Nothing Then
        Set ResolveInventoryWorkbookBridge = inventoryWb
        Exit Function
    End If

    On Error GoTo FailResolve
    result = RunInventoryDomainMacro1("modInventoryBridgeApi.ResolveInventoryWorkbookBridgeResult", warehouseId)
    If IsObject(result) Then Set ResolveInventoryWorkbookBridge = result
    Exit Function

FailResolve:
    Set ResolveInventoryWorkbookBridge = Nothing
End Function

Public Function EnsureInventorySchemaBridge(Optional ByVal targetWb As Workbook = Nothing, _
                                           Optional ByRef report As String = "") As Boolean
    Dim result As Variant
    Dim payload As Object

    On Error GoTo FailEnsure
    result = RunInventoryDomainMacro1("modInventoryBridgeApi.EnsureInventorySchemaBridgeResult", targetWb)
    If IsObject(result) Then
        Set payload = result
        EnsureInventorySchemaBridge = GetBridgeBool(payload, "Success")
        report = GetBridgeString(payload, "Report")
    End If
    Exit Function

FailEnsure:
    report = Err.Description
    EnsureInventorySchemaBridge = False
End Function

Public Function ApplyInventoryEventBridge(ByVal evt As Object, _
                                         Optional ByVal inventoryWb As Workbook = Nothing, _
                                         Optional ByVal runId As String = "", _
                                         Optional ByRef statusOut As String = "", _
                                         Optional ByRef errorCode As String = "", _
                                         Optional ByRef errorMessage As String = "") As Boolean
    Dim result As Variant
    Dim payload As Object

    On Error GoTo FailApply
    result = RunInventoryDomainMacro3("modInventoryBridgeApi.ApplyEventBridgeResult", evt, inventoryWb, runId)
    If IsObject(result) Then
        Set payload = result
        ApplyInventoryEventBridge = GetBridgeBool(payload, "Success")
        statusOut = GetBridgeString(payload, "StatusOut")
        errorCode = GetBridgeString(payload, "ErrorCode")
        errorMessage = GetBridgeString(payload, "ErrorMessage")
    End If
    Exit Function

FailApply:
    errorCode = "INVENTORY_DOMAIN_CALL_FAILED"
    errorMessage = Err.Description
    ApplyInventoryEventBridge = False
End Function

Public Function RemoveLastBulkLogEntriesBridge(ByVal countToRemove As Long) As Collection
    Dim result As Variant

    On Error GoTo FailRemove
    result = RunInventoryDomainMacro1("modInventoryBridgeApi.RemoveLastBulkLogEntriesBridgeResult", countToRemove)
    If IsObject(result) Then Set RemoveLastBulkLogEntriesBridge = result
    Exit Function

FailRemove:
    Set RemoveLastBulkLogEntriesBridge = New Collection
End Function

Public Sub ReAddBulkLogEntriesBridge(ByVal logDataCollection As Collection)
    On Error Resume Next
    Call RunInventoryDomainMacro1("modInventoryBridgeApi.ReAddBulkLogEntriesBridgeResult", logDataCollection)
    On Error GoTo 0
End Sub

Private Function RunInventoryDomainMacro0(ByVal macroName As String) As Variant
    RunInventoryDomainMacro0 = Application.Run(ResolveInventoryDomainMacroName(macroName))
End Function

Private Function RunInventoryDomainMacro1(ByVal macroName As String, ByVal arg0 As Variant) As Variant
    RunInventoryDomainMacro1 = Application.Run(ResolveInventoryDomainMacroName(macroName), arg0)
End Function

Private Function RunInventoryDomainMacro2(ByVal macroName As String, ByVal arg0 As Variant, ByVal arg1 As Variant) As Variant
    RunInventoryDomainMacro2 = Application.Run(ResolveInventoryDomainMacroName(macroName), arg0, arg1)
End Function

Private Function RunInventoryDomainMacro3(ByVal macroName As String, ByVal arg0 As Variant, ByVal arg1 As Variant, ByVal arg2 As Variant) As Variant
    RunInventoryDomainMacro3 = Application.Run(ResolveInventoryDomainMacroName(macroName), arg0, arg1, arg2)
End Function

Private Function ResolveInventoryDomainMacroName(ByVal macroName As String) As String
    Dim wb As Workbook

    Set wb = FindInventoryDomainAddin()
    If wb Is Nothing Then
        Err.Raise vbObjectError + 2601, "modInventoryDomainBridge.ResolveInventoryDomainMacroName", _
                  "Inventory Domain add-in is not open."
    End If

    ResolveInventoryDomainMacroName = "'" & wb.Name & "'!" & macroName
End Function

Private Function FindInventoryDomainAddin() As Workbook
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If StrComp(wb.Name, INVENTORY_DOMAIN_ADDIN_NAME, vbTextCompare) = 0 Then
            Set FindInventoryDomainAddin = wb
            Exit Function
        End If
    Next wb

    For Each wb In Application.Workbooks
        If InStr(1, wb.Name, "Inventory.Domain", vbTextCompare) > 0 Then
            Set FindInventoryDomainAddin = wb
            Exit Function
        End If
    Next wb
End Function

Private Function GetBridgeString(ByVal payload As Object, ByVal key As String) As String
    On Error Resume Next
    If Not payload Is Nothing Then
        If payload.Exists(key) Then GetBridgeString = CStr(payload(key))
    End If
    On Error GoTo 0
End Function

Private Function GetBridgeBool(ByVal payload As Object, ByVal key As String) As Boolean
    On Error Resume Next
    If Not payload Is Nothing Then
        If payload.Exists(key) Then GetBridgeBool = CBool(payload(key))
    End If
    On Error GoTo 0
End Function
