Attribute VB_Name = "modOperationsInit"
Option Explicit

Public Sub Auto_Open()
    ' Shadow startup is registration-only. Role launchers remain responsible
    ' for creating or refreshing operator workbook surfaces.
    modReceivingInit.ReceivingPackageAutoOpen
    modProductionInit.ProductionPackageAutoOpen
    modShippingInit.ShippingPackageAutoOpen
End Sub

Public Function OperationsShadowStartupForTest() As String
    On Error GoTo Failed

    Auto_Open
    OperationsShadowStartupForTest = _
        "OK|Receiving=True|Production=True|Shipping=True"
    Exit Function

Failed:
    OperationsShadowStartupForTest = _
        "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function
