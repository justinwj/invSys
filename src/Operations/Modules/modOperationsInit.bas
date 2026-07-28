Attribute VB_Name = "modOperationsInit"
Option Explicit

Private mStartupReport As String

Public Sub Auto_Open()
    Dim coexistenceReport As String
    Dim roleReport As String

    coexistenceReport = LegacyRoleAddinCoexistenceReport()
    If coexistenceReport <> "" Then
        mStartupReport = "FAIL|LEGACY_ROLE_ADDINS_LOADED|" & coexistenceReport & _
                         "|Close Excel and run tools\register_current_addins.ps1."
        MsgBox "invSys Operations did not initialize because a retired standalone role add-in is also loaded." & _
               vbCrLf & vbCrLf & coexistenceReport & vbCrLf & vbCrLf & _
               "Close Excel and run tools\register_current_addins.ps1.", _
               vbExclamation, "invSys Operations"
        Exit Sub
    End If

    roleReport = InitializeReceivingOperations()
    If Left$(roleReport, 3) <> "OK|" Then GoTo RoleFailed

    roleReport = InitializeProductionOperations()
    If Left$(roleReport, 3) <> "OK|" Then GoTo RoleFailed

    roleReport = InitializeShippingOperations()
    If Left$(roleReport, 3) <> "OK|" Then GoTo RoleFailed

    mStartupReport = "OK|Receiving=True|Production=True|Shipping=True"
    Exit Sub

RoleFailed:
    mStartupReport = roleReport
    MsgBox "invSys Operations loaded, but one role failed to initialize." & _
           vbCrLf & vbCrLf & roleReport, vbExclamation, "invSys Operations"
End Sub

Public Function OperationsShadowStartupForTest() As String
    Auto_Open
    OperationsShadowStartupForTest = mStartupReport
End Function

Public Function OperationsStartupReport() As String
    OperationsStartupReport = mStartupReport
End Function

Public Function LegacyRoleAddinCoexistenceReport() As String
    Dim legacyNames As Variant
    Dim legacyName As Variant
    Dim wb As Workbook
    Dim found As String

    legacyNames = Array( _
        "invSys.Receiving.xlam", _
        "invSys.Production.xlam", _
        "invSys.Shipping.xlam")

    For Each wb In Application.Workbooks
        If Not (wb Is ThisWorkbook) Then
            For Each legacyName In legacyNames
                If StrComp(CStr(wb.Name), CStr(legacyName), vbTextCompare) = 0 Then
                    If found <> "" Then found = found & ","
                    found = found & CStr(legacyName)
                    Exit For
                End If
            Next legacyName
        End If
    Next wb

    If found <> "" Then
        LegacyRoleAddinCoexistenceReport = "Loaded=" & found
    End If
End Function

Private Function InitializeReceivingOperations() As String
    On Error GoTo Failed
    modReceivingInit.ReceivingPackageAutoOpen
    InitializeReceivingOperations = "OK|Role=Receiving"
    Exit Function
Failed:
    InitializeReceivingOperations = _
        "FAIL|Role=Receiving|Error=" & CStr(Err.Number) & "|" & Err.Description
End Function

Private Function InitializeProductionOperations() As String
    On Error GoTo Failed
    modProductionInit.ProductionPackageAutoOpen
    InitializeProductionOperations = "OK|Role=Production"
    Exit Function
Failed:
    InitializeProductionOperations = _
        "FAIL|Role=Production|Error=" & CStr(Err.Number) & "|" & Err.Description
End Function

Private Function InitializeShippingOperations() As String
    On Error GoTo Failed
    modShippingInit.ShippingPackageAutoOpen
    InitializeShippingOperations = "OK|Role=Shipping"
    Exit Function
Failed:
    InitializeShippingOperations = _
        "FAIL|Role=Shipping|Error=" & CStr(Err.Number) & "|" & Err.Description
End Function
