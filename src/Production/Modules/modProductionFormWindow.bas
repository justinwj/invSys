Attribute VB_Name = "modProductionFormWindow"
Option Explicit

#If Not Mac Then

#If VBA7 Then
    Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Private Declare PtrSafe Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pIUnk As IUnknown, ByRef hwnd As LongPtr) As Long
#Else
    Private Declare Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Private Declare Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pIUnk As IUnknown, ByRef hwnd As Long) As Long
#End If

Private Const GWL_STYLE As Long = -16
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_FRAMECHANGED As Long = &H20
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_MAXIMIZEBOX As Long = &H10000

#End If

Public Function EnableResizable(ByVal productionForm As Object, _
                                Optional ByVal allowMinimize As Boolean = True, _
                                Optional ByVal allowMaximize As Boolean = True) As Boolean
#If Mac Then
    EnableResizable = False
#Else
    Dim hwnd As LongPtr

    On Error GoTo FailEnable
    hwnd = ResolveProductionFormWindowHandle(productionForm)
    If hwnd = 0 Then Exit Function

    EnableResizable = ApplyProductionWindowStyle(hwnd, allowMinimize, allowMaximize)
    Exit Function

FailEnable:
    EnableResizable = False
#End If
End Function

Public Function ApplyDpiLayoutZoom(ByVal productionForm As Object) As Long
    On Error Resume Next
    CallByName productionForm, "Zoom", VbLet, 100
    On Error GoTo 0
    ApplyDpiLayoutZoom = 100
End Function

Public Function DiagnoseWindowStyle(ByVal productionForm As Object) As String
#If Mac Then
    DiagnoseWindowStyle = "Platform=Mac|Supported=False"
#Else
    Dim hwnd As LongPtr

    On Error GoTo FailDiagnostic
    hwnd = ResolveProductionFormWindowHandle(productionForm)
    If hwnd = 0 Then
        DiagnoseWindowStyle = "Handle=False|Resizable=False|Minimize=False|Maximize=False"
        Exit Function
    End If
    DiagnoseWindowStyle = ProductionWindowStyleReport(hwnd)
    Exit Function

FailDiagnostic:
    DiagnoseWindowStyle = "Handle=False|Resizable=False|Minimize=False|Maximize=False|Error=" & Err.Description
#End If
End Function

#If Not Mac Then

Private Function ApplyProductionWindowStyle(ByVal hwnd As LongPtr, _
                                            ByVal allowMinimize As Boolean, _
                                            ByVal allowMaximize As Boolean) As Boolean
    Dim currentStyle As LongPtr
    Dim requestedStyle As LongPtr

    currentStyle = GetWindowLongPtr(hwnd, GWL_STYLE)
    requestedStyle = currentStyle Or WS_THICKFRAME
    If allowMinimize Then requestedStyle = requestedStyle Or WS_MINIMIZEBOX
    If allowMaximize Then requestedStyle = requestedStyle Or WS_MAXIMIZEBOX

    If requestedStyle <> currentStyle Then
        Call SetWindowLongPtr(hwnd, GWL_STYLE, requestedStyle)
        Call SetWindowPos(hwnd, 0, 0, 0, 0, 0, _
                          SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED)
    End If
    ApplyProductionWindowStyle = True
End Function

Private Function ProductionWindowStyleReport(ByVal hwnd As LongPtr) As String
    Dim styleFlags As LongPtr

    styleFlags = GetWindowLongPtr(hwnd, GWL_STYLE)
    ProductionWindowStyleReport = _
        "Handle=True|Resizable=" & CStr((styleFlags And WS_THICKFRAME) <> 0) & _
        "|Minimize=" & CStr((styleFlags And WS_MINIMIZEBOX) <> 0) & _
        "|Maximize=" & CStr((styleFlags And WS_MAXIMIZEBOX) <> 0)
End Function

Private Function ResolveProductionFormWindowHandle(ByVal productionForm As Object) As LongPtr
    On Error Resume Next
    IUnknown_GetWindow productionForm, ResolveProductionFormWindowHandle
    On Error GoTo 0
End Function

#End If
