Attribute VB_Name = "modSyntheticSurface"
Option Explicit

Public Const TABLE_INVENTORY_ENTITIES As String = "tblInventoryEntities"
Public Const TABLE_LEGACY_VIOLATION As String = "tblLegacyViolation"
Public Const CONFIG_AUTO_REFRESH_SECONDS As String = "AutoRefreshIntervalSeconds"
Public Const EVENT_RECEIVE As String = "RECEIVE"
Public Const CAPABILITY_RECEIVE_POST As String = "RECEIVE_POST"

Public Sub Auto_Open()
    DirectWorker "AUTO"
End Sub

Public Sub RibbonSyntheticOnAction(ByVal control As Object)
    DirectWorker CStr(control.Id)
End Sub

Public Function RibbonSyntheticGetEnabled(ByVal control As Object) As Boolean
    RibbonSyntheticGetEnabled = (Len(CStr(control.Id)) > 0)
End Function

Public Sub InvokeLiteralBridge()
    Application.Run "'invSys.Core.xlam'!SyntheticBridge", "payload"
End Sub

Public Sub InvokeDynamicBridge(ByVal dynamicTarget As String)
    Application.Run dynamicTarget
End Sub

Public Sub SyntheticProcessorHandler(ByVal payload As String)
    DirectWorker payload
End Sub

Public Sub StringDispatchedTarget()
    DirectWorker "STRING"
End Sub

Public Function SyntheticWindowProc( _
    ByVal hwnd As LongPtr, _
    ByVal messageId As Long, _
    ByVal wParam As LongPtr, _
    ByVal lParam As LongPtr) As LongPtr

    SyntheticWindowProc = 0
End Function

Public Sub RetainedCompatibilityShim()
    DirectWorker "COMPAT"
End Sub

Private Sub DirectWorker(ByVal payload As String)
    Debug.Print payload
End Sub

Private Function DuplicateAlpha(ByVal value As Long) As Long
    DuplicateAlpha = value + 1
End Function

Private Function DuplicateBeta(ByVal value As Long) As Long
    DuplicateBeta = value + 1
End Function

Private Sub UnreferencedCandidate()
    Debug.Print "unreachable"
End Sub

Public Function ManagedInventoryHeaders() As Variant
    ManagedInventoryHeaders = Array( _
        "System_Key", _
        "SKU", _
        "QtyOnHand", _
        "Location", _
        "Condition", _
        "Custom_Color")
End Function

Public Function RetiredLegacyHeaders() As Variant
    RetiredLegacyHeaders = Array("ROW", "ITEM_CODE", "Location")
End Function
