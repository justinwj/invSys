Attribute VB_Name = "TestSyntheticSurface"
Option Explicit

Public Sub RunSyntheticSurfaceTests()
    TestLiteralBridge
    TestSystemKeyHeaders
End Sub

Private Sub TestLiteralBridge()
    modSyntheticSurface.InvokeLiteralBridge
End Sub

Private Sub TestSystemKeyHeaders()
    Dim headers As Variant
    headers = modSyntheticSurface.ManagedInventoryHeaders()
    Debug.Assert headers(0) = "System_Key"
    Debug.Assert headers(4) = "Condition"
End Sub
