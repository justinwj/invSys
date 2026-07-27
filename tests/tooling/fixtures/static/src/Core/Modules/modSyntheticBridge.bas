Attribute VB_Name = "modSyntheticBridge"
Option Explicit

Public Const SYNTHETIC_BRIDGE_CONTRACT_VERSION As String = "1.0.0"

Public Function SyntheticBridge(ByVal payload As String) As String
    SyntheticBridge = "{""ok"":true,""payloadLength"":" & CStr(Len(payload)) & "}"
End Function
