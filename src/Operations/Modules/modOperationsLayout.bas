Attribute VB_Name = "modOperationsLayout"
Option Explicit

Public Const OPERATIONS_ANCHOR_LEFT As Long = 1
Public Const OPERATIONS_ANCHOR_TOP As Long = 2
Public Const OPERATIONS_ANCHOR_RIGHT As Long = 4
Public Const OPERATIONS_ANCHOR_BOTTOM As Long = 8

Public Function OperationsAnchorManager() As cOperationsAnchorManager
    Set OperationsAnchorManager = New cOperationsAnchorManager
End Function
