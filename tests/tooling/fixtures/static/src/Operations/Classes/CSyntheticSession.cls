VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSyntheticSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mSystemKey As String

Public Property Get SystemKey() As String
    SystemKey = mSystemKey
End Property

Public Property Let SystemKey(ByVal value As String)
    mSystemKey = value
End Property
