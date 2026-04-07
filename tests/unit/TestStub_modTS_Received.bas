Attribute VB_Name = "modTS_Received"
Option Explicit

Private mInitializeCount As Long
Private mLastWorkbookName As String

Public Sub InitializeReceivingUiForWorkbook(Optional ByVal targetWb As Workbook = Nothing)
    mInitializeCount = mInitializeCount + 1
    If Not targetWb Is Nothing Then mLastWorkbookName = targetWb.Name
End Sub

Public Sub ResetReceivingUiStub()
    mInitializeCount = 0
    mLastWorkbookName = vbNullString
End Sub

Public Function GetReceivingUiStubInitializeCount() As Long
    GetReceivingUiStubInitializeCount = mInitializeCount
End Function

Public Function GetReceivingUiStubLastWorkbookName() As String
    GetReceivingUiStubLastWorkbookName = mLastWorkbookName
End Function
