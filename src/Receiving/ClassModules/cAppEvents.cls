VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents App As Application

Public Sub Init()
    Set App = Application
End Sub

Private Sub App_NewWorkbook(ByVal Wb As Workbook)
    ' Surface creation is explicit through Receiving ribbon/form entry points.
End Sub

Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
    ' Never mutate a workbook because it was opened while this XLAM is loaded.
End Sub

Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
    ' Never mutate a workbook because it became active.
End Sub

Private Sub App_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    If Not IsReceivingRoleWorkbook(Wb) Then Exit Sub
    modOperatorReadModel.UnregisterAutoSnapshotWorkbook Wb
End Sub

Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    If Not IsReceivingRoleSheet(Sh) Then Exit Sub
    If Not Application.EnableEvents Then Exit Sub
    On Error Resume Next
    modTS_Received.HandleReceivingSelectionChange Target
    On Error GoTo 0
End Sub

Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If Not IsReceivingRoleSheet(Sh) Then Exit Sub
    If Not Application.EnableEvents Then Exit Sub
    On Error Resume Next
    modTS_Received.HandleReceivingSheetChange Target
    On Error GoTo 0
End Sub

Private Function IsReceivingRoleWorkbook(ByVal Wb As Workbook) As Boolean
    On Error GoTo NotRole
    If Wb Is Nothing Then Exit Function
    IsReceivingRoleWorkbook = Not Wb.Worksheets("ReceivedTally").ListObjects("ReceivedTally") Is Nothing
NotRole:
End Function

Private Function IsReceivingRoleSheet(ByVal Sh As Object) As Boolean
    On Error GoTo NotRole
    If Sh Is Nothing Then Exit Function
    If StrComp(CStr(Sh.Name), "ReceivedTally", vbTextCompare) <> 0 Then Exit Function
    IsReceivingRoleSheet = IsReceivingRoleWorkbook(Sh.Parent)
NotRole:
End Function
