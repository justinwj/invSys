VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cShippingAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents App As Application

Public Sub Init()
    Set App = Application
End Sub

Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
    ' Never mutate a workbook because it was opened while this XLAM is loaded.
End Sub

Private Sub App_NewWorkbook(ByVal Wb As Workbook)
    ' Surface creation is explicit through Shipping ribbon/form entry points.
End Sub

Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
    ' Never mutate a workbook because it became active.
End Sub

Private Sub App_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    If Not IsShippingRoleWorkbook(Wb) Then Exit Sub
    If modUiQuiet.QuietUiIsActive() Then Exit Sub
    modTS_Shipments.HandleShippingOperatorWorkbookClosing Wb
    modOperatorReadModel.UnregisterAutoSnapshotWorkbook Wb
End Sub

Private Sub App_SheetActivate(ByVal Sh As Object)
    If Not IsShippingRoleSheet(Sh) Then Exit Sub
    If modUiQuiet.QuietUiIsActive() Then Exit Sub
End Sub

Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    If Not IsShippingRoleSheet(Sh) Then Exit Sub
    If modUiQuiet.QuietUiIsActive() Then Exit Sub
    On Error Resume Next
    modTS_Shipments.HandleShippingSelectionChange Target
    On Error GoTo 0
End Sub

Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If Not IsShippingRoleSheet(Sh) Then Exit Sub
    If modUiQuiet.QuietUiIsActive() Then Exit Sub
    On Error Resume Next
    modTS_Shipments.HandleShippingSheetChange Target
    On Error GoTo 0
End Sub

Private Function IsShippingRoleWorkbook(ByVal Wb As Workbook) As Boolean
    Dim ws As Worksheet

    On Error GoTo NotRole
    If Wb Is Nothing Then Exit Function
    If Wb.Worksheets("ShipmentsTally") Is Nothing Then Exit Function
    For Each ws In Wb.Worksheets
        On Error Resume Next
        IsShippingRoleWorkbook = Not ws.ListObjects("ShipmentsTally") Is Nothing
        On Error GoTo NotRole
        If IsShippingRoleWorkbook Then Exit Function
    Next ws
NotRole:
End Function

Private Function IsShippingRoleSheet(ByVal Sh As Object) As Boolean
    On Error GoTo NotRole
    If Sh Is Nothing Then Exit Function
    If StrComp(CStr(Sh.Name), "ShipmentsTally", vbTextCompare) <> 0 Then Exit Function
    IsShippingRoleSheet = IsShippingRoleWorkbook(Sh.Parent)
NotRole:
End Function
