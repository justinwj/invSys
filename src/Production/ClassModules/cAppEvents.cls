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
    ' Surface creation is explicit through Production ribbon/form entry points.
End Sub

Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
    ' Never mutate a workbook because it was opened while this XLAM is loaded.
End Sub

Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
    ' Never mutate a workbook because it became active.
End Sub

Private Sub App_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    If Not IsProductionRoleWorkbook(Wb) Then Exit Sub
    modOperatorReadModel.UnregisterAutoSnapshotWorkbook Wb
End Sub

Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    If Not IsProductionRoleSheet(Sh) Then Exit Sub
    On Error Resume Next
    mProduction.HandleProductionSelectionChange Target
    On Error GoTo 0
End Sub

Private Sub App_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    If Not IsProductionRoleSheet(Sh) Then Exit Sub
    On Error Resume Next
    mProduction.HandleProductionBeforeDoubleClick Target, Cancel
    On Error GoTo 0
End Sub

Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If Not IsProductionRoleSheet(Sh) Then Exit Sub
    On Error Resume Next
    mProduction.HandleProductionChange Target
    On Error GoTo 0
End Sub

Private Function IsProductionRoleWorkbook(ByVal Wb As Workbook) As Boolean
    On Error GoTo NotRole
    If Wb Is Nothing Then Exit Function
    IsProductionRoleWorkbook = Not Wb.Worksheets("Production").ListObjects("RecipeBuilder") Is Nothing
NotRole:
End Function

Private Function IsProductionRoleSheet(ByVal Sh As Object) As Boolean
    On Error GoTo NotRole
    If Sh Is Nothing Then Exit Function
    If StrComp(CStr(Sh.Name), "Production", vbTextCompare) <> 0 Then Exit Function
    IsProductionRoleSheet = IsProductionRoleWorkbook(Sh.Parent)
NotRole:
End Function
