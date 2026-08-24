VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cProductionAppEvents"
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
    If Not modOperationsInit.WorkbookOwnsRoleTable( _
            Wb, "Production", "RecipeBuilder") Then Exit Sub
    mProduction.HandleProductionOperatorWorkbookClosing Wb
    modOperationsPrimitiveBridge.UnregisterProductionAutoSnapshot Wb.Name
End Sub

Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next
    If modProductionProcessWorksheet.IsProcessWorksheetItemSearchTarget(Target) Then
        mProduction.ShowProductionProcessItemSearch Target
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    If Not modOperationsInit.IsRoleWorksheet( _
            Sh, "Production", "RecipeBuilder") Then Exit Sub
    On Error Resume Next
    mProduction.HandleProductionSelectionChange Target
    On Error GoTo 0
End Sub

Private Sub App_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    If Not modOperationsInit.IsRoleWorksheet( _
            Sh, "Production", "RecipeBuilder") Then Exit Sub
    On Error Resume Next
    mProduction.HandleProductionBeforeDoubleClick Target, Cancel
    On Error GoTo 0
End Sub

Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next
    If modProductionProcessWorksheet.IsProcessWorksheetTableTarget(Target) Then
        modProductionProcessWorksheet.RefreshProcessWorksheetManagedColumns Target
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    If Not modOperationsInit.IsRoleWorksheet( _
            Sh, "Production", "RecipeBuilder") Then Exit Sub
    On Error Resume Next
    mProduction.HandleProductionChange Target
    On Error GoTo 0
End Sub
