Attribute VB_Name = "modProductionInit"
Option Explicit

Private gAppEvents As cProductionAppEvents

Public Sub InitProductionAddin()
    Dim prevEvents As Boolean
    Dim prevScreenUpdating As Boolean

    prevEvents = Application.EnableEvents
    prevScreenUpdating = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    InitializeProductionEventHooks
    mProduction.InitializeProductionUiForWorkbook ThisWorkbook
    EnsureProductionSurfaceForWorkbook Application.ActiveWorkbook
    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEvents
End Sub

Public Sub ProductionPackageAutoOpen()
    ' Loading one role XLAM must not create or repair surfaces in whichever
    ' operator workbook happens to be active. Ribbon/form entry points call
    ' InitProductionAddin explicitly when Production is actually requested.
    InitializeProductionEventHooks
End Sub

Private Sub InitializeProductionEventHooks()
    If gAppEvents Is Nothing Then
        Set gAppEvents = New cProductionAppEvents
        gAppEvents.Init
    End If
End Sub

Public Sub EnsureProductionSurfaceForWorkbook(ByVal wb As Workbook)
    Dim prevEvents As Boolean

    If wb Is Nothing Then Exit Sub
    If Not modRoleWorkbookSurfaces.ShouldBootstrapRoleWorkbookSurface(wb) Then Exit Sub
    If Not IsLikelyProductionWorkbook(wb) Then Exit Sub
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    mProduction.InitializeProductionUiForWorkbook wb
    Application.EnableEvents = prevEvents
End Sub

Private Function IsLikelyProductionWorkbook(ByVal wb As Workbook) As Boolean
    Dim wbName As String

    If wb Is Nothing Then Exit Function
    wbName = LCase$(Trim$(wb.Name))
    If wbName Like "*.production.operator.xls*" Then
        IsLikelyProductionWorkbook = True
        Exit Function
    End If
    If WorkbookSheetExistsProductionInit(wb, "Production") _
       And WorkbookSheetExistsProductionInit(wb, "Recipes") _
       And WorkbookTableExistsProductionInit(wb, "RecipeBuilder") Then
        IsLikelyProductionWorkbook = True
    End If
End Function

Private Function WorkbookSheetExistsProductionInit(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    On Error Resume Next
    WorkbookSheetExistsProductionInit = Not wb.Worksheets(sheetName) Is Nothing
    On Error GoTo 0
End Function

Private Function WorkbookTableExistsProductionInit(ByVal wb As Workbook, ByVal tableName As String) As Boolean
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        On Error Resume Next
        WorkbookTableExistsProductionInit = Not ws.ListObjects(tableName) Is Nothing
        On Error GoTo 0
        If WorkbookTableExistsProductionInit Then Exit Function
    Next ws
End Function
