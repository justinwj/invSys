Attribute VB_Name = "modOperationsPrimitiveBridge"
Option Explicit

Public Function EnsureProductionWorkbookSurface(ByVal workbookName As String, _
                                                ByRef report As String) As Boolean
    Dim wb As Workbook

    Set wb = ResolveOpenWorkbook(workbookName)
    If wb Is Nothing Then
        report = "Production operator workbook is not open: " & Trim$(workbookName)
        Exit Function
    End If
    EnsureProductionWorkbookSurface = _
        modRoleWorkbookSurfaces.EnsureProductionWorkbookSurface(wb, report)
End Function

Public Function ShouldBootstrapRoleWorkbookSurface(ByVal workbookName As String) As Boolean
    Dim wb As Workbook

    Set wb = ResolveOpenWorkbook(workbookName)
    If wb Is Nothing Then Exit Function
    ShouldBootstrapRoleWorkbookSurface = _
        modRoleWorkbookSurfaces.ShouldBootstrapRoleWorkbookSurface(wb)
End Function

Public Function BeginQuietUiForWorkbook(ByVal workbookName As String) As Boolean
    Dim wb As Workbook

    Set wb = ResolveOpenWorkbook(workbookName)
    If wb Is Nothing Then Exit Function
    modUiQuiet.BeginQuietUi wb
    BeginQuietUiForWorkbook = True
End Function

Public Function InitializeProductionAutoSnapshot(ByVal workbookName As String) As Boolean
    Dim wb As Workbook

    Set wb = ResolveOpenWorkbook(workbookName)
    If wb Is Nothing Then Exit Function
    modOperatorReadModel.InitializeAutoSnapshotForWorkbook wb
    InitializeProductionAutoSnapshot = True
End Function

Public Sub UnregisterProductionAutoSnapshot(ByVal workbookName As String)
    Dim wb As Workbook

    Set wb = ResolveOpenWorkbook(workbookName)
    If wb Is Nothing Then Exit Sub
    modOperatorReadModel.UnregisterAutoSnapshotWorkbook wb
End Sub

Public Function RefreshInventoryReadModel(ByVal workbookName As String, _
                                          Optional ByVal warehouseId As String = "", _
                                          Optional ByVal sourceType As String = "LOCAL", _
                                          Optional ByRef report As String = "") As Boolean
    Dim wb As Workbook

    Set wb = ResolveOpenWorkbook(workbookName)
    If wb Is Nothing Then
        report = "Operator workbook is not open: " & Trim$(workbookName)
        Exit Function
    End If
    RefreshInventoryReadModel = _
        modOperatorReadModel.RefreshInventoryReadModelForWorkbook( _
            wb, warehouseId, sourceType, report)
End Function

Public Function DiagnoseInventoryReadModel(ByVal workbookName As String, _
                                           Optional ByVal warehouseId As String = "", _
                                           Optional ByVal sourceType As String = "LOCAL") As String
    Dim wb As Workbook

    Set wb = ResolveOpenWorkbook(workbookName)
    If wb Is Nothing Then
        DiagnoseInventoryReadModel = _
            "Operator workbook is not open: " & Trim$(workbookName)
        Exit Function
    End If
    DiagnoseInventoryReadModel = _
        modOperatorReadModel.DiagnoseInventoryReadModelRefresh( _
            wb, warehouseId, sourceType)
End Function

Public Function RunBatchAndRefreshOperatorWorkbook(ByVal workbookName As String, _
                                                   Optional ByVal warehouseId As String = "", _
                                                   Optional ByVal sourceType As String = "LOCAL", _
                                                   Optional ByRef report As String = "", _
                                                   Optional ByVal requireQueuedWork As Boolean = True, _
                                                   Optional ByRef queuedWorkHandled As Boolean = False) As Boolean
    Dim wb As Workbook

    Set wb = ResolveOpenWorkbook(workbookName)
    If wb Is Nothing Then
        report = "Operator workbook is not open: " & Trim$(workbookName)
        Exit Function
    End If
    RunBatchAndRefreshOperatorWorkbook = _
        modOperatorReadModel.RunBatchAndRefreshOperatorWorkbook( _
            wb, warehouseId, sourceType, report, requireQueuedWork, queuedWorkHandled)
End Function

Public Sub ApplyShapeCapability(ByVal workbookName As String, _
                                ByVal worksheetName As String, _
                                ByVal shapeName As String, _
                                ByVal capability As String)
    Dim wb As Workbook
    Dim ws As Worksheet

    Set wb = ResolveOpenWorkbook(workbookName)
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    Set ws = wb.Worksheets(worksheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    modRoleUiAccess.ApplyShapeCapability ws, shapeName, capability
End Sub

Public Function ListDesigns(Optional ByVal statusFilter As String = "") As Variant
    ListDesigns = modDesignsDomainBridge.ListDesignsBridge(Nothing, statusFilter)
End Function

Public Function GetDesignBom(ByVal designId As String, _
                             ByVal designVersion As String) As Variant
    GetDesignBom = _
        modDesignsDomainBridge.GetDesignBOMBridge(designId, designVersion, Nothing)
End Function

Public Function GetDesignBomForStatus(ByVal designId As String, _
                                      ByVal designVersion As String, _
                                      ByVal requiredStatus As String) As Variant
    GetDesignBomForStatus = _
        modDesignsDomainBridge.GetDesignBOMForStatusBridge( _
            designId, designVersion, requiredStatus, Nothing)
End Function

Private Function ResolveOpenWorkbook(ByVal workbookName As String) As Workbook
    Dim wb As Workbook

    workbookName = Trim$(workbookName)
    If workbookName = "" Then Exit Function

    On Error Resume Next
    Set ResolveOpenWorkbook = Application.Workbooks(workbookName)
    On Error GoTo 0
    If Not ResolveOpenWorkbook Is Nothing Then Exit Function

    For Each wb In Application.Workbooks
        On Error Resume Next
        If StrComp(wb.FullName, workbookName, vbTextCompare) = 0 Then
            Set ResolveOpenWorkbook = wb
            Exit Function
        End If
        On Error GoTo 0
    Next wb
End Function
