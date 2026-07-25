Attribute VB_Name = "modDesignsRuntime"
Option Explicit

Public Function ResolveDesignsWorkbook(Optional ByVal warehouseId As String = "", _
                                       Optional ByVal designsWb As Workbook = Nothing, _
                                       Optional ByRef report As String = "") As Workbook
    On Error GoTo FailResolve

    Dim targetPath As String
    Dim wb As Workbook

    If Not designsWb Is Nothing Then
        If designsWb.IsAddin Then
            report = "An XLAM cannot be the authoritative Designs workbook."
            Exit Function
        End If
        Set ResolveDesignsWorkbook = designsWb
        Exit Function
    End If

    warehouseId = ResolveDesignsWarehouseId(warehouseId)
    If warehouseId = "" Then
        report = "WarehouseId was not resolved."
        Exit Function
    End If
    targetPath = BuildDesignsWorkbookPath(warehouseId)
    If targetPath = "" Then
        report = "Designs runtime path was not resolved."
        Exit Function
    End If

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, targetPath, vbTextCompare) = 0 Then
            Set ResolveDesignsWorkbook = wb
            Exit Function
        End If
    Next wb

    If Len(Dir$(targetPath)) > 0 Then
        Set wb = Application.Workbooks.Open(targetPath)
    Else
        Set wb = Application.Workbooks.Add(xlWBATWorksheet)
        wb.SaveAs Filename:=targetPath, FileFormat:=50
    End If
    If wb Is Nothing Then
        report = "Authoritative Designs workbook could not be opened or created."
        Exit Function
    End If
    If Not modDesignsSchema.EnsureDesignsSchema(wb, report) Then Exit Function
    If Not wb.ReadOnly Then wb.Save
    Set ResolveDesignsWorkbook = wb
    report = "OK"
    Exit Function

FailResolve:
    report = "ResolveDesignsWorkbook failed: " & Err.Description
End Function

Public Function BuildDesignsWorkbookPath(Optional ByVal warehouseId As String = "") As String
    Dim rootPath As String

    warehouseId = ResolveDesignsWarehouseId(warehouseId)
    If warehouseId = "" Then Exit Function
    rootPath = Trim$(modRuntimeWorkbooks.GetCoreDataRootOverride())
    If rootPath = "" Then rootPath = Trim$(modConfig.GetString("PathDataRoot", ""))
    If rootPath = "" Then rootPath = modDeploymentPaths.DefaultWarehouseRuntimeRootPath(warehouseId, True)
    If rootPath = "" Then Exit Function
    rootPath = Replace$(rootPath, "/", "\")
    If Right$(rootPath, 1) <> "\" Then rootPath = rootPath & "\"
    BuildDesignsWorkbookPath = rootPath & warehouseId & ".invSys.Data.Designs.xlsb"
End Function

Private Function ResolveDesignsWarehouseId(ByVal warehouseId As String) As String
    warehouseId = Trim$(warehouseId)
    If warehouseId = "" And modConfig.IsLoaded() Then warehouseId = Trim$(modConfig.GetWarehouseId())
    ResolveDesignsWarehouseId = warehouseId
End Function
