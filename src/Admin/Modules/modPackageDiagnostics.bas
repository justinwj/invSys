Attribute VB_Name = "modPackageDiagnostics"
Option Explicit

Private Const REPORT_PREFIX As String = "invSys_loaded_package_report"

Private Type PackageDiagnosticContext
    WarehouseId As String
    StationId As String
    PathDataRoot As String
    PathSharePointRoot As String
    ConfigLoaded As Boolean
    ConfigReport As String
End Type

Public Function BuildLoadedPackageReport(Optional ByVal warehouseId As String = "", _
                                         Optional ByVal stationId As String = "") As String
    Dim lines As Collection
    Dim ctx As PackageDiagnosticContext

    On Error GoTo FailBuild

    Set lines = New Collection
    ctx = ResolvePackageDiagnosticContext(warehouseId, stationId)

    AddLinePackageDiagnostic lines, String$(80, "=")
    AddLinePackageDiagnostic lines, "invSys Loaded Package Report " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    AddLinePackageDiagnostic lines, String$(80, "=")

    AppendSessionLinesPackageDiagnostic lines
    AppendConfigLinesPackageDiagnostic lines, ctx
    AppendInstalledAddinsLinesPackageDiagnostic lines
    AppendOpenWorkbookLinesPackageDiagnostic lines
    AppendRuntimeArtifactLinesPackageDiagnostic lines, ctx
    AppendSharePointAddinLinesPackageDiagnostic lines, ctx

    BuildLoadedPackageReport = JoinLinesPackageDiagnostic(lines)
    Exit Function

FailBuild:
    BuildLoadedPackageReport = "BuildLoadedPackageReport failed: " & Err.Description
End Function

Public Function ExportLoadedPackageReport(Optional ByVal outputPath As String = "", _
                                          Optional ByVal warehouseId As String = "", _
                                          Optional ByVal stationId As String = "", _
                                          Optional ByRef actualPathOut As String = "", _
                                          Optional ByRef report As String = "") As Boolean
    Dim textOut As String
    Dim targetPath As String
    Dim fileNo As Integer

    On Error GoTo FailExport

    textOut = BuildLoadedPackageReport(warehouseId, stationId)
    If Len(Trim$(textOut)) = 0 Then
        report = "Loaded package report was empty."
        Exit Function
    End If

    targetPath = ResolveReportPathPackageDiagnostic(outputPath)
    EnsureParentFolderPackageDiagnostic targetPath

    fileNo = FreeFile
    Open targetPath For Output As #fileNo
    Print #fileNo, textOut
    Close #fileNo

    actualPathOut = targetPath
    report = "OK"
    ExportLoadedPackageReport = True
    Exit Function

FailExport:
    On Error Resume Next
    If fileNo <> 0 Then Close #fileNo
    On Error GoTo 0
    report = "ExportLoadedPackageReport failed: " & Err.Description
End Function

Private Function ResolvePackageDiagnosticContext(ByVal warehouseId As String, _
                                                 ByVal stationId As String) As PackageDiagnosticContext
    Dim ctx As PackageDiagnosticContext
    Dim resolvedWh As String
    Dim resolvedSt As String

    resolvedWh = SafeTrimPackageDiagnostic(warehouseId)
    resolvedSt = SafeTrimPackageDiagnostic(stationId)

    If resolvedWh = "" Then resolvedWh = SafeTrimPackageDiagnostic(modConfig.GetWarehouseId())
    If resolvedSt = "" Then resolvedSt = SafeTrimPackageDiagnostic(modConfig.GetStationId())
    If resolvedWh = "" Then resolvedWh = ResolveWarehouseIdFromWorkbookPackageDiagnostic(Application.ActiveWorkbook)
    If resolvedSt = "" Then resolvedSt = ResolveStationIdFromWorkbookPackageDiagnostic(Application.ActiveWorkbook)

    ctx.WarehouseId = resolvedWh
    ctx.StationId = resolvedSt

    If ctx.WarehouseId = "" Then
        ResolvePackageDiagnosticContext = ctx
        Exit Function
    End If

    If modConfig.LoadConfig(ctx.WarehouseId, ctx.StationId) Then
        ctx.ConfigLoaded = True
        ctx.WarehouseId = SafeTrimPackageDiagnostic(modConfig.GetWarehouseId())
        If ctx.WarehouseId = "" Then ctx.WarehouseId = resolvedWh
        ctx.StationId = SafeTrimPackageDiagnostic(modConfig.GetStationId())
        If ctx.StationId = "" Then ctx.StationId = resolvedSt
        ctx.PathDataRoot = NormalizeFolderPathPackageDiagnostic(modConfig.GetString("PathDataRoot", ""), False)
        ctx.PathSharePointRoot = NormalizeFolderPathPackageDiagnostic(modConfig.GetString("PathSharePointRoot", ""), False)
    Else
        ctx.ConfigReport = SafeTrimPackageDiagnostic(modConfig.Validate())
    End If

    ResolvePackageDiagnosticContext = ctx
End Function

Private Sub AppendSessionLinesPackageDiagnostic(ByVal lines As Collection)
    AddLinePackageDiagnostic lines, "Session"
    AddLinePackageDiagnostic lines, "  ComputerName: " & SafeTrimPackageDiagnostic(Environ$("COMPUTERNAME"))
    AddLinePackageDiagnostic lines, "  UserName: " & SafeTrimPackageDiagnostic(Environ$("USERNAME"))
    AddLinePackageDiagnostic lines, "  ExcelVersion: " & SafeTrimPackageDiagnostic(Application.Version)
    AddLinePackageDiagnostic lines, "  ActiveWorkbook: " & SafeWorkbookNamePackageDiagnostic(Application.ActiveWorkbook)
    AddLinePackageDiagnostic lines, "  ActiveWorkbookFullName: " & SafeWorkbookPathPackageDiagnostic(Application.ActiveWorkbook)
    AddLinePackageDiagnostic lines, vbNullString
End Sub

Private Sub AppendConfigLinesPackageDiagnostic(ByVal lines As Collection, _
                                               ByRef ctx As PackageDiagnosticContext)
    AddLinePackageDiagnostic lines, "ConfigContext"
    AddLinePackageDiagnostic lines, "  WarehouseId=" & ctx.WarehouseId & " | StationId=" & ctx.StationId
    AddLinePackageDiagnostic lines, "  ConfigLoaded=" & CStr(ctx.ConfigLoaded)
    If ctx.ConfigReport <> "" Then AddLinePackageDiagnostic lines, "  ConfigReport=" & ctx.ConfigReport
    AddLinePackageDiagnostic lines, "  PathDataRoot=" & ctx.PathDataRoot
    AddLinePackageDiagnostic lines, "  PathSharePointRoot=" & ctx.PathSharePointRoot
    AddLinePackageDiagnostic lines, vbNullString
End Sub

Private Sub AppendInstalledAddinsLinesPackageDiagnostic(ByVal lines As Collection)
    Dim addin As AddIn
    Dim foundAny As Boolean

    AddLinePackageDiagnostic lines, "InstalledAddIns"
    On Error Resume Next
    For Each addin In Application.AddIns
        If ShouldIncludeAddinPackageDiagnostic(addin) Then
            AddLinePackageDiagnostic lines, _
                "  " & SafeTrimPackageDiagnostic(addin.Name) & _
                " | Installed=" & CStr(CBool(addin.Installed)) & _
                " | FullName=" & SafeAddinFullNamePackageDiagnostic(addin)
            foundAny = True
        End If
    Next addin
    On Error GoTo 0

    If Not foundAny Then AddLinePackageDiagnostic lines, "  <none>"
    AddLinePackageDiagnostic lines, vbNullString
End Sub

Private Sub AppendOpenWorkbookLinesPackageDiagnostic(ByVal lines As Collection)
    Dim wb As Workbook
    Dim foundAny As Boolean

    AddLinePackageDiagnostic lines, "OpenInvSysWorkbooks"
    For Each wb In Application.Workbooks
        If ShouldIncludeWorkbookPackageDiagnostic(wb) Then
            AddLinePackageDiagnostic lines, _
                "  " & SafeWorkbookNamePackageDiagnostic(wb) & _
                " | IsAddin=" & CStr(wb.IsAddin) & _
                " | ReadOnly=" & CStr(wb.ReadOnly) & _
                " | FullName=" & SafeWorkbookPathPackageDiagnostic(wb)
            foundAny = True
        End If
    Next wb

    If Not foundAny Then AddLinePackageDiagnostic lines, "  <none>"
    AddLinePackageDiagnostic lines, vbNullString
End Sub

Private Sub AppendRuntimeArtifactLinesPackageDiagnostic(ByVal lines As Collection, _
                                                        ByRef ctx As PackageDiagnosticContext)
    Dim rootPath As String

    AddLinePackageDiagnostic lines, "ExpectedRuntimeArtifacts"
    rootPath = NormalizeFolderPathPackageDiagnostic(ctx.PathDataRoot, False)
    If ctx.WarehouseId = "" Or rootPath = "" Then
        AddLinePackageDiagnostic lines, "  <warehouse/data root unresolved>"
        AddLinePackageDiagnostic lines, vbNullString
        Exit Sub
    End If

    AddArtifactLinePackageDiagnostic lines, "ConfigWorkbook", rootPath & "\" & ctx.WarehouseId & ".invSys.Config.xlsb"
    AddArtifactLinePackageDiagnostic lines, "AuthWorkbook", rootPath & "\" & ctx.WarehouseId & ".invSys.Auth.xlsb"
    AddArtifactLinePackageDiagnostic lines, "InventoryWorkbook", rootPath & "\" & ctx.WarehouseId & ".invSys.Data.Inventory.xlsb"
    AddArtifactLinePackageDiagnostic lines, "SnapshotWorkbook", rootPath & "\" & ctx.WarehouseId & ".invSys.Snapshot.Inventory.xlsb"
    If ctx.StationId <> "" Then
        AddArtifactLinePackageDiagnostic lines, "ReceivingInboxWorkbook", rootPath & "\invSys.Inbox.Receiving." & ctx.StationId & ".xlsb"
    End If
    AddLinePackageDiagnostic lines, vbNullString
End Sub

Private Sub AppendSharePointAddinLinesPackageDiagnostic(ByVal lines As Collection, _
                                                        ByRef ctx As PackageDiagnosticContext)
    Dim addinNames As Variant
    Dim addinName As Variant
    Dim addinsRoot As String
    Dim manifestPath As String

    AddLinePackageDiagnostic lines, "ExpectedSharePointAddins"
    If ctx.PathSharePointRoot = "" Then
        AddLinePackageDiagnostic lines, "  <sharepoint root unresolved>"
        AddLinePackageDiagnostic lines, vbNullString
        Exit Sub
    End If

    addinsRoot = NormalizeFolderPathPackageDiagnostic(ctx.PathSharePointRoot, False) & "\Addins"
    manifestPath = addinsRoot & "\addins-manifest.json"
    AddLinePackageDiagnostic lines, _
        "  AddinsRoot=" & addinsRoot & " | Exists=" & CStr(FolderExistsPackageDiagnostic(addinsRoot))
    AddLinePackageDiagnostic lines, _
        "  Manifest=" & manifestPath & " | Exists=" & CStr(FileExistsPackageDiagnostic(manifestPath))

    addinNames = GetRequiredAddinNamesPackageDiagnostic()
    For Each addinName In addinNames
        AddArtifactLinePackageDiagnostic lines, "SharePointAddin " & CStr(addinName), addinsRoot & "\" & CStr(addinName)
    Next addinName
    AddLinePackageDiagnostic lines, vbNullString
End Sub

Private Sub AddArtifactLinePackageDiagnostic(ByVal lines As Collection, _
                                             ByVal labelText As String, _
                                             ByVal fullPath As String)
    Dim openWb As Workbook

    Set openWb = FindOpenWorkbookByPathPackageDiagnostic(fullPath)
    AddLinePackageDiagnostic lines, _
        "  " & labelText & ": " & fullPath & _
        " | Exists=" & CStr(FileExistsPackageDiagnostic(fullPath)) & _
        " | Open=" & CStr(Not openWb Is Nothing)
End Sub

Private Function GetRequiredAddinNamesPackageDiagnostic() As Variant
    GetRequiredAddinNamesPackageDiagnostic = Array( _
        "invSys.Core.xlam", _
        "invSys.Inventory.Domain.xlam", _
        "invSys.Designs.Domain.xlam", _
        "invSys.Receiving.xlam", _
        "invSys.Shipping.xlam", _
        "invSys.Production.xlam", _
        "invSys.Admin.xlam")
End Function

Private Function ShouldIncludeAddinPackageDiagnostic(ByVal addin As AddIn) As Boolean
    Dim addinName As String
    Dim addinPath As String

    On Error Resume Next
    addinName = LCase$(SafeTrimPackageDiagnostic(addin.Name))
    addinPath = LCase$(SafeTrimPackageDiagnostic(addin.FullName))
    On Error GoTo 0

    ShouldIncludeAddinPackageDiagnostic = _
        (addinName Like "invsys*.xlam") Or _
        (InStr(1, addinName, "invsys", vbTextCompare) > 0) Or _
        (InStr(1, addinPath, "\invsys", vbTextCompare) > 0)
End Function

Private Function ShouldIncludeWorkbookPackageDiagnostic(ByVal wb As Workbook) As Boolean
    Dim wbName As String

    If wb Is Nothing Then Exit Function
    wbName = LCase$(SafeTrimPackageDiagnostic(wb.Name))
    If wbName = "" Then Exit Function

    If wb.IsAddin Then
        ShouldIncludeWorkbookPackageDiagnostic = (InStr(1, wbName, "invsys", vbTextCompare) > 0)
        Exit Function
    End If

    If InStr(1, wbName, "invsys", vbTextCompare) > 0 Then
        ShouldIncludeWorkbookPackageDiagnostic = True
        Exit Function
    End If

    If wbName Like "*.receiving.operator.xls*" _
       Or wbName Like "*.shipping.operator.xls*" _
       Or wbName Like "*.production.operator.xls*" _
       Or wbName Like "*.admin.xls*" Then
        ShouldIncludeWorkbookPackageDiagnostic = True
    End If
End Function

Private Function ResolveWarehouseIdFromWorkbookPackageDiagnostic(ByVal wb As Workbook) As String
    Dim wbName As String
    Dim markerPos As Long

    If wb Is Nothing Then Exit Function
    wbName = SafeWorkbookNamePackageDiagnostic(wb)
    If wbName = "" Then Exit Function

    markerPos = InStr(1, wbName, ".Receiving.Operator.xls", vbTextCompare)
    If markerPos > 1 Then
        ResolveWarehouseIdFromWorkbookPackageDiagnostic = Left$(wbName, markerPos - 1)
        Exit Function
    End If

    markerPos = InStr(1, wbName, ".Shipping.Operator.xls", vbTextCompare)
    If markerPos > 1 Then
        ResolveWarehouseIdFromWorkbookPackageDiagnostic = Left$(wbName, markerPos - 1)
        Exit Function
    End If

    markerPos = InStr(1, wbName, ".Production.Operator.xls", vbTextCompare)
    If markerPos > 1 Then
        ResolveWarehouseIdFromWorkbookPackageDiagnostic = Left$(wbName, markerPos - 1)
        Exit Function
    End If

    markerPos = InStr(1, wbName, ".invSys.", vbTextCompare)
    If markerPos > 1 Then ResolveWarehouseIdFromWorkbookPackageDiagnostic = Left$(wbName, markerPos - 1)
End Function

Private Function ResolveStationIdFromWorkbookPackageDiagnostic(ByVal wb As Workbook) As String
    Dim wbName As String
    Dim namePrefix As String
    Dim parts() As String

    If wb Is Nothing Then Exit Function
    wbName = SafeWorkbookNamePackageDiagnostic(wb)
    If wbName = "" Then Exit Function

    If InStr(1, wbName, "_Receiving_Operator.xls", vbTextCompare) > 0 Then
        namePrefix = Left$(wbName, InStr(1, wbName, "_Receiving_Operator.xls", vbTextCompare) - 1)
        parts = Split(namePrefix, "_")
        If UBound(parts) >= 1 Then ResolveStationIdFromWorkbookPackageDiagnostic = parts(UBound(parts))
    End If
End Function

Private Function FindOpenWorkbookByPathPackageDiagnostic(ByVal fullPath As String) As Workbook
    Dim wb As Workbook

    fullPath = NormalizeFilePathPackageDiagnostic(fullPath)
    If fullPath = "" Then Exit Function

    For Each wb In Application.Workbooks
        If StrComp(NormalizeFilePathPackageDiagnostic(SafeWorkbookPathPackageDiagnostic(wb)), fullPath, vbTextCompare) = 0 Then
            Set FindOpenWorkbookByPathPackageDiagnostic = wb
            Exit Function
        End If
    Next wb
End Function

Private Function ResolveReportPathPackageDiagnostic(ByVal outputPath As String) As String
    If SafeTrimPackageDiagnostic(outputPath) <> "" Then
        ResolveReportPathPackageDiagnostic = NormalizeFilePathPackageDiagnostic(outputPath)
    Else
        ResolveReportPathPackageDiagnostic = NormalizeFilePathPackageDiagnostic(Environ$("TEMP")) & "\" & _
                                             REPORT_PREFIX & "_" & Format$(Now, "yyyymmdd_hhnnss") & ".txt"
    End If
End Function

Private Sub EnsureParentFolderPackageDiagnostic(ByVal filePath As String)
    Dim parentPath As String

    parentPath = GetParentFolderPackageDiagnostic(filePath)
    If parentPath = "" Then Exit Sub
    EnsureFolderRecursivePackageDiagnostic parentPath
End Sub

Private Sub EnsureFolderRecursivePackageDiagnostic(ByVal folderPath As String)
    Dim parentPath As String

    folderPath = NormalizeFolderPathPackageDiagnostic(folderPath, False)
    If folderPath = "" Then Exit Sub
    If FolderExistsPackageDiagnostic(folderPath) Then Exit Sub

    parentPath = GetParentFolderPackageDiagnostic(folderPath)
    If parentPath <> "" And Not FolderExistsPackageDiagnostic(parentPath) Then
        EnsureFolderRecursivePackageDiagnostic parentPath
    End If

    MkDir folderPath
End Sub

Private Function JoinLinesPackageDiagnostic(ByVal lines As Collection) As String
    Dim item As Variant

    If lines Is Nothing Then Exit Function
    For Each item In lines
        If Len(JoinLinesPackageDiagnostic) > 0 Then JoinLinesPackageDiagnostic = JoinLinesPackageDiagnostic & vbCrLf
        JoinLinesPackageDiagnostic = JoinLinesPackageDiagnostic & CStr(item)
    Next item
End Function

Private Sub AddLinePackageDiagnostic(ByVal lines As Collection, ByVal textOut As String)
    If lines Is Nothing Then Exit Sub
    lines.Add CStr(textOut)
End Sub

Private Function GetParentFolderPackageDiagnostic(ByVal pathIn As String) As String
    Dim slashPos As Long

    pathIn = NormalizeFilePathPackageDiagnostic(pathIn)
    If pathIn = "" Then Exit Function
    slashPos = InStrRev(pathIn, "\")
    If slashPos > 1 Then GetParentFolderPackageDiagnostic = Left$(pathIn, slashPos - 1)
End Function

Private Function SafeWorkbookNamePackageDiagnostic(ByVal wb As Workbook) As String
    On Error Resume Next
    SafeWorkbookNamePackageDiagnostic = SafeTrimPackageDiagnostic(wb.Name)
    On Error GoTo 0
End Function

Private Function SafeWorkbookPathPackageDiagnostic(ByVal wb As Workbook) As String
    On Error Resume Next
    SafeWorkbookPathPackageDiagnostic = NormalizeFilePathPackageDiagnostic(wb.FullName)
    On Error GoTo 0
End Function

Private Function SafeAddinFullNamePackageDiagnostic(ByVal addin As AddIn) As String
    On Error Resume Next
    SafeAddinFullNamePackageDiagnostic = NormalizeFilePathPackageDiagnostic(addin.FullName)
    On Error GoTo 0
End Function

Private Function NormalizeFolderPathPackageDiagnostic(ByVal folderPath As String, _
                                                      Optional ByVal withTrailingSlash As Boolean = False) As String
    NormalizeFolderPathPackageDiagnostic = modConfig.NormalizeFolderPathForRuntime(folderPath, withTrailingSlash)
End Function

Private Function NormalizeFilePathPackageDiagnostic(ByVal filePath As String) As String
    filePath = SafeTrimPackageDiagnostic(filePath)
    If filePath = "" Then Exit Function
    NormalizeFilePathPackageDiagnostic = Replace$(filePath, "/", "\")
End Function

Private Function FolderExistsPackageDiagnostic(ByVal folderPath As String) As Boolean
    folderPath = NormalizeFolderPathPackageDiagnostic(folderPath, False)
    If folderPath = "" Then Exit Function
    FolderExistsPackageDiagnostic = (Len(Dir$(folderPath, vbDirectory)) > 0)
End Function

Private Function FileExistsPackageDiagnostic(ByVal filePath As String) As Boolean
    filePath = NormalizeFilePathPackageDiagnostic(filePath)
    If filePath = "" Then Exit Function
    FileExistsPackageDiagnostic = (Len(Dir$(filePath, vbNormal)) > 0)
End Function

Private Function SafeTrimPackageDiagnostic(ByVal valueIn As Variant) As String
    On Error Resume Next
    SafeTrimPackageDiagnostic = Trim$(CStr(valueIn))
    On Error GoTo 0
End Function
