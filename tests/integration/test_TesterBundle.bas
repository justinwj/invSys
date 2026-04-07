Attribute VB_Name = "test_TesterBundle"
Option Explicit

Private mCaseNames() As String
Private mCaseResults() As String
Private mCaseDetails() As String
Private mCaseCount As Long
Private mSummary As String
Private Const TEST_WAREHOUSE_ID As String = "WHBUND1"

Public Function TestTesterBundle_EndToEnd() As Long
    Dim runtimeBase As String
    Dim runtimeRoot As String
    Dim templateRoot As String
    Dim sharePointRoot As String
    Dim outputRoot As String
    Dim extractRoot As String
    Dim detailText As String
    Dim zipPath As String
    Dim readmePath As String

    On Error GoTo FailTest

    runtimeBase = BuildTesterBundleTempRoot("bundle_e2e")
    runtimeRoot = runtimeBase & "\runtime"
    templateRoot = runtimeBase & "\templates"
    sharePointRoot = runtimeBase & "\sharepoint"
    outputRoot = runtimeBase & "\output"
    extractRoot = runtimeBase & "\extracted"

    ResetTesterBundleEvidence
    If Not SetupTesterBundleRuntime(runtimeRoot, templateRoot, sharePointRoot) Then
        RecordTesterBundleCase "Harness.SetupRuntime", False, "Runtime setup failed: " & modWarehouseBootstrap.GetLastWarehouseBootstrapReport()
        mSummary = "Tester bundle runtime setup failed."
        GoTo CleanExit
    End If

    modRuntimeWorkbooks.SetCoreDataRootOverride runtimeRoot
    modTesterBundle.SetTesterBundleSharePointRootOverride sharePointRoot
    CreateRequiredAddinsForTesterBundle sharePointRoot & "\Addins"

    RecordTesterBundleCase "WriteBundle.Created", modTesterBundle.WriteTesterBundle(TEST_WAREHOUSE_ID, outputRoot), modTesterBundle.GetLastTesterBundleReport()
    zipPath = modTesterBundle.GetLastTesterBundleZipPath()
    readmePath = modTesterBundle.GetLastTesterBundleReadmePath()

    RecordTesterBundleCase "WriteBundle.VerifyPasses", modTesterBundle.VerifyTesterBundle(zipPath), modTesterBundle.GetLastTesterBundleReport()
    RecordTesterBundleCase "WriteBundle.ReadmeSidecar", (Len(Dir$(readmePath, vbNormal)) > 0), readmePath
    RecordTesterBundleCase "WriteBundle.ExtractForInspection", modTesterBundle.ExtractTesterBundleToFolder(zipPath, extractRoot), modTesterBundle.GetLastTesterBundleReport()
    RecordTesterBundleCase "WriteBundle.NoCredentials", AssertNoCredentialsInBundle(extractRoot, detailText), detailText

    RecordTesterBundleCase "PublishBundle.Idempotent", RunPublishIdempotentCase(sharePointRoot, detailText), detailText

    If AllTesterBundleCasesPassed() Then
        mSummary = "Tester bundle was created, verified, sanitized, and published idempotently."
        TestTesterBundle_EndToEnd = 1
    Else
        mSummary = "One or more tester bundle cases failed."
    End If

CleanExit:
    modTesterBundle.ClearTesterBundleSharePointRootOverride
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    CleanupTesterBundleTempRoot runtimeBase
    Exit Function

FailTest:
    RecordTesterBundleCase "Harness.Exception", False, Err.Description
    mSummary = "Tester bundle integration raised an unexpected exception."
    Resume CleanExit
End Function

Public Function GetTesterBundleContextPacked() As String
    GetTesterBundleContextPacked = "Summary=" & SafeTesterBundleText(mSummary)
End Function

Public Function GetTesterBundleEvidenceRows() As String
    Dim i As Long

    For i = 1 To mCaseCount
        If Len(GetTesterBundleEvidenceRows) > 0 Then GetTesterBundleEvidenceRows = GetTesterBundleEvidenceRows & vbLf
        GetTesterBundleEvidenceRows = GetTesterBundleEvidenceRows & _
            mCaseNames(i) & vbTab & mCaseResults(i) & vbTab & mCaseDetails(i)
    Next i
End Function

Private Function SetupTesterBundleRuntime(ByVal runtimeRoot As String, ByVal templateRoot As String, ByVal sharePointRoot As String) As Boolean
    Dim spec As modWarehouseBootstrap.WarehouseSpec

    On Error GoTo FailSetup

    spec.WarehouseId = TEST_WAREHOUSE_ID
    spec.WarehouseName = "Warehouse One"
    spec.StationId = "R1"
    spec.AdminUser = "admin.bundle"
    spec.PathLocal = runtimeRoot
    spec.PathSharePoint = sharePointRoot

    modWarehouseBootstrap.SetWarehouseBootstrapTemplateRootOverride templateRoot
    modRuntimeWorkbooks.SetCoreDataRootOverride runtimeRoot
    If Not modWarehouseBootstrap.BootstrapWarehouseLocal(spec) Then GoTo CleanExit
    SetupTesterBundleRuntime = True

CleanExit:
    modWarehouseBootstrap.ClearWarehouseBootstrapTemplateRootOverride
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    Exit Function

FailSetup:
    Resume CleanExit
End Function

Private Sub CreateRequiredAddinsForTesterBundle(ByVal addinsRoot As String)
    WriteTextTesterBundle addinsRoot & "\invSys.Core.xlam", "core"
    WriteTextTesterBundle addinsRoot & "\invSys.Inventory.Domain.xlam", "inventory"
    WriteTextTesterBundle addinsRoot & "\invSys.Receiving.xlam", "receiving"
    WriteTextTesterBundle addinsRoot & "\invSys.Admin.xlam", "admin"
End Sub

Private Function AssertNoCredentialsInBundle(ByVal extractRoot As String, ByRef detailText As String) As Boolean
    Dim authText As String
    Dim configText As String
    Dim manifestText As String

    authText = ReadAllTextTesterBundle(extractRoot & "\auth\tester-auth-template.csv")
    configText = ReadAllTextTesterBundle(extractRoot & "\config\tblWarehouseConfig.csv")
    manifestText = ReadAllTextTesterBundle(extractRoot & "\manifest.json")

    If authText <> "UserId,WarehouseId,StationId,PasswordHash,Capabilities,Status" & vbCrLf Then
        detailText = "Auth template was not blank headers only."
        Exit Function
    End If
    If InStr(1, configText, "PathSharePointRoot", vbTextCompare) > 0 Then
        detailText = "Config export leaked PathSharePointRoot."
        Exit Function
    End If
    If InStr(1, configText, "admin.bundle", vbTextCompare) > 0 Or InStr(1, manifestText, "admin.bundle", vbTextCompare) > 0 Then
        detailText = "Bundle leaked admin identity."
        Exit Function
    End If
    If InStr(1, authText, "PinHash", vbTextCompare) > 0 Then
        detailText = "Bundle used runtime auth schema instead of blank tester template."
        Exit Function
    End If

    AssertNoCredentialsInBundle = True
    detailText = "Bundle output contained only sanitized config, blank auth headers, and no live credentials."
End Function

Private Function RunPublishIdempotentCase(ByVal sharePointRoot As String, ByRef detailText As String) As Boolean
    Dim bundleTarget As String
    Dim readmeTarget As String
    Dim manifestPath As String
    Dim manifestText As String

    If Not modTesterBundle.PublishTesterBundle(TEST_WAREHOUSE_ID) Then
        detailText = modTesterBundle.GetLastTesterBundleReport()
        Exit Function
    End If
    If Not modTesterBundle.PublishTesterBundle(TEST_WAREHOUSE_ID) Then
        detailText = "Second publish failed: " & modTesterBundle.GetLastTesterBundleReport()
        Exit Function
    End If

    bundleTarget = sharePointRoot & "\TesterPackage\" & TEST_WAREHOUSE_ID & "\" & TEST_WAREHOUSE_ID & ".TesterBundle.zip"
    readmeTarget = sharePointRoot & "\TesterPackage\" & TEST_WAREHOUSE_ID & "\" & TEST_WAREHOUSE_ID & ".TesterReadme.md"
    manifestPath = sharePointRoot & "\Addins\addins-manifest.json"
    manifestText = ReadAllTextTesterBundle(manifestPath)

    If Len(Dir$(bundleTarget, vbNormal)) = 0 Then
        detailText = "Published bundle not found."
        Exit Function
    End If
    If Len(Dir$(readmeTarget, vbNormal)) = 0 Then
        detailText = "Published readme not found."
        Exit Function
    End If
    If InStr(1, manifestText, """tester_bundle_published_utc"":", vbTextCompare) = 0 Then
        detailText = "Add-ins manifest missing tester bundle timestamp."
        Exit Function
    End If
    If InStr(1, manifestText, """tester_bundle_warehouse_id"": """ & TEST_WAREHOUSE_ID & """", vbTextCompare) = 0 Then
        detailText = "Add-ins manifest missing tester bundle warehouse id."
        Exit Function
    End If

    RunPublishIdempotentCase = True
    detailText = "PublishTesterBundle succeeded twice and published the tester bundle, readme, and an updated addins-manifest.json."
End Function

Private Function BuildTesterBundleTempRoot(ByVal suffix As String) As String
    Randomize
    BuildTesterBundleTempRoot = Environ$("TEMP") & "\invSys_testerbundle_integration_" & suffix & "_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & Format$(CLng(Rnd() * 10000), "0000")
End Function

Private Sub CleanupTesterBundleTempRoot(ByVal rootPath As String)
    Dim fso As Object

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then
        If fso.FolderExists(rootPath) Then fso.DeleteFolder rootPath, True
    End If
    Set fso = Nothing
    On Error GoTo 0
End Sub

Private Sub WriteTextTesterBundle(ByVal filePath As String, ByVal textOut As String)
    Dim fileNum As Integer
    Dim parentPath As String
    Dim slashPos As Long

    slashPos = InStrRev(filePath, "\")
    If slashPos > 0 Then
        parentPath = Left$(filePath, slashPos - 1)
        EnsureFolderRecursiveTesterBundle parentPath
    End If

    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, textOut;
    Close #fileNum
End Sub

Private Sub EnsureFolderRecursiveTesterBundle(ByVal folderPath As String)
    Dim fso As Object
    Dim parentPath As String
    Dim slashPos As Long

    folderPath = Trim$(Replace$(folderPath, "/", "\"))
    If folderPath = "" Then Exit Sub

    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(folderPath) Then GoTo CleanExit
    slashPos = InStrRev(folderPath, "\")
    If slashPos > 3 Then
        parentPath = Left$(folderPath, slashPos - 1)
        If Not fso.FolderExists(parentPath) Then EnsureFolderRecursiveTesterBundle parentPath
    End If
    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath

CleanExit:
    Set fso = Nothing
End Sub

Private Function ReadAllTextTesterBundle(ByVal filePath As String) As String
    Dim fileNum As Integer

    If Len(Dir$(filePath, vbNormal)) = 0 Then Exit Function
    fileNum = FreeFile
    Open filePath For Binary Access Read As #fileNum
    ReadAllTextTesterBundle = Space$(LOF(fileNum))
    Get #fileNum, , ReadAllTextTesterBundle
    Close #fileNum
End Function

Private Sub ResetTesterBundleEvidence()
    mCaseCount = 0
    Erase mCaseNames
    Erase mCaseResults
    Erase mCaseDetails
    mSummary = vbNullString
End Sub

Private Sub RecordTesterBundleCase(ByVal caseName As String, ByVal passed As Boolean, ByVal detailText As String)
    mCaseCount = mCaseCount + 1
    ReDim Preserve mCaseNames(1 To mCaseCount)
    ReDim Preserve mCaseResults(1 To mCaseCount)
    ReDim Preserve mCaseDetails(1 To mCaseCount)
    mCaseNames(mCaseCount) = caseName
    mCaseResults(mCaseCount) = IIf(passed, "PASS", "FAIL")
    mCaseDetails(mCaseCount) = SafeTesterBundleText(detailText)
End Sub

Private Function AllTesterBundleCasesPassed() As Boolean
    Dim i As Long

    If mCaseCount = 0 Then Exit Function
    For i = 1 To mCaseCount
        If mCaseResults(i) <> "PASS" Then Exit Function
    Next i
    AllTesterBundleCasesPassed = True
End Function

Private Function SafeTesterBundleText(ByVal textIn As String) As String
    SafeTesterBundleText = Replace$(Replace$(Trim$(textIn), vbCr, " "), vbLf, " ")
End Function
