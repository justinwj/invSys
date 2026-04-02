Attribute VB_Name = "TestAddinsPublish"
Option Explicit

Public Function TestVerifyAddinsPublished_AllPresent() As Long
    Dim rootPath As String
    Dim addinsRoot As String

    rootPath = BuildTempRootAddins("verify_all_present")
    addinsRoot = rootPath & "\Addins"

    On Error GoTo CleanFail
    CreateRequiredAddinsSet addinsRoot, False
    modAddinsPublish.SetAddinsPublishSharePointRootOverride rootPath
    modDiagnostics.ResetDiagnosticCapture

    If Not modAddinsPublish.VerifyAddinsPublished() Then GoTo CleanExit
    If StrComp(modAddinsPublish.GetLastAddinsPublishReport(), "OK", vbTextCompare) <> 0 Then GoTo CleanExit
    If modDiagnostics.GetDiagnosticEventCount() <> 0 Then GoTo CleanExit

    TestVerifyAddinsPublished_AllPresent = 1

CleanExit:
    modAddinsPublish.ClearAddinsPublishSharePointRootOverride
    DeleteFolderRecursiveAddinsTest rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestVerifyAddinsPublished_OneMissingLogsDiagnostic() As Long
    Dim rootPath As String
    Dim addinsRoot As String

    rootPath = BuildTempRootAddins("verify_one_missing")
    addinsRoot = rootPath & "\Addins"

    On Error GoTo CleanFail
    CreateRequiredAddinsSet addinsRoot, False
    Kill addinsRoot & "\invSys.Admin.xlam"

    modAddinsPublish.SetAddinsPublishSharePointRootOverride rootPath
    modDiagnostics.ResetDiagnosticCapture

    If modAddinsPublish.VerifyAddinsPublished() Then GoTo CleanExit
    If modDiagnostics.GetDiagnosticEventCount() <= 0 Then GoTo CleanExit
    If InStr(1, modDiagnostics.GetLastDiagnosticMessage(), "invSys.Admin.xlam", vbTextCompare) = 0 Then GoTo CleanExit

    TestVerifyAddinsPublished_OneMissingLogsDiagnostic = 1

CleanExit:
    modAddinsPublish.ClearAddinsPublishSharePointRootOverride
    DeleteFolderRecursiveAddinsTest rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestVerifyAddinsPublished_ZeroByteFileLogsDiagnostic() As Long
    Dim rootPath As String
    Dim addinsRoot As String

    rootPath = BuildTempRootAddins("verify_zero_byte")
    addinsRoot = rootPath & "\Addins"

    On Error GoTo CleanFail
    CreateRequiredAddinsSet addinsRoot, False
    WriteTextFileAddinsTest addinsRoot & "\invSys.Core.xlam", vbNullString

    modAddinsPublish.SetAddinsPublishSharePointRootOverride rootPath
    modDiagnostics.ResetDiagnosticCapture

    If modAddinsPublish.VerifyAddinsPublished() Then GoTo CleanExit
    If modDiagnostics.GetDiagnosticEventCount() <= 0 Then GoTo CleanExit
    If InStr(1, modDiagnostics.GetLastDiagnosticMessage(), "Zero-byte add-in", vbTextCompare) = 0 Then GoTo CleanExit

    TestVerifyAddinsPublished_ZeroByteFileLogsDiagnostic = 1

CleanExit:
    modAddinsPublish.ClearAddinsPublishSharePointRootOverride
    DeleteFolderRecursiveAddinsTest rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestPublishAddins_IdempotentRepublishWritesManifest() As Long
    Dim rootPath As String
    Dim sourceRoot As String
    Dim addinsRoot As String
    Dim manifestPath As String
    Dim manifestText As String

    rootPath = BuildTempRootAddins("publish_idempotent")
    sourceRoot = rootPath & "\source"
    addinsRoot = rootPath & "\share\Addins"
    manifestPath = addinsRoot & "\addins-manifest.json"

    On Error GoTo CleanFail
    CreateRequiredAddinsSet sourceRoot, False
    modAddinsPublish.SetAddinsPublishSharePointRootOverride rootPath & "\share"
    modDiagnostics.ResetDiagnosticCapture

    If Not modAddinsPublish.PublishAddins(sourceRoot) Then GoTo CleanExit
    If Not modAddinsPublish.VerifyAddinsPublished() Then GoTo CleanExit
    If Len(Dir$(manifestPath, vbNormal)) = 0 Then GoTo CleanExit

    manifestText = ReadAllTextAddinsTest(manifestPath)
    If InStr(1, manifestText, """published_utc"":", vbTextCompare) = 0 Then GoTo CleanExit
    If InStr(1, manifestText, """invSys.Core.xlam""", vbTextCompare) = 0 Then GoTo CleanExit
    If InStr(1, manifestText, """size_bytes"":", vbTextCompare) = 0 Then GoTo CleanExit

    If Not modAddinsPublish.PublishAddins(sourceRoot) Then GoTo CleanExit
    If InStr(1, modAddinsPublish.GetLastAddinsPublishReport(), "SKIPPED", vbTextCompare) = 0 Then GoTo CleanExit

    TestPublishAddins_IdempotentRepublishWritesManifest = 1

CleanExit:
    modAddinsPublish.ClearAddinsPublishSharePointRootOverride
    DeleteFolderRecursiveAddinsTest rootPath
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Private Sub CreateRequiredAddinsSet(ByVal folderPath As String, ByVal zeroByte As Boolean)
    EnsureFolderRecursiveAddinsTest folderPath
    WriteTextFileAddinsTest folderPath & "\invSys.Core.xlam", IIf(zeroByte, vbNullString, "core-addin")
    WriteTextFileAddinsTest folderPath & "\invSys.Inventory.Domain.xlam", IIf(zeroByte, vbNullString, "inventory-domain-addin")
    WriteTextFileAddinsTest folderPath & "\invSys.Receiving.xlam", IIf(zeroByte, vbNullString, "receiving-addin")
    WriteTextFileAddinsTest folderPath & "\invSys.Admin.xlam", IIf(zeroByte, vbNullString, "admin-addin")
End Sub

Private Function BuildTempRootAddins(ByVal suffix As String) As String
    BuildTempRootAddins = Environ$("TEMP") & "\invSys_addins_" & suffix & "_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & Format$(CLng(Rnd() * 10000), "0000")
End Function

Private Sub EnsureFolderRecursiveAddinsTest(ByVal folderPath As String)
    Dim fso As Object
    Dim parentPath As String
    Dim slashPos As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    folderPath = Trim$(Replace$(folderPath, "/", "\"))
    If folderPath = "" Then GoTo CleanExit
    If fso.FolderExists(folderPath) Then GoTo CleanExit

    slashPos = InStrRev(folderPath, "\")
    If slashPos > 3 Then
        parentPath = Left$(folderPath, slashPos - 1)
        If Not fso.FolderExists(parentPath) Then EnsureFolderRecursiveAddinsTest parentPath
    End If
    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath

CleanExit:
    Set fso = Nothing
End Sub

Private Sub DeleteFolderRecursiveAddinsTest(ByVal folderPath As String)
    Dim fso As Object

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then
        If fso.FolderExists(folderPath) Then fso.DeleteFolder folderPath, True
    End If
    Set fso = Nothing
    On Error GoTo 0
End Sub

Private Sub WriteTextFileAddinsTest(ByVal filePath As String, ByVal textOut As String)
    Dim fileNum As Integer
    Dim folderPath As String
    Dim slashPos As Long

    slashPos = InStrRev(filePath, "\")
    If slashPos > 0 Then
        folderPath = Left$(filePath, slashPos - 1)
        EnsureFolderRecursiveAddinsTest folderPath
    End If

    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, textOut;
    Close #fileNum
End Sub

Private Function ReadAllTextAddinsTest(ByVal filePath As String) As String
    Dim fileNum As Integer

    If Len(Dir$(filePath, vbNormal)) = 0 Then Exit Function

    fileNum = FreeFile
    Open filePath For Binary Access Read As #fileNum
    ReadAllTextAddinsTest = Space$(LOF(fileNum))
    Get #fileNum, , ReadAllTextAddinsTest
    Close #fileNum
End Function
