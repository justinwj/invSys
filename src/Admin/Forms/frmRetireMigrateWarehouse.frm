VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRetireMigrateWarehouse 
   Caption         =   "Retire / Migrate Warehouse"
   ClientHeight    =   4200
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   6400
   OleObjectBlob   =   "frmRetireMigrateWarehouse.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRetireMigrateWarehouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const PANEL_SELECTION As String = "SELECTION"
Private Const PANEL_CONFIRM As String = "CONFIRM"
Private Const PANEL_RESULT As String = "RESULT"

Private Const COLOR_ERROR As Long = 255
Private Const COLOR_SUCCESS As Long = 32768
Private Const COLOR_INFO As Long = 0
Private Const COLOR_WARNING As Long = 192

Private mFormBusy As Boolean
Private mCurrentPanel As String
Private mPendingSpec As modWarehouseRetire.RetireMigrateSpec
Private mReAuthPassed As Boolean

Private Sub UserForm_Initialize()
    mFormBusy = True
    Me.Caption = "Retire / Migrate Warehouse"
    Me.StartUpPosition = 1

    Me.optArchiveOnly.Value = True
    Me.chkPublishTombstone.Value = True
    Me.chkConfirmAction.Value = False
    Me.lblDeleteWarning.ForeColor = COLOR_ERROR

    ClearAllInlineErrors
    PopulateWarehouseDropdowns
    ApplyDefaultSelections
    ShowSelectionPanel
    ShowFormMessage "Select a source warehouse and operation mode, then click OK.", COLOR_INFO

    mFormBusy = False
End Sub

Private Sub cmbSourceWarehouse_Change()
    If mFormBusy Then Exit Sub
    ClearInlineError Me.lblSourceWarehouseError
    SuggestArchiveDestination False
End Sub

Private Sub cmbTargetWarehouse_Change()
    If mFormBusy Then Exit Sub
    ClearInlineError Me.lblTargetWarehouseError
End Sub

Private Sub txtArchiveDestPath_Change()
    If mFormBusy Then Exit Sub
    ClearInlineError Me.lblArchiveDestPathError
End Sub

Private Sub optArchiveOnly_Click()
    If mFormBusy Then Exit Sub
    UpdateModeUi
End Sub

Private Sub optArchiveMigrate_Click()
    If mFormBusy Then Exit Sub
    UpdateModeUi
End Sub

Private Sub optArchiveRetire_Click()
    If mFormBusy Then Exit Sub
    UpdateModeUi
End Sub

Private Sub optArchiveRetireDelete_Click()
    If mFormBusy Then Exit Sub
    UpdateModeUi
End Sub

Private Sub chkConfirmAction_Click()
    If mCurrentPanel <> PANEL_CONFIRM Then Exit Sub
    ClearInlineError Me.lblConfirmError
    UpdateConfirmOkState
End Sub

Private Sub btnBack_Click()
    If mCurrentPanel = PANEL_CONFIRM Then
        ShowSelectionPanel
        ShowFormMessage "Adjust the selection, then click OK to continue.", COLOR_INFO
    End If
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()
    Select Case mCurrentPanel
        Case PANEL_SELECTION
            HandleSelectionOk
        Case PANEL_CONFIRM
            HandleConfirmOk
        Case PANEL_RESULT
            Unload Me
    End Select
End Sub

Private Sub HandleSelectionOk()
    Dim spec As modWarehouseRetire.RetireMigrateSpec

    ClearAllInlineErrors
    If Not BuildSpecFromSelection(spec) Then
        ShowFormMessage "Fix the highlighted fields and try again.", COLOR_ERROR
        Exit Sub
    End If

    If Not modWarehouseRetire.RequireReAuth("ADMIN_MAINT") Then
        SetInlineError Me.lblReAuthError, "Re-authentication required to continue"
        ShowFormMessage "Re-authentication required to continue.", COLOR_ERROR
        Exit Sub
    End If

    mPendingSpec = spec
    mReAuthPassed = True
    ShowConfirmPanel
End Sub

Private Sub HandleConfirmOk()
    Dim spec As modWarehouseRetire.RetireMigrateSpec
    Dim summaryText As String
    Dim failureText As String

    ClearInlineError Me.lblConfirmError
    If Not mReAuthPassed Then
        SetInlineError Me.lblConfirmError, "Re-authentication required to continue"
        Exit Sub
    End If
    If Not CBool(Me.chkConfirmAction.Value) Then
        SetInlineError Me.lblConfirmError, "You must confirm this action before continuing."
        Exit Sub
    End If

    spec = mPendingSpec
    spec.ConfirmedByUser = True
    If Not modWarehouseRetire.ValidateRetireMigrateSpec(spec, failureText) Then
        SetInlineError Me.lblConfirmError, failureText
        Exit Sub
    End If

    If Not modWarehouseRetire.WriteArchivePackage(spec) Then
        ShowResultPanel False, "WriteArchivePackage failed: " & modWarehouseRetire.GetLastWarehouseRetireReport()
        Exit Sub
    End If
    summaryText = "WriteArchivePackage: " & modWarehouseRetire.GetLastWarehouseRetireReport()

    If spec.OperationMode = modWarehouseRetire.MODE_ARCHIVE_MIGRATE Then
        If Not modWarehouseRetire.MigrateInventoryToTarget(spec) Then
            ShowResultPanel False, "MigrateInventoryToTarget failed: " & modWarehouseRetire.GetLastWarehouseRetireReport()
            Exit Sub
        End If
        summaryText = summaryText & vbCrLf & "MigrateInventoryToTarget: " & modWarehouseRetire.GetLastWarehouseRetireReport()
    End If

    If spec.OperationMode = modWarehouseRetire.MODE_ARCHIVE_RETIRE Or _
       spec.OperationMode = modWarehouseRetire.MODE_ARCHIVE_RETIRE_DELETE Then
        If Not modWarehouseRetire.RetireSourceWarehouse(spec) Then
            ShowResultPanel False, "RetireSourceWarehouse failed: " & modWarehouseRetire.GetLastWarehouseRetireReport()
            Exit Sub
        End If
        summaryText = summaryText & vbCrLf & "RetireSourceWarehouse: " & modWarehouseRetire.GetLastWarehouseRetireReport()
    End If

    If spec.OperationMode = modWarehouseRetire.MODE_ARCHIVE_RETIRE_DELETE Then
        If Not modWarehouseRetire.DeleteLocalRuntime(spec) Then
            ShowResultPanel False, "DeleteLocalRuntime failed: " & modWarehouseRetire.GetLastWarehouseRetireReport()
            Exit Sub
        End If
        summaryText = summaryText & vbCrLf & "DeleteLocalRuntime: " & modWarehouseRetire.GetLastWarehouseRetireReport()
    End If

    ShowResultPanel True, summaryText
End Sub

Private Function BuildSpecFromSelection(ByRef spec As modWarehouseRetire.RetireMigrateSpec) As Boolean
    Dim isValid As Boolean

    spec.SourceWarehouseId = Trim$(CStr(Me.cmbSourceWarehouse.Value))
    spec.TargetWarehouseId = Trim$(CStr(Me.cmbTargetWarehouse.Value))
    spec.OperationMode = ResolveSelectedMode()
    spec.AdminUser = ResolveCurrentUserForm()
    spec.ConfirmedByUser = False
    spec.ArchiveDestPath = Trim$(CStr(Me.txtArchiveDestPath.Value))
    spec.PublishTombstone = CBool(Me.chkPublishTombstone.Value)

    If spec.SourceWarehouseId = "" Then
        SetInlineError Me.lblSourceWarehouseError, "Source warehouse is required."
        isValid = False
    Else
        isValid = True
    End If

    If spec.OperationMode = modWarehouseRetire.MODE_ARCHIVE_MIGRATE Then
        If spec.TargetWarehouseId = "" Then
            SetInlineError Me.lblTargetWarehouseError, "Target warehouse is required for migrate mode."
            isValid = False
        ElseIf StrComp(spec.SourceWarehouseId, spec.TargetWarehouseId, vbTextCompare) = 0 Then
            SetInlineError Me.lblTargetWarehouseError, "Target warehouse must be different from the source."
            isValid = False
        End If
    End If

    If spec.ArchiveDestPath = "" Then
        SetInlineError Me.lblArchiveDestPathError, "Archive destination path is required."
        isValid = False
    End If

    If isValid Then
        If Not ValidateSelectionSpecForm(spec) Then
            isValid = False
        End If
    End If

    BuildSpecFromSelection = isValid
End Function

Private Function ValidateSelectionSpecForm(ByRef spec As modWarehouseRetire.RetireMigrateSpec) As Boolean
    Dim report As String
    Dim confirmedSpec As modWarehouseRetire.RetireMigrateSpec

    confirmedSpec = spec
    confirmedSpec.ConfirmedByUser = True

    If modWarehouseRetire.ValidateRetireMigrateSpec(confirmedSpec, report) Then
        ValidateSelectionSpecForm = True
        Exit Function
    End If

    If InStr(1, report, "SourceWarehouseId", vbTextCompare) > 0 Then
        SetInlineError Me.lblSourceWarehouseError, report
    ElseIf InStr(1, report, "TargetWarehouseId", vbTextCompare) > 0 Then
        SetInlineError Me.lblTargetWarehouseError, report
    ElseIf InStr(1, report, "ArchiveDestPath", vbTextCompare) > 0 Then
        SetInlineError Me.lblArchiveDestPathError, report
    Else
        ShowFormMessage report, COLOR_ERROR
    End If
End Function

Private Sub PopulateWarehouseDropdowns()
    Dim warehouseIds As Collection
    Dim item As Variant

    Set warehouseIds = DiscoverWarehouseIdsForm()
    Me.cmbSourceWarehouse.Clear
    Me.cmbTargetWarehouse.Clear

    For Each item In warehouseIds
        Me.cmbSourceWarehouse.AddItem CStr(item)
        Me.cmbTargetWarehouse.AddItem CStr(item)
    Next item
End Sub

Private Function DiscoverWarehouseIdsForm() As Collection
    Dim results As Collection
    Dim seen As Object
    Dim rootPath As String
    Dim scanRoots As Collection
    Dim candidate As Variant

    Set results = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    Set scanRoots = New Collection

    rootPath = ResolveWarehouseScanRootForm()
    If rootPath <> "" Then scanRoots.Add rootPath
    If StrComp(rootPath, "C:\invSys", vbTextCompare) <> 0 Then scanRoots.Add "C:\invSys"

    For Each candidate In scanRoots
        AddWarehousesFromRootForm results, seen, CStr(candidate)
    Next candidate

    Set DiscoverWarehouseIdsForm = results
End Function

Private Function ResolveWarehouseScanRootForm() As String
    Dim rootPath As String
    Dim parentPath As String

    rootPath = Trim$(modRuntimeWorkbooks.GetCoreDataRootOverride())
    If rootPath = "" Then rootPath = Trim$(modRuntimeWorkbooks.ResolveCoreDataRoot("", ""))
    rootPath = NormalizePathForm(rootPath)
    If rootPath = "" Then
        ResolveWarehouseScanRootForm = "C:\invSys"
        Exit Function
    End If

    parentPath = GetParentFolderForm(rootPath)
    If parentPath = "" Then
        ResolveWarehouseScanRootForm = rootPath
    Else
        ResolveWarehouseScanRootForm = parentPath
    End If
End Function

Private Sub AddWarehousesFromRootForm(ByVal results As Collection, ByVal seen As Object, ByVal rootPath As String)
    Dim folderName As String
    Dim folderPath As String

    rootPath = NormalizePathForm(rootPath)
    If rootPath = "" Then Exit Sub
    If Len(Dir$(rootPath, vbDirectory)) = 0 Then Exit Sub

    folderName = Dir$(rootPath & "\*", vbDirectory)
    Do While folderName <> ""
        If folderName <> "." And folderName <> ".." Then
            folderPath = rootPath & "\" & folderName
            If (GetAttr(folderPath) And vbDirectory) = vbDirectory Then
                If Len(Dir$(folderPath & "\" & folderName & ".invSys.Config.xlsb", vbNormal)) > 0 Then
                    If Not seen.Exists(folderName) Then
                        seen.Add folderName, True
                        results.Add folderName
                    End If
                End If
            End If
        End If
        folderName = Dir$
    Loop
End Sub

Private Sub ApplyDefaultSelections()
    If Me.cmbSourceWarehouse.ListCount > 0 Then
        Me.cmbSourceWarehouse.ListIndex = 0
    End If
    If Me.cmbTargetWarehouse.ListCount > 0 Then
        Me.cmbTargetWarehouse.ListIndex = 0
    End If
    SuggestArchiveDestination True
    UpdateModeUi
End Sub

Private Sub SuggestArchiveDestination(ByVal forceApply As Boolean)
    Dim suggestedPath As String

    suggestedPath = ResolveArchiveDefaultForm(Trim$(CStr(Me.cmbSourceWarehouse.Value)))
    If forceApply Or Trim$(CStr(Me.txtArchiveDestPath.Value)) = "" Then
        Me.txtArchiveDestPath.Value = suggestedPath
    End If
End Sub

Private Function ResolveArchiveDefaultForm(ByVal warehouseId As String) As String
    Dim priorRoot As String
    Dim pathValue As String

    If warehouseId = "" Then
        ResolveArchiveDefaultForm = "C:\invSys\Archive"
        Exit Function
    End If

    priorRoot = modRuntimeWorkbooks.GetCoreDataRootOverride()
    On Error Resume Next
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    If modConfig.LoadConfig(warehouseId, "") Then
        pathValue = Trim$(modConfig.GetString("PathBackupRoot", ""))
    End If
    On Error GoTo 0
    RestoreRootOverrideForm priorRoot

    pathValue = NormalizePathForm(pathValue)
    If pathValue = "" Then
        ResolveArchiveDefaultForm = "C:\invSys\Archive"
    Else
        ResolveArchiveDefaultForm = pathValue
    End If
End Function

Private Sub UpdateModeUi()
    Dim migrateMode As Boolean
    Dim retireMode As Boolean
    Dim deleteMode As Boolean

    migrateMode = (ResolveSelectedMode() = modWarehouseRetire.MODE_ARCHIVE_MIGRATE)
    retireMode = (ResolveSelectedMode() = modWarehouseRetire.MODE_ARCHIVE_RETIRE Or ResolveSelectedMode() = modWarehouseRetire.MODE_ARCHIVE_RETIRE_DELETE)
    deleteMode = (ResolveSelectedMode() = modWarehouseRetire.MODE_ARCHIVE_RETIRE_DELETE)

    Me.cmbTargetWarehouse.Enabled = migrateMode
    Me.lblTargetWarehouse.Enabled = migrateMode
    Me.chkPublishTombstone.Visible = retireMode
    Me.chkPublishTombstone.Enabled = retireMode
    If Not retireMode Then Me.chkPublishTombstone.Value = False

    Me.lblDeleteWarning.Visible = deleteMode
    ClearInlineError Me.lblTargetWarehouseError
    ClearInlineError Me.lblReAuthError
End Sub

Private Function ResolveSelectedMode() As modWarehouseRetire.RetireMigrateOperationMode
    If CBool(Me.optArchiveMigrate.Value) Then
        ResolveSelectedMode = modWarehouseRetire.MODE_ARCHIVE_MIGRATE
    ElseIf CBool(Me.optArchiveRetire.Value) Then
        ResolveSelectedMode = modWarehouseRetire.MODE_ARCHIVE_RETIRE
    ElseIf CBool(Me.optArchiveRetireDelete.Value) Then
        ResolveSelectedMode = modWarehouseRetire.MODE_ARCHIVE_RETIRE_DELETE
    Else
        ResolveSelectedMode = modWarehouseRetire.MODE_ARCHIVE_ONLY
    End If
End Function

Private Sub ShowSelectionPanel()
    mCurrentPanel = PANEL_SELECTION
    SetSelectionControlsVisible True
    Me.fraConfirm.Visible = False
    Me.fraResult.Visible = False
    Me.btnBack.Visible = False
    Me.btnCancel.Caption = "Cancel"
    Me.btnOK.Caption = "OK"
    UpdateModeUi
End Sub

Private Sub ShowConfirmPanel()
    mCurrentPanel = PANEL_CONFIRM
    SetSelectionControlsVisible False
    Me.fraConfirm.Visible = True
    Me.fraResult.Visible = False
    Me.btnBack.Visible = True
    Me.btnCancel.Caption = "Cancel"
    Me.btnOK.Caption = "Run"
    Me.chkConfirmAction.Value = False
    Me.lblConfirmSummary.Caption = BuildConfirmationSummaryForm(mPendingSpec)
    Me.lblDeleteWarning.Visible = (mPendingSpec.OperationMode = modWarehouseRetire.MODE_ARCHIVE_RETIRE_DELETE)
    ClearInlineError Me.lblConfirmError
    UpdateConfirmOkState
End Sub

Private Sub ShowResultPanel(ByVal wasSuccessful As Boolean, ByVal detailText As String)
    mCurrentPanel = PANEL_RESULT
    SetSelectionControlsVisible False
    Me.fraConfirm.Visible = False
    Me.fraResult.Visible = True
    Me.btnBack.Visible = False
    Me.btnCancel.Caption = "Close"
    Me.btnOK.Caption = "Close"
    Me.lblResultSummary.Caption = Trim$(detailText)
    Me.lblResultSummary.ForeColor = IIf(wasSuccessful, COLOR_SUCCESS, COLOR_ERROR)
End Sub

Private Function BuildConfirmationSummaryForm(ByRef spec As modWarehouseRetire.RetireMigrateSpec) As String
    Select Case spec.OperationMode
        Case modWarehouseRetire.MODE_ARCHIVE_ONLY
            BuildConfirmationSummaryForm = _
                "Archive only will create an archive package for " & spec.SourceWarehouseId & "." & vbCrLf & _
                "No migration, retirement, or deletion will occur." & vbCrLf & _
                "Archive destination: " & spec.ArchiveDestPath
        Case modWarehouseRetire.MODE_ARCHIVE_MIGRATE
            BuildConfirmationSummaryForm = _
                "Archive + Migrate will archive " & spec.SourceWarehouseId & " and seed current inventory into " & spec.TargetWarehouseId & "." & vbCrLf & _
                "The target remains locally authoritative. No auth, config identity, or inbox files are copied." & vbCrLf & _
                "Archive destination: " & spec.ArchiveDestPath
        Case modWarehouseRetire.MODE_ARCHIVE_RETIRE
            BuildConfirmationSummaryForm = _
                "Archive + Retire will archive " & spec.SourceWarehouseId & ", mark it RETIRED locally, and write a tombstone." & vbCrLf & _
                IIf(spec.PublishTombstone, "A best-effort SharePoint tombstone publish will also be attempted.", "SharePoint tombstone publish is disabled.") & vbCrLf & _
                "Archive destination: " & spec.ArchiveDestPath
        Case modWarehouseRetire.MODE_ARCHIVE_RETIRE_DELETE
            BuildConfirmationSummaryForm = _
                "Archive + Retire + Delete will archive " & spec.SourceWarehouseId & ", mark it RETIRED, write a tombstone, then delete the local runtime folder." & vbCrLf & _
                IIf(spec.PublishTombstone, "A best-effort SharePoint tombstone publish will also be attempted before deletion.", "SharePoint tombstone publish is disabled.") & vbCrLf & _
                "Archive destination: " & spec.ArchiveDestPath
    End Select
End Function

Private Sub UpdateConfirmOkState()
    Me.btnOK.Enabled = CBool(Me.chkConfirmAction.Value)
End Sub

Private Sub SetSelectionControlsVisible(ByVal isVisible As Boolean)
    Me.lblTitle.Visible = isVisible
    Me.lblSelectionIntro.Visible = isVisible
    Me.lblSourceWarehouse.Visible = isVisible
    Me.cmbSourceWarehouse.Visible = isVisible
    Me.lblSourceWarehouseError.Visible = isVisible
    Me.lblTargetWarehouse.Visible = isVisible
    Me.cmbTargetWarehouse.Visible = isVisible
    Me.lblTargetWarehouseError.Visible = isVisible
    Me.fraMode.Visible = isVisible
    Me.optArchiveOnly.Visible = isVisible
    Me.optArchiveMigrate.Visible = isVisible
    Me.optArchiveRetire.Visible = isVisible
    Me.optArchiveRetireDelete.Visible = isVisible
    Me.lblArchiveDestPath.Visible = isVisible
    Me.txtArchiveDestPath.Visible = isVisible
    Me.lblArchiveDestPathError.Visible = isVisible
    Me.chkPublishTombstone.Visible = isVisible And (ResolveSelectedMode() = modWarehouseRetire.MODE_ARCHIVE_RETIRE Or ResolveSelectedMode() = modWarehouseRetire.MODE_ARCHIVE_RETIRE_DELETE)
    Me.lblReAuthError.Visible = isVisible
End Sub

Private Sub ClearAllInlineErrors()
    ClearInlineError Me.lblSourceWarehouseError
    ClearInlineError Me.lblTargetWarehouseError
    ClearInlineError Me.lblArchiveDestPathError
    ClearInlineError Me.lblReAuthError
    ClearInlineError Me.lblConfirmError
End Sub

Private Sub ClearInlineError(ByVal lbl As MSForms.Label)
    If lbl Is Nothing Then Exit Sub
    lbl.Caption = ""
    lbl.ForeColor = COLOR_ERROR
End Sub

Private Sub SetInlineError(ByVal lbl As MSForms.Label, ByVal messageText As String)
    If lbl Is Nothing Then Exit Sub
    lbl.Caption = Trim$(messageText)
    lbl.ForeColor = COLOR_ERROR
End Sub

Private Sub ShowFormMessage(ByVal messageText As String, ByVal foreColor As Long)
    Me.lblSelectionIntro.Caption = Trim$(messageText)
    Me.lblSelectionIntro.ForeColor = foreColor
End Sub

Private Function ResolveCurrentUserForm() As String
    ResolveCurrentUserForm = Trim$(Environ$("USERNAME"))
    If ResolveCurrentUserForm = "" Then ResolveCurrentUserForm = Trim$(Application.UserName)
End Function

Private Function NormalizePathForm(ByVal pathText As String) As String
    pathText = Trim$(Replace$(pathText, "/", "\"))
    Do While Len(pathText) > 3 And Right$(pathText, 1) = "\"
        pathText = Left$(pathText, Len(pathText) - 1)
    Loop
    NormalizePathForm = pathText
End Function

Private Function GetParentFolderForm(ByVal pathText As String) As String
    Dim sepPos As Long

    pathText = NormalizePathForm(pathText)
    sepPos = InStrRev(pathText, "\")
    If sepPos > 1 Then GetParentFolderForm = Left$(pathText, sepPos - 1)
End Function

Private Sub RestoreRootOverrideForm(ByVal priorRoot As String)
    If Trim$(priorRoot) = "" Then
        modRuntimeWorkbooks.ClearCoreDataRootOverride
    Else
        modRuntimeWorkbooks.SetCoreDataRootOverride priorRoot
    End If
End Sub
