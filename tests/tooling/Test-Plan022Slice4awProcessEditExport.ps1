[CmdletBinding()]
param(
    [string]$RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
)

$ErrorActionPreference = "Stop"
$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$docs = (Resolve-Path -LiteralPath (Join-Path $repo "..\invSys_docs")).Path

$form = Get-Content -LiteralPath (Join-Path $repo "src\Production\Forms\frmProduction.frm") -Raw
$bridge = Get-Content -LiteralPath (Join-Path $repo "src\Production\Modules\mProduction.bas") -Raw
$validator = Get-Content -LiteralPath (Join-Path $repo "tools\validate_plan022_packaged_launchers.ps1") -Raw
$spec = Get-Content -LiteralPath (Join-Path $docs "0 plan docs\xlam_invSys\invSys-Design-v4.11.md") -Raw
$plan = Get-Content -LiteralPath (Join-Path $docs "expert guidance docs\022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md") -Raw
$controls = Get-Content -LiteralPath (Join-Path $docs "0 plan docs\xlam_invSys\invSys-Controls-v1.md") -Raw

$checks = [ordered]@{}
$checks["Docs.ReleasedProcessEditExportContract"] =
    $spec.Contains("A saved/released Process sent for") -and
    $spec.Contains("editing becomes a new generated DRAFT version") -and
    $plan.Contains("Slice 4aw -- Released Process editing and worksheet export") -and
    $controls.Contains("Edit as New Version") -and
    $controls.Contains("Send Process to Sheet")
$checks["Form.OperatorWording"] =
    $form.Contains('"View Process"') -and
    $form.Contains('"Edit as New Version"') -and
    $form.Contains('"Send Process to Sheet"')
$checks["Form.SuccessorVersionRebasesOutputs"] =
    $form.Contains("Private Sub SetEditableProcessDraftVersion") -and
    $form.Contains("mLstProcessOutputs.List(rowIndex, 4) = processVersion") -and
    $form.Contains("SetEditableProcessDraftVersion _")
$checks["Form.RowSelectionSurvivesEditorFocus"] =
    $form.Contains("Private mSelectedProcessRequirementIndex As Long") -and
    $form.Contains("Private mSelectedProcessOutputIndex As Long") -and
    $form.Contains("If updateExisting And idx < 0 Then") -and
    $form.Contains("idx = mSelectedProcessRequirementIndex") -and
    $form.Contains("idx = mSelectedProcessOutputIndex")
$checks["Form.RealHandlersExportAndRetrieve"] =
    $form.Contains("mBtnProcessReuse_Click") -and
    $form.Contains("mBtnProcessWorksheetCreate_Click") -and
    $form.Contains("mBtnProcessWorksheetRetrieve_Click") -and
    $form.Contains("|ReleasedProcessEditable=True") -and
    $form.Contains("|ExistingProcessExported=True") -and
    $form.Contains("|ExportRoundTrip=True") -and
    $form.Contains("|OutputDesignVersionRebased=True") -and
    $form.Contains("|OutputYieldRebased=True")
$checks["Bridge.PublicPackagedEntry"] =
    $bridge.Contains("RunProcessEditExportContractTest") -and
    $bridge.Contains("TestProcessEditExportContract")
$checks["Packaged.RequiresEditExportEvidence"] =
    $validator.Contains("ReleasedProcessEditable=True") -and
    $validator.Contains("ExistingProcessExported=True") -and
    $validator.Contains("ExportRoundTrip=True") -and
    $validator.Contains("OutputDesignVersionRebased=True") -and
    $validator.Contains("OutputYieldRebased=True")

$passed = 0
$red = 0
foreach ($entry in $checks.GetEnumerator()) {
    if ($entry.Value) {
        Write-Host ("PASS " + $entry.Key)
        $passed++
    }
    else {
        Write-Host ("RED  " + $entry.Key)
        $red++
    }
}

Write-Host ("PLAN022_SLICE4AW_SOURCE passed={0} red={1} total={2}" -f $passed, $red, $checks.Count)
if ($red -gt 0) { exit 1 }
