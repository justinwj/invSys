[CmdletBinding()]
param(
    [string]$RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
)

$ErrorActionPreference = "Stop"
$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$docs = (Resolve-Path -LiteralPath (Join-Path $repo "..\invSys_docs")).Path

$form = Get-Content -LiteralPath (Join-Path $repo "src\Production\Forms\frmProduction.frm") -Raw
$run = Get-Content -LiteralPath (Join-Path $repo "src\Production\Modules\modProductionReusableRun.bas") -Raw
$validator = Get-Content -LiteralPath (Join-Path $repo "tools\validate_plan022_packaged_launchers.ps1") -Raw
$spec = Get-Content -LiteralPath (Join-Path $docs "0 plan docs\xlam_invSys\invSys-Design-v4.11.md") -Raw
$plan = Get-Content -LiteralPath (Join-Path $docs "expert guidance docs\022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md") -Raw
$controls = Get-Content -LiteralPath (Join-Path $docs "0 plan docs\xlam_invSys\invSys-Controls-v1.md") -Raw

$checks = [ordered]@{}
$checks["Docs.ProcessScopedRunContract"] =
    $spec.Contains("Production Run executes one selected Process at a time") -and
    $plan.Contains("Slice 4av -- Process-scoped Production execution, plan sufficiency, and run instructions") -and
    $controls.Contains('header-backed `lstRunInstructions`')
$checks["Run.ProcessScopedState"] =
    $run.Contains("Private mCompletedNodes As Object") -and
    $run.Contains("Private mCheckedInNodeId As String") -and
    $run.Contains("Public Function CheckInReusableProcess") -and
    $run.Contains("Public Function CompleteReusableProcess")
$checks["Run.SelectedRequirementsOnly"] =
    $run.Contains("ValidateProcessRequirementsReady") -and
    $run.Contains("ValidateProcessAllocationsLive") -and
    $run.Contains("Upstream output is not ready")
$checks["Run.WholeRecipeStatus"] =
    $run.Contains("ReusableRunLineStatus") -and
    $run.Contains('"! INSUFFICIENT"') -and
    $run.Contains('"WAITING UPSTREAM"') -and
    $run.Contains('"NEEDS ALLOCATION"') -and
    $run.Contains('"READY"') -and
    $run.Contains('"COMPLETE"')
$checks["Run.SelectedInstructions"] =
    $run.Contains("Private mInstructions As Collection") -and
    $run.Contains("Public Function ReusableRunInstructionRows") -and
    $form.Contains("Private WithEvents mLstRunInstructions As MSForms.ListBox") -and
    $form.Contains("RefreshReusableRunInstructions")
$checks["Form.EightRowPaletteAndInstructions"] =
    $form.Contains('AddList(pg, "lstRunPalette", 12, 250, 1018, 96, 10') -and
    $form.Contains('AddColumnHeaders pg, "RunInstructions", Array("Step", "Instruction")') -and
    $form.Contains('AddList(pg, "lstRunInstructions"')
$checks["Form.RealHandlersUseSelectedProcess"] =
    $form.Contains("modProductionReusableRun.CheckInReusableProcess(ActiveRunProcess()") -and
    $form.Contains("modProductionReusableRun.CompleteReusableProcess(ActiveRunProcess()") -and
    $form.Contains("|SelectedProcessOnly=") -and
    $form.Contains("|RunInstructionsVisible=") -and
    $form.Contains("|WholeRecipeStatus=")
$checks["Packaged.RequiresSlice4avEvidence"] =
    $validator.Contains("SelectedProcessOnly=True") -and
    $validator.Contains("RunInstructionsVisible=True") -and
    $validator.Contains("WholeRecipeStatus=True") -and
    $validator.Contains("EightPaletteRows=True")

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

Write-Host ("PLAN022_SLICE4AV_SOURCE passed={0} red={1} total={2}" -f $passed, $red, $checks.Count)
if ($red -gt 0) { exit 1 }
