[CmdletBinding()]
param(
    [string]$RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
)

$ErrorActionPreference = "Stop"
$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$docs = (Resolve-Path -LiteralPath (Join-Path $repo "..\invSys_docs")).Path

$form = Get-Content -LiteralPath (Join-Path $repo "src\Production\Forms\frmProduction.frm") -Raw
$validator = Get-Content -LiteralPath (Join-Path $repo "tools\validate_plan022_packaged_launchers.ps1") -Raw
$spec = Get-Content -LiteralPath (Join-Path $docs "0 plan docs\xlam_invSys\invSys-Design-v4.11.md") -Raw
$plan = Get-Content -LiteralPath (Join-Path $docs "expert guidance docs\022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md") -Raw
$controls = Get-Content -LiteralPath (Join-Path $docs "0 plan docs\xlam_invSys\invSys-Controls-v1.md") -Raw

$checks = [ordered]@{}
$checks["Docs.OutputYieldAndRunPlanContract"] =
    $spec -match 'Output Qty, Yield %, and UOM always come from the producing Process' -and
    $spec -match '(?s)Scale from target output Qty \(coming\s+later\)' -and
    $plan.Contains("Slice 4ar -- Output-yield clarity, list headers, and multi-Process run plan") -and
    $controls.Contains("Slice 4ar output-yield clarity and multi-Process run plan")
$checks["Form.OutputYieldProjection"] =
    $form.Contains('Array("Stage", "Produced by", "Output", "Feeds Process", "Output Qty", "Yield %", "UOM", "", "", "")') -and
    $form.Contains('"Produced by / Output / Feeds Process / Required Qty / Required % / UOM"') -and
    $form -match 'AddRecipeOutputFlowRow[\s\S]{0,500}outputQty[\s\S]{0,120}outputPercent'
$checks["Form.OutputYieldDefaultsAndUpdate"] =
    $form.Contains("NormalizeOutputYieldEditorDefaults") -and
    $form.Contains("NormalizedOutputYieldPercent") -and
    $form.Contains("NormalizedOutputYieldBasis") -and
    $form.Contains("|OutputYieldDefaults=")
$checks["Form.ProcessAndAssignmentHeaders"] =
    $form.Contains('AddColumnHeaders pg, "SavedProcesses"') -and
    $form.Contains('AddColumnHeaders pg, "ProcessRequirements"') -and
    $form.Contains('AddColumnHeaders pg, "ProcessOutputs"') -and
    $form.Contains('AddColumnHeaders pg, "ProcessInstructions"') -and
    $form.Contains('AddColumnHeaders pg, "AssignmentProcesses"') -and
    $form.Contains('AddColumnHeaders pg, "AssignmentRequirements"') -and
    $form.Contains('AddColumnHeaders pg, "AssignmentInventory"') -and
    $form.Contains('AddColumnHeaders pg, "AssignmentAllowed"')
$checks["Form.AcceptableItemsNamed"] =
    $form.Contains("ManagedItemDisplayForAssignmentCode") -and
    $form.Contains('Array("", "Managed Item", "UOM", "Item Code"') -and
    $form.Contains("|AcceptableItemsNamed=")
$checks["Form.MultiProcessRunProjection"] =
    $form.Contains('"Multi-Process Run Plan"') -and
    $form.Contains('"Process filter"') -and
    $form.Contains('"Process / Ingredient"') -and
    $form.Contains("ApplyReusablePaletteProcessProjection") -and
    $form.Contains("|MultiProcessRunPlan=")
$checks["Form.TargetOutputScaleStub"] =
    $form.Contains('"Scale from target output Qty (coming later)"') -and
    $form -match '(?s)mChkRunTargetOutputScale.*?\.Enabled = False' -and
    $form.Contains("mCmbRunTargetOutput.Enabled = False") -and
    $form.Contains("mTxtRunTargetOutputQty.Enabled = False") -and
    $form.Contains("|TargetOutputScaleStub=")
$checks["Packaged.RequiresSlice4arEvidence"] =
    $validator.Contains("OutputYieldDefaults=True") -and
    $validator.Contains("OutputFlowUsesProcessYield=True") -and
    $validator.Contains("ProcessAssignmentHeaders=True") -and
    $validator.Contains("AcceptableItemsNamed=True") -and
    $validator.Contains("MultiProcessRunPlan=True") -and
    $validator.Contains("TargetOutputScaleStub=True")

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

Write-Host ("PLAN022_SLICE4AR_SOURCE passed={0} red={1} total={2}" -f $passed, $red, $checks.Count)
if ($red -gt 0) { exit 1 }
