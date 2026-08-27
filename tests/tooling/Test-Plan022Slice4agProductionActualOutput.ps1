[CmdletBinding()]
param(
    [string]$RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
)

$ErrorActionPreference = "Stop"
$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$docs = (Resolve-Path -LiteralPath (Join-Path $repo "..\invSys_docs")).Path

$spec = Get-Content -LiteralPath (Join-Path $docs "0 plan docs\xlam_invSys\invSys-Design-v4.11.md") -Raw
$plan = Get-Content -LiteralPath (Join-Path $docs "expert guidance docs\022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md") -Raw
$controls = Get-Content -LiteralPath (Join-Path $docs "0 plan docs\xlam_invSys\invSys-Controls-v1.md") -Raw
$form = Get-Content -LiteralPath (Join-Path $repo "src\Production\Forms\frmProduction.frm") -Raw
$run = Get-Content -LiteralPath (Join-Path $repo "src\Production\Modules\modProductionReusableRun.bas") -Raw
$validator = Get-Content -LiteralPath (Join-Path $repo "tools\validate_plan022_packaged_launchers.ps1") -Raw

$checks = [ordered]@{}
$checks["Docs.ActualOutputContract"] =
    $spec.Contains("operator-entered actual quantity") -and
    $plan.Contains("Slice 4ag -- reusable-run actual output and list readability") -and
    $controls.Contains("Slice 4ag reusable-run actual output and list readability:")
$checks["Form.ReadableRunHeaders"] =
    $form.Contains('Array("", "", "Ingredient", "System_Key", "Inventory Item"') -and
    $form.Contains('Array("System_Key", "Code", "Item", "UOM", "Used", "Total Inv")') -and
    $form.Contains('Array("Process", "Output", "UOM", "Last Actual", "Batch", "Used Goods", "Process Total", "Recall", "System_Key")') -and
    $form.Contains('AddLabel pg, "Actual Output"')
$checks["Run.PerOutputActualState"] =
    $run.Contains("Private mActualOutputQty As Object") -and
    $run.Contains("Private mLastOutputQty As Object") -and
    $run.Contains("Public Function StageReusableRunActualOutput") -and
    $run.Contains("Public Function ReusableRunActualOutput")
$checks["Form.PublicHandlersStageActual"] =
    $form -match 'Private Sub mLstManagerOutput_Click\(\)(?s).*?LoadSelectedReusableActualOutput' -and
    $form -match 'Private Sub mTxtOutputReal_Change\(\)(?s).*?StageSelectedReusableActualOutput' -and
    $form -match 'Private Sub CompleteProductionRun\(\)(?s).*?StageSelectedReusableActualOutput'
$checks["Run.ActualQuantityIsInventoryAuthority"] =
    $run.Contains("ValidateReusableActualOutputs") -and
    $run.Contains("ActualOutputQty(output)") -and
    $run.Contains('""PlannedQty""') -and
    $run.Contains('""ActualQty""')
$checks["Packaged.RequiresVisibleEvidence"] =
    $form.Contains("ActualOutputAccepted=") -and
    $form.Contains("LastActualDisplayed=") -and
    $form.Contains("ActualInventoryQty=") -and
    $form.Contains("SystemKeyHeadersReadable=") -and
    $validator.Contains("ActualOutputAccepted=True") -and
    $validator.Contains("LastActualDisplayed=True") -and
    $validator.Contains("ActualInventoryQty=True") -and
    $validator.Contains("SystemKeyHeadersReadable=True")

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

Write-Host ("PLAN022_SLICE4AG_SOURCE passed={0} red={1} total={2}" -f $passed, $red, $checks.Count)
if ($red -gt 0) { exit 1 }
