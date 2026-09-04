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
$queries = Get-Content -LiteralPath (Join-Path $repo "src\InventoryDomain\Modules\modInventoryQueries.bas") -Raw
$validator = Get-Content -LiteralPath (Join-Path $repo "tools\validate_plan022_packaged_launchers.ps1") -Raw

$checks = [ordered]@{}
$checks["Docs.Slice4ajContract"] =
    $spec.Contains("Production Output retains a separate row") -and
    $plan.Contains("Slice 4aj -- Production batch history and Utility projection") -and
    $controls.Contains("Slice 4aj Production batch history and Utility projection")
$checks["Form.OutputHistoryHeaders"] =
    $form.Contains('Array("Process", "Output", "UOM", "Last Actual", "Batch", "Used Goods", "Process Total", "Recall", "System_Key")') -and
    -not $form.Contains('Array("Process", "Output", "UOM", "Last Actual", "Batch", "Planned", "Recall", "System_Key")')
$checks["Run.RetainedBatchHistory"] =
    $run.Contains("Private mOutputHistory As Collection") -and
    $run.Contains("CaptureCompletedOutputHistory") -and
    $run.Contains("ReusableRunOutputDefinitionIndex")
$checks["Run.UsedGoodsAndProcessTotal"] =
    $run.Contains("ProcessUsedGoodsQty") -and
    $run.Contains("ProcessOutputTotal") -and
    $run.Contains('record("ProcessTotal")')
$checks["Form.ActiveRowHandlers"] =
    $form -match 'LoadSelectedReusableActualOutput(?s).*?ReusableRunOutputDefinitionIndex' -and
    $form -match 'StageSelectedReusableActualOutput(?s).*?ReusableRunOutputDefinitionIndex'
$checks["Inventory.UtilityMetadataProjection"] =
    $queries.Contains('result(outRow, 11) = trackQty') -and
    $queries.Contains('result(outRow, 12) = itemKind') -and
    $queries.Contains('result(outRow, 13) = categoryValue') -and
    $queries.Contains('columnIndex = InventoryQueryColumn(lo, columnName)') -and
    $run.Contains("ExactEntityInventoryDisplay") -and
    $run.Contains('ExactEntityInventoryDisplay(entities, entityRow)')
$checks["Packaged.RealHandlerEvidence"] =
    $form.Contains('"|BatchHistoryRows="') -and
    $form.Contains('"|ProcessTotal="') -and
    $form.Contains('"|UtilityDisplay="') -and
    $validator.Contains("BatchHistoryRows=True") -and
    $validator.Contains("UtilityDisplay=True")

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

Write-Host ("PLAN022_SLICE4AJ_SOURCE passed={0} red={1} total={2}" -f $passed, $red, $checks.Count)
if ($red -gt 0) { exit 1 }
