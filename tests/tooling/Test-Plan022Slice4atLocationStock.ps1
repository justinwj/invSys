[CmdletBinding()]
param(
    [string]$RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
)

$ErrorActionPreference = "Stop"
$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$docs = (Resolve-Path -LiteralPath (Join-Path $repo "..\invSys_docs")).Path

$form = Get-Content -LiteralPath (Join-Path $repo "src\Production\Forms\frmProduction.frm") -Raw
$run = Get-Content -LiteralPath (Join-Path $repo "src\Production\Modules\modProductionReusableRun.bas") -Raw
$receiving = Get-Content -LiteralPath (Join-Path $repo "src\Receiving\Forms\frmReceiving.frm") -Raw
$validator = Get-Content -LiteralPath (Join-Path $repo "tools\validate_plan022_packaged_launchers.ps1") -Raw
$spec = Get-Content -LiteralPath (Join-Path $docs "0 plan docs\xlam_invSys\invSys-Design-v4.11.md") -Raw
$plan = Get-Content -LiteralPath (Join-Path $docs "expert guidance docs\022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md") -Raw
$controls = Get-Content -LiteralPath (Join-Path $docs "0 plan docs\xlam_invSys\invSys-Controls-v1.md") -Raw

$checks = [ordered]@{}
$checks["Docs.LocationStockContract"] =
    $spec -match 'stock bucket rather than\s+an entity' -and
    $spec -match 'expand the chosen bucket deterministically into its contributing exact keys' -and
    $spec.Contains("Capacity (coming later)") -and
    $plan.Contains("Slice 4at -- Location-stock Production allocation and Receiving capacity stub") -and
    $controls.Contains("Slice 4at location-stock allocation and Capacity stub")
$checks["Production.AssignmentSystemKeyReadable"] =
    $form.Contains("Private Function AssignmentSystemKeyReadable") -and
    $form.Contains('12, 312, "190 pt;105 pt;35 pt;45 pt;70 pt;65 pt;0 pt"')
$checks["Production.LocationStockProjection"] =
    $run.Contains("Private Function StockBucketKey") -and
    $run.Contains("Private Function StockBucketAvailableDisplay") -and
    $form.Contains('Array("", "", "Process / Ingredient", "", "Inventory Stock", "% Req", "Qty", "UOM", "Available", "Location")') -and
    $form.Contains('"|LocationStockBuckets=" & CStr(locationStockBuckets)')
$checks["Production.ExactExpansion"] =
    $run.Contains("Public Function ApplyReusableRunStockAllocation") -and
    $run.Contains("Exact allocation expansion saved:") -and
    $form.Contains("modProductionReusableRun.ApplyReusableRunStockAllocation") -and
    $form.Contains('"|LocationStockExactExpansion=" & CStr(locationStockExactExpansion)')
$checks["Receiving.CapacityStub"] =
    $receiving.Contains('Array("Code", "Item", "UOM", "Available", "Location", "Capacity (coming later)", "Lot", "Condition", "Description", "Vendor")') -and
    $receiving.Contains("Private mReceiveItemSystemKeys As Collection") -and
    $receiving.Contains("Private Sub FillReceiveItemResults") -and
    $receiving.Contains('|CapacityStub=" & CStr(capacityStub)')
$checks["PublicActions.ExerciseStockAndCapacity"] =
    $form.Contains("|AssignmentSystemKeyReadable=") -and
    $form.Contains("|LocationStockBuckets=") -and
    $form.Contains("|LocationStockExactExpansion=") -and
    $receiving.Contains("|CapacityStub=")
$checks["Packaged.RequiresSlice4atEvidence"] =
    $validator.Contains("AssignmentSystemKeyReadable=True") -and
    $validator.Contains("LocationStockBuckets=True") -and
    $validator.Contains("LocationStockExactExpansion=True") -and
    $validator.Contains("CapacityStub=True")

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

Write-Host ("PLAN022_SLICE4AT_SOURCE passed={0} red={1} total={2}" -f $passed, $red, $checks.Count)
if ($red -gt 0) { exit 1 }
