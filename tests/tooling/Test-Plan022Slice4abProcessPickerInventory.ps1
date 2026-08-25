[CmdletBinding()]
param(
    [string]$RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
)

$ErrorActionPreference = "Stop"
$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$docs = (Resolve-Path -LiteralPath (Join-Path $repo "..\invSys_docs")).Path

$checks = [ordered]@{}
$plan = Get-Content -LiteralPath (Join-Path $docs "expert guidance docs\022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md") -Raw
$controls = Get-Content -LiteralPath (Join-Path $docs "0 plan docs\xlam_invSys\invSys-Controls-v1.md") -Raw
$picker = Get-Content -LiteralPath (Join-Path $repo "src\Core\ClassModules\cDynItemSearch.cls") -Raw
$form = Get-Content -LiteralPath (Join-Path $repo "src\Production\Forms\frmProduction.frm") -Raw
$validator = Get-Content -LiteralPath (Join-Path $repo "tools\validate_plan022_packaged_launchers.ps1") -Raw

$checks["Docs.Slice4abContract"] =
    $plan.Contains("Slice 4ab -- Process picker managed-inventory projection blocker") -and
    $controls.Contains("Slice 4ab Process picker inventory projection:")
$checks["Picker.SystemKeyIdentity"] =
    $picker.Contains("LoadProcessManagedInventoryItems") -and
    $picker.Contains('modInventoryDomainBridge.ListAvailableInventoryEntitiesBridge("")') -and
    $picker.Contains('systemKey = Trim$(NzStr(sourceRows(r, 1)))')
$checks["Packaged.PublicPickerRows"] =
    $form.Contains("PickerInventoryRows=True") -and
    $form.Contains("ProductionProcessItemSearchResultCountForTest")
$checks["Validator.RequiresPickerRows"] =
    $validator.Contains("PickerInventoryRows=True")

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

Write-Host ("PLAN022_SLICE4AB_SOURCE passed={0} red={1} total={2}" -f $passed, $red, $checks.Count)
if ($red -gt 0) { exit 1 }
