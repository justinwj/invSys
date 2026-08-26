[CmdletBinding()]
param(
    [string]$RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
)

$ErrorActionPreference = "Stop"
$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$docs = (Resolve-Path -LiteralPath (Join-Path $repo "..\invSys_docs")).Path

$checks = [ordered]@{}
$spec = Get-Content -LiteralPath (Join-Path $docs "0 plan docs\xlam_invSys\invSys-Design-v4.11.md") -Raw
$plan = Get-Content -LiteralPath (Join-Path $docs "expert guidance docs\022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md") -Raw
$controls = Get-Content -LiteralPath (Join-Path $docs "0 plan docs\xlam_invSys\invSys-Controls-v1.md") -Raw
$worksheet = Get-Content -LiteralPath (Join-Path $repo "src\Production\Modules\modProductionProcessWorksheet.bas") -Raw
$picker = Get-Content -LiteralPath (Join-Path $repo "src\Core\ClassModules\cDynItemSearch.cls") -Raw
$form = Get-Content -LiteralPath (Join-Path $repo "src\Production\Forms\frmProduction.frm") -Raw
$validator = Get-Content -LiteralPath (Join-Path $repo "tools\validate_plan022_packaged_launchers.ps1") -Raw

$checks["Docs.RecordTypeReachability"] =
    $spec.Contains("same Core item-search interaction as INPUT") -and
    $plan.Contains("Slice 4ad -- Process picker INPUT/OUTPUT record-type reachability") -and
    $controls.Contains("Slice 4ad Process picker INPUT/OUTPUT reachability:")
$checks["Worksheet.OutputManagedItemTarget"] =
    $worksheet.Contains("IsProcessWorksheetOutputManagedItemTarget") -and
    $worksheet.Contains('Set managedItemColumn = lo.ListColumns("Acceptable Managed Item 1")')
$checks["Picker.OutputManagedItemCommit"] =
    $picker.Contains('If processRecordType = "OUTPUT" Then') -and
    $picker.Contains('cProcessItem = ColumnIndex(lo, "Acceptable Managed Item 1")') -and
    $picker.Contains('cProcessOutputSku = ColumnIndex(lo, "Output SKU")')
$checks["Worksheet.OutputSkuImport"] =
    $worksheet.Contains('If outputSku = "" Then outputSku = acceptedSku') -and
    $worksheet.Contains('record("ITEM_CODE") = outputSku')
$checks["Packaged.ActualOutputManagedItemCell"] =
    $form.Contains('outputItemColumn = ProcessWorksheetColumnForTest(lo, "Acceptable Managed Item 1")') -and
    $form.Contains("OutputNameRetained=True")
$checks["Validator.RequiresRetainedName"] =
    $validator.Contains("OutputNameRetained=True")

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

Write-Host ("PLAN022_SLICE4AD_SOURCE passed={0} red={1} total={2}" -f $passed, $red, $checks.Count)
if ($red -gt 0) { exit 1 }
