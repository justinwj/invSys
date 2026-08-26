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

$checks["Docs.OutputSkuContract"] =
    $spec.Contains("system-managed **Output SKU**") -and
    $plan.Contains("Slice 4ac -- Process OUTPUT managed-item picker") -and
    $controls.Contains("Slice 4ac Process OUTPUT managed-item picker:")
$checks["Worksheet.OutputManagedItemTarget"] =
    $worksheet.Contains("ProcessManagedItemPairNumber") -and
    $worksheet.Contains('Case "OUTPUT"') -and
    $worksheet.Contains('lo.ListColumns("Output SKU")')
$checks["Worksheet.OutputSkuRoundTrip"] =
    $worksheet.Contains('rowRecord("OutputSku")') -and
    $worksheet.Contains('record("ITEM_CODE") = outputSku')
$checks["Picker.OutputCommit"] =
    $picker.Contains('cProcessOutputSku = ColumnIndex(lo, "Output SKU")') -and
    $picker.Contains('cProcessName = ColumnIndex(lo, "Name")')
$checks["Packaged.PublicOutputPicker"] =
    $form.Contains("TestProcessWorksheetOutputPickerContract") -and
    $form.Contains("OutputPickerOpened=True") -and
    $form.Contains("OutputSkuRoundTrip=True")
$checks["Validator.RequiresOutputPicker"] =
    $validator.Contains("OutputPickerOpened=True") -and
    $validator.Contains("OutputSkuRoundTrip=True")

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

Write-Host ("PLAN022_SLICE4AC_SOURCE passed={0} red={1} total={2}" -f $passed, $red, $checks.Count)
if ($red -gt 0) { exit 1 }
