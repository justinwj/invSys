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
$json = Get-Content -LiteralPath (Join-Path $repo "src\Production\Modules\modProductionJson.bas") -Raw
$form = Get-Content -LiteralPath (Join-Path $repo "src\Production\Forms\frmProduction.frm") -Raw
$validator = Get-Content -LiteralPath (Join-Path $repo "tools\validate_plan022_packaged_launchers.ps1") -Raw

$checks["Docs.CellAndIdentityContract"] =
    $spec -match 'only from an\s+\*\*Acceptable Managed Item n\*\* cell' -and
    $spec.Contains("one table-wide row-ID namespace") -and
    $plan.Contains("Slice 4ae -- Process picker cell boundary and table-wide row identity") -and
    $controls.Contains("Slice 4ae Process picker cell boundary and row identity:")
$checks["Worksheet.AcceptableCellsOnly"] =
    $worksheet.Contains("ProcessManagedItemPairNumber") -and
    -not $worksheet.Contains("IsProcessWorksheetOutputManagedItemTarget")
$checks["Worksheet.OneRowIdNamespace"] =
    $worksheet.Contains("Set usedRowIds = CreateObject") -and
    $worksheet.Contains('Case "INPUT", "REQUIREMENT", "OUTPUT", "INSTRUCTION"') -and
    -not $worksheet.Contains("Set outputIds = CreateObject") -and
    $form.Contains("NextProcessComponentBase36Id") -and
    $form.Contains("NormalizeProcessComponentIdentities") -and
    $form.Contains("usedComponentIds.Exists(componentId)") -and
    $json.IndexOf("Case VarType(valueIn) = vbString") -lt
        $json.IndexOf("Case IsNumeric(valueIn)")
$checks["Packaged.NameSuppression"] =
    $form.Contains("OutputNamePickerSuppressed=True") -and
    $form.Contains("ShowProductionProcessItemSearch outputNameCell")
$checks["Packaged.ChangeHandlerIdentity"] =
    $form.Contains("mProduction.HandleProductionChange") -and
    $form.Contains("UniqueRowIds=True") -and
    $form.Contains("FirstAssignedIdRetained=True")
$checks["Validator.Requires4aeEvidence"] =
    $validator.Contains("OutputNamePickerSuppressed=True") -and
    $validator.Contains("UniqueRowIds=True") -and
    $validator.Contains("FirstAssignedIdRetained=True")

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

Write-Host ("PLAN022_SLICE4AE_SOURCE passed={0} red={1} total={2}" -f $passed, $red, $checks.Count)
if ($red -gt 0) { exit 1 }
