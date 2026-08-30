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
$checks["Docs.CompactOutputEditorContract"] =
    $spec -match 'Hidden output SKU state must not reserve a visible gap' -and
    $spec -match 'Output\s+UOM is selected from a\s+dropdown' -and
    $plan.Contains("Slice 4as -- Compact Process Output editor and catalog UOM") -and
    $controls.Contains("Slice 4as compact Process Output editor")
$checks["Form.HiddenSkuDoesNotReserveGap"] =
    $form.Contains('Set mTxtProcessOutputItemCode = AddText(pg, "txtProcessOutputItemCode", 463, 342, 1, 1)') -and
    $form.Contains('Set mTxtProcessOutputDesignId = AddText(pg, "txtProcessOutputDesignId", 463, 342, 48, 22)')
$checks["Form.AllVisibleOutputFieldsOneRow"] =
    $form.Contains('Set mTxtProcessOutputDesignVersion = AddText(pg, "txtProcessOutputDesignVersion", 515, 342, 31, 22)') -and
    $form.Contains('Set mTxtProcessOutputQty = AddText(pg, "txtProcessOutputQty", 550, 342, 36, 22)') -and
    $form.Contains('Set mTxtProcessOutputPercent = AddText(pg, "txtProcessOutputPercent", 590, 342, 36, 22)') -and
    $form.Contains('Set mTxtProcessOutputYieldBasis = AddText(pg, "txtProcessOutputYieldBasis", 630, 342, 40, 22)') -and
    $form.Contains('Set mCmbProcessOutputUom = AddCombo(pg, "cmbProcessOutputUom", 674, 342, 56, 22)')
$checks["Form.OutputUomUsesCatalog"] =
    $form.Contains("Private WithEvents mCmbProcessOutputUom As MSForms.ComboBox") -and
    $form.Contains("Private Sub RefreshProcessOutputUomCatalog") -and
    $form.Contains("modUomSettings.GetConfiguredUoms()") -and
    $form.Contains("SelectComboText mCmbProcessOutputUom, selectedUom") -and
    $form.Contains("OutputUomIsCatalogValue")
$checks["PublicAction.ExercisesOutputEditor"] =
    $form.Contains("|ProcessOutputEditorCompact=") -and
    $form.Contains("|ProcessOutputUomCatalog=")
$checks["Packaged.RequiresOutputEditorEvidence"] =
    $validator.Contains("ProcessOutputEditorCompact=True") -and
    $validator.Contains("ProcessOutputUomCatalog=True")

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

Write-Host ("PLAN022_SLICE4AS_SOURCE passed={0} red={1} total={2}" -f $passed, $red, $checks.Count)
if ($red -gt 0) { exit 1 }
