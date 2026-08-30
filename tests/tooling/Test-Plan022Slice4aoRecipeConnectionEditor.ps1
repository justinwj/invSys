[CmdletBinding()]
param([string]$RepoRoot = ".")

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$docs = (Resolve-Path -LiteralPath (Join-Path $repo "..\invSys_docs")).Path

function Read-Text([string]$Path) {
    Get-Content -Raw -LiteralPath $Path
}

$form = Read-Text (Join-Path $repo "src\Production\Forms\frmProduction.frm")
$validator = Read-Text (Join-Path $repo "tools\validate_plan022_packaged_launchers.ps1")
$spec = Read-Text (Join-Path $docs "0 plan docs\xlam_invSys\invSys-Design-v4.11.md")
$plan = Read-Text (Join-Path $docs "expert guidance docs\022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md")
$controls = Read-Text (Join-Path $docs "0 plan docs\xlam_invSys\invSys-Controls-v1.md")

$checks = @(
    [pscustomobject]@{ Name = "Docs.NameAndCatalogContract"; Pass =
        $spec -match 'connection output by its Process output name' -and
        $spec -match 'generated `OutputId`' -and
        $spec -match 'remains hidden control/domain identity' -and
        $spec -match 'Connection UOM is selected from the current warehouse' -and
        $plan -match 'Slice 4ao -- Recipe connection output names and catalog UOM selection' -and
        $controls -match 'Slice 4ao Recipe connection labels and UOM control' },
    [pscustomobject]@{ Name = "Form.ConnectionUomIsCatalogCombo"; Pass =
        $form -match 'Private WithEvents mCmbConnectionUom As MSForms\.ComboBox' -and
        $form -match 'AddCombo\(pg, "cmbConnectionUom"' -and
        $form -notmatch 'mTxtConnectionUom As MSForms\.TextBox' -and
        $form -notmatch 'AddText\(pg, "txtConnectionUom"' },
    [pscustomobject]@{ Name = "Form.ConnectionUomLoadsSettingsCatalog"; Pass =
        $form -match 'RefreshConnectionUomCatalog' -and
        $form -match 'modUomSettings\.GetConfiguredUoms\(\)' -and
        $form -match 'SelectComboText mCmbConnectionUom' },
    [pscustomobject]@{ Name = "Form.OutputShowsNameAndBindsId"; Pass =
        $form -match 'With mCmbConnectionOutput' -and
        $form -match '\.ColumnCount = 2' -and
        $form -match '\.BoundColumn = 1' -and
        $form -match '\.TextColumn = 2' -and
        $form -match 'ReusableRecordText\(record, "OutputName"\)' },
    [pscustomobject]@{ Name = "Form.ConnectUpdatePreserveHiddenId"; Pass =
        $form -match 'mBtnRecipeConnect_Click' -and
        $form -match 'mBtnRecipeUpdateConnection_Click' -and
        $form -match 'ConnectionOutputId\(\)' -and
        $form -match '\.List\(idx, 1\) = ConnectionOutputId\(\)' -and
        $form -match '\.List\(idx, 6\) = ComboText\(mCmbConnectionUom\)' },
    [pscustomobject]@{ Name = "PublicAction.ExercisesVisibleNameAndCatalogUom"; Pass =
        $form -match '\|RecipeOutputNameVisible=' -and
        $form -match '\|RecipeOutputIdPreserved=' -and
        $form -match '\|RecipeUomCatalog=' -and
        $form -match '\|RecipeConnectionUpdated=' },
    [pscustomobject]@{ Name = "Packaged.RequiresRecipeConnectionContract"; Pass =
        $validator -match 'RecipeOutputNameVisible=True' -and
        $validator -match 'RecipeOutputIdPreserved=True' -and
        $validator -match 'RecipeUomCatalog=True' -and
        $validator -match 'RecipeConnectionUpdated=True' }
)

$failed = @($checks | Where-Object { -not $_.Pass })
foreach ($check in $checks) {
    "{0} {1}" -f $(if ($check.Pass) { "PASS" } else { "FAIL" }), $check.Name
}

"PLAN022_SLICE4AO_SOURCE passed=$($checks.Count - $failed.Count) red=$($failed.Count) total=$($checks.Count)"
if ($failed.Count -gt 0) {
    throw "Plan 022 Slice 4ao source contract RED: $($failed.Name -join ', ')"
}
