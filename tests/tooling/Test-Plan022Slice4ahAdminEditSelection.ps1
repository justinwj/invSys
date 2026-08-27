[CmdletBinding()]
param(
    [string]$RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
)

$ErrorActionPreference = "Stop"
$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$docs = (Resolve-Path -LiteralPath (Join-Path $repo "..\invSys_docs")).Path

$plan = Get-Content -LiteralPath (Join-Path $docs "expert guidance docs\022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md") -Raw
$controls = Get-Content -LiteralPath (Join-Path $docs "0 plan docs\xlam_invSys\invSys-Controls-v1.md") -Raw
$form = Get-Content -LiteralPath (Join-Path $repo "src\Admin\Forms\frmAddInventoryItem.frm") -Raw
$admin = Get-Content -LiteralPath (Join-Path $repo "src\Admin\Modules\modAdmin.bas") -Raw
$validator = Get-Content -LiteralPath (Join-Path $repo "tools\validate_phase6_packaged_xlams.ps1") -Raw

$checks = [ordered]@{}
$checks["Docs.Slice4ahContract"] =
    $plan.Contains("Slice 4ah -- Admin inventory edit selection binding") -and
    $controls.Contains("Slice 4ah Admin inventory edit selection binding:")
$checks["Form.ComboSelectionCommitsIdentity"] =
    $form.Contains("Private Function CommitSelectedEditItemFromCombo") -and
    $form -match 'Private Sub mCmbEditItem_Change\(\)(?s).*?CommitSelectedEditItemFromCombo'
$checks["Form.OperatorHandlerTest"] =
    $form.Contains("Public Function TestEditItemComboSelectionContract") -and
    $form -match 'TestEditItemComboSelectionContract\(\)(?s).*?mCmbEditItem_Change' -and
    $form.Contains("ComboSelected=") -and
    $form.Contains("UtilityReady=")
$checks["Admin.PackagedEntryPoint"] =
    $admin.Contains("Public Function InventoryEditSelectionContractForAutomation") -and
    $admin.Contains("frmAddInventoryItem.TestEditItemComboSelectionContract")
$checks["Packaged.ValidatorRequiresSelection"] =
    $validator.Contains("Admin.EditItemComboSelection") -and
    $validator.Contains("ComboSelected=True") -and
    $validator.Contains("FieldsLoaded=True") -and
    $validator.Contains("UtilityReady=True")

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

Write-Host ("PLAN022_SLICE4AH_SOURCE passed={0} red={1} total={2}" -f $passed, $red, $checks.Count)
if ($red -gt 0) { exit 1 }
