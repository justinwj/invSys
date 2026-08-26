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
$validator = Get-Content -LiteralPath (Join-Path $repo "tools\validate_plan022_packaged_launchers.ps1") -Raw

$checks = [ordered]@{}
$checks["Docs.RecipeIdentityContract"] =
    $spec -match 'Recipe ID\s+control is a locked projection' -and
    $spec -match 'Recipe version control is an editable\s+operator field' -and
    $plan.Contains("Slice 4af -- Recipe identity initialization and editable version") -and
    $controls.Contains("Slice 4af Recipe identity initialization and editable version:")
$checks["Form.IdentityControlState"] =
    $form.Contains("mTxtReusableRecipeId.Locked = True") -and
    $form.Contains("mTxtReusableRecipeVersion.Locked = False")
$checks["Form.AutomaticIdentityHelper"] =
    $form.Contains("Private Sub EnsureRecipeDraftIdentity()") -and
    $form.Contains("mTxtReusableRecipeId.Text = NextListBase36Id(mLstRecipes, 0)") -and
    $form.Contains("modProductionReusableDesigns.NextReusableDefinitionVersion")
$checks["Form.InitialAndHandlerPaths"] =
    $form -match 'RefreshAllViews\s+EnsureRecipeDraftIdentity' -and
    $form -match 'Private Sub mBtnRecipeSave_Click\(\)\s+Dim report As String\s+EnsureRecipeDraftIdentity' -and
    $form -match 'Private Sub mBtnRecipeRelease_Click\(\)\s+Dim report As String\s+EnsureRecipeDraftIdentity'
$checks["Packaged.OperatorEvidence"] =
    $form.Contains("RecipeIdentityInitialized=True") -and
    $form.Contains("RecipeVersionGenerated=True") -and
    $form.Contains("RecipeIdLocked=True") -and
    $form.Contains("RecipeVersionEditable=True") -and
    $form.Contains('"|EditedRecipeVersionRetained="')
$checks["Validator.RequiresRecipeEvidence"] =
    $validator.Contains("RecipeIdentityInitialized=True") -and
    $validator.Contains("RecipeVersionGenerated=True") -and
    $validator.Contains("RecipeIdLocked=True") -and
    $validator.Contains("RecipeVersionEditable=True") -and
    $validator.Contains("EditedRecipeVersionRetained=True")

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

Write-Host ("PLAN022_SLICE4AF_SOURCE passed={0} red={1} total={2}" -f $passed, $red, $checks.Count)
if ($red -gt 0) { exit 1 }
