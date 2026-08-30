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
    [pscustomobject]@{ Name = "Docs.DirectionalNamedGraphContract"; Pass =
        $spec -match 'Upstream Process / Output / Downstream' -and
        $spec -match 'same-node connection is an invalid circular self-reference' -and
        $spec -match 'final or co-product output requires no terminal connection' -and
        $plan -match 'Slice 4ap -- Named Recipe graph layout and finished-output guidance' -and
        $controls -match 'same-node self-reference' },
    [pscustomobject]@{ Name = "Form.VisibleProjectionKeepsInternalIds"; Pass =
        $form -match 'Private WithEvents mLstRecipeConnectionDisplay As MSForms\.ListBox' -and
        $form -match 'Set mLstRecipeConnectionDisplay = AddList\(pg, "lstRecipeConnectionDisplay", 12,' -and
        $form -match 'mLstRecipeConnections\.Visible = False' -and
        $form -match 'Private Sub RefreshRecipeConnectionDisplay' },
    [pscustomobject]@{ Name = "Form.NamedBoundNodeAndRequirementSelectors"; Pass =
        $form -match 'ConfigureNamedIdCombo mCmbConnectionFromNode' -and
        $form -match 'ConfigureCompatibleTargetCombo mCmbConnectionToNode' -and
        $form -match 'mCmbConnectionRequirement\.Visible = False' -and
        $form -match 'Private Sub BindSelectedCompatibleRequirement' -and
        $form -match 'mCmbConnectionToNode\.List\(rowIndex, 1\) = matchedRequirementId' -and
        $form -match 'mCmbConnectionToNode\.List\(rowIndex, 2\) = NzStr\(mLstRecipeNodes\.List\(nodeIndex, 3\)\)' -and
        $form -match 'mCmbConnectionFromNode\.List\([\s\S]{0,100}, 1\) =[\s\S]{0,100}NzStr\(mLstRecipeNodes\.List\([\s\S]{0,100}, 3\)\)' },
    [pscustomobject]@{ Name = "Form.HeadersAndFullWidthLayout"; Pass =
        $form -match 'AddColumnHeaders pg, "RecipeNodes"' -and
        $form -match 'AddColumnHeaders pg, "RecipeConnections"' -and
        $form -match '"Stage", "Produced by", "Output", "Feeds Process", "Qty", "%", "UOM"' -and
        $form -match 'AddList\(pg, "lstRecipeConnectionDisplay", 12, [0-9]+, 1018,' },
    [pscustomobject]@{ Name = "Form.FinishedOutputGuidance"; Pass =
        $form -match 'Final output: leave unconnected; Production creates it as finished inventory\.' },
    [pscustomobject]@{ Name = "Form.VisibleSelectionAndDisconnectUsePublicHandlers"; Pass =
        $form -match 'Private Sub mLstRecipeConnectionDisplay_Click\(\)' -and
        $form -match 'LoadConnectionEditorFromIndex' -and
        $form -match 'mBtnRecipeDisconnect_Click[\s\S]{0,500}RefreshRecipeConnectionDisplay' },
    [pscustomobject]@{ Name = "Form.SelfReferenceRemainsRejected"; Pass =
        $form -match 'A Process output cannot connect back to the same Recipe node\.' -and
        $form -match '\|RecipeSelfReferenceRejected=' },
    [pscustomobject]@{ Name = "PublicAction.ExercisesNamedGraphHandlers"; Pass =
        $form -match '\|RecipeNodeNamesVisible=' -and
        $form -match '\|RecipeRequirementNameVisible=' -and
        $form -match '\|RecipeConnectionNamesVisible=' -and
        $form -match '\|RecipeConnectionHeaders=' -and
        $form -match '\|RecipeConnectionsFullWidth=' -and
        $form -match '\|RecipeFinishedOutputGuidance=' -and
        $form -match '\|RecipeConnectionSelected=' -and
        $form -match '\|RecipeDisconnected=' },
    [pscustomobject]@{ Name = "Packaged.RequiresNamedGraphContract"; Pass =
        $validator -match 'RecipeNodeNamesVisible=True' -and
        $validator -match 'RecipeRequirementNameVisible=True' -and
        $validator -match 'RecipeConnectionNamesVisible=True' -and
        $validator -match 'RecipeConnectionHeaders=True' -and
        $validator -match 'RecipeConnectionsFullWidth=True' -and
        $validator -match 'RecipeFinishedOutputGuidance=True' -and
        $validator -match 'RecipeConnectionSelected=True' -and
        $validator -match 'RecipeDisconnected=True' -and
        $validator -match 'RecipeSelfReferenceRejected=True' }
)

$failed = @($checks | Where-Object { -not $_.Pass })
foreach ($check in $checks) {
    "{0} {1}" -f $(if ($check.Pass) { "PASS" } else { "FAIL" }), $check.Name
}

"PLAN022_SLICE4AP_SOURCE passed=$($checks.Count - $failed.Count) red=$($failed.Count) total=$($checks.Count)"
if ($failed.Count -gt 0) {
    throw "Plan 022 Slice 4ap source contract RED: $($failed.Name -join ', ')"
}
