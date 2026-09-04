[CmdletBinding()]
param([string]$RepoRoot = ".")

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$docs = (Resolve-Path -LiteralPath (Join-Path $repo "..\invSys_docs")).Path

function Read-Text([string]$Path) {
    Get-Content -Raw -LiteralPath $Path
}

$run = Read-Text (Join-Path $repo "src\Production\Modules\modProductionReusableRun.bas")
$form = Read-Text (Join-Path $repo "src\Production\Forms\frmProduction.frm")
$validator = Read-Text (Join-Path $repo "tools\validate_plan022_packaged_launchers.ps1")
$spec = Read-Text (Join-Path $docs "0 plan docs\xlam_invSys\invSys-Design-v4.11.md")
$plan = Read-Text (Join-Path $docs "expert guidance docs\022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md")
$controls = Read-Text (Join-Path $docs "0 plan docs\xlam_invSys\invSys-Controls-v1.md")

$checks = @(
    [pscustomobject]@{ Name = "Docs.RoutedInputAndGroupedUsedGoods"; Pass =
        $spec -match 'read-only routed-input row' -and
        $spec -match 'deterministic, normalized-UOM' -and
        $plan -match 'Approved 2026-08-30 contract refinement' -and
        $controls -match 'Lists external allocations plus read-only routed inputs' },
    [pscustomobject]@{ Name = "Run.RoutedInputRowsExposeAuditFields"; Pass =
        $run -match 'Public Function ReusableRunManagerCheckRows' -and
        $run -match 'BuildRoutedInputCheckRow' -and
        $run -match 'Source Process / Output' -and
        $run -match 'Remaining Balance' },
    [pscustomobject]@{ Name = "Run.UsedGoodsGroupsUoms"; Pass =
        $run -match 'ProcessUsedGoodsDisplay' -and
        $run -match 'UsedGoodsByNormalizedUom' -and
        $run -match 'FormatUsedGoodsQuantity' -and
        $run -notmatch 'Private Function ProcessUsedGoodsQty' },
    [pscustomobject]@{ Name = "Form.ReadOnlyRoutedInputProjection"; Pass =
        $form -match '"Source Process / Output"' -and
        $form -match '"Remaining Balance"' -and
        $form -match 'mLstManagerCheck.*9' },
    [pscustomobject]@{ Name = "Form.PublicHandlerProof"; Pass =
        $form -match '\|RoutedInputVisible=' -and
        $form -match '\|UpstreamWaitThenReady=' -and
        $form -match '\|GroupedUsedGoods=' },
    [pscustomobject]@{ Name = "Form.NamedChaiFourProcessHandlerProof"; Pass =
        $form -match 'Public Function TestChaiForkConvergenceRunActionContract' -and
        $form -match '\|ChaiFourProcessesCompleted=' -and
        $form -match '\|ChaiFinalBottlingCompleted=' -and
        $form -match '\|ChaiRunNotRestarted=' },
    [pscustomobject]@{ Name = "Packaged.RequiresSlice4axEvidence"; Pass =
        $validator -match 'RoutedInputVisible=True' -and
        $validator -match 'UpstreamWaitThenReady=True' -and
        $validator -match 'GroupedUsedGoods=True' -and
        $validator -match 'ChaiFourProcessesCompleted=True' -and
        $validator -match 'ChaiFinalBottlingCompleted=True' -and
        $validator -match 'ChaiRunNotRestarted=True' }
)

$failed = @($checks | Where-Object { -not $_.Pass })
foreach ($check in $checks) {
    "{0} {1}" -f $(if ($check.Pass) { "PASS" } else { "RED" }), $check.Name
}

"PLAN022_SLICE4AX_SOURCE passed=$($checks.Count - $failed.Count) red=$($failed.Count) total=$($checks.Count)"
if ($failed.Count -gt 0) {
    throw "Plan 022 Slice 4ax source contract RED: $($failed.Name -join ', ')"
}
