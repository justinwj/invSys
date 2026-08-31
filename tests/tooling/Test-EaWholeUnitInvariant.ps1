[CmdletBinding()]
param([string]$RepoRoot = ".")

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$docs = (Resolve-Path -LiteralPath (Join-Path $repo "..\invSys_docs")).Path

function Read-Text([string]$Path) {
    Get-Content -Raw -LiteralPath $Path
}

$uom = Read-Text (Join-Path $repo "src\Core\Modules\modUomSettings.bas")
$domain = Read-Text (Join-Path $repo "src\InventoryDomain\Modules\modInventoryApply.bas")
$run = Read-Text (Join-Path $repo "src\Production\Modules\modProductionReusableRun.bas")
$form = Read-Text (Join-Path $repo "src\Production\Forms\frmProduction.frm")
$production = Read-Text (Join-Path $repo "src\Production\Modules\mProduction.bas")
$validator = Read-Text (Join-Path $repo "tools\validate_plan022_packaged_launchers.ps1")
$spec = Read-Text (Join-Path $docs "0 plan docs\xlam_invSys\invSys-Design-v4.11.md")
$plan = Read-Text (Join-Path $docs "expert guidance docs\022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md")
$controls = Read-Text (Join-Path $docs "0 plan docs\xlam_invSys\invSys-Controls-v1.md")

$checks = @(
    [pscustomobject]@{ Name = "Docs.EaIsWholeOnly"; Pass =
        $spec -match 'operator-entered `ea`' -and
        $plan -match "case-insensitively after UOM normalization" -and
        $controls -match 'Normalized `EA` is whole-unit' },
    [pscustomobject]@{ Name = "Core.NormalizesAndRejectsFractionalEa"; Pass =
        $uom -match "Public Function UomRequiresWholeQuantity" -and
        $uom -match "Public Function ValidateQuantityForUom" -and
        $uom -match "NormalizeUomName\(uomName\)" },
    [pscustomobject]@{ Name = "Domain.GuardsReceiveAndPayload"; Pass =
        $domain -match "ValidateWholeUnitQuantityApply" -and
        $domain -match "BuildReceiveLines" -and
        $domain -match "BuildPayloadLines" -and
        $domain -match "CatalogUomForSkuApply" },
    [pscustomobject]@{ Name = "Domain.PublicFunctionalProof"; Pass =
        $domain -match "Public Function InventoryEaWholeUnitContractForAutomation" -and
        $domain -match "FRACTIONAL_EA_QTY" },
    [pscustomobject]@{ Name = "Production.PublicHandlersRejectFractionalEa"; Pass =
        $form -match "WriteRequirementEditorToList" -and
        $form -match "WriteOutputEditorToList" -and
        $form -match "WriteConnectionEditorToList" -and
        $form -match "ValidateQuantityForUom" -and
        $run -match "ApplyReusableRunStockAllocation" -and
        $run -match "StageReusableRunActualOutput" -and
        $run -match "ValidateQuantityForUom" },
    [pscustomobject]@{ Name = "Production.PackagedRequirementHandlerProof"; Pass =
        $form -match "Public Function TestEaWholeUnitActionContract" -and
        $form -match "mBtnProcessRequirementAdd_Click" -and
        $production -match "RunEaWholeUnitActionContractTest" -and
        $validator -match "PRODUCTION_EA_WHOLE_UNIT" }
)

$failed = @($checks | Where-Object { -not $_.Pass })
foreach ($check in $checks) {
    "{0} {1}" -f $(if ($check.Pass) { "PASS" } else { "RED" }), $check.Name
}

"EA_WHOLE_UNIT_SOURCE passed=$($checks.Count - $failed.Count) red=$($failed.Count) total=$($checks.Count)"
if ($failed.Count -gt 0) {
    throw "EA whole-unit invariant RED: $($failed.Name -join ', ')"
}
