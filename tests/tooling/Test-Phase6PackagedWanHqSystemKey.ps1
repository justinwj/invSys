[CmdletBinding()]
param([string]$RepoRoot = ".")

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$validator = Get-Content -Raw -LiteralPath (Join-Path $repo "tools\validate_phase6_packaged_wan_hq.ps1")
$aggregator = Get-Content -Raw -LiteralPath (Join-Path $repo "src\Core\Modules\modHqAggregator.bas")

$checks = @(
    [pscustomobject]@{ Name = "Fixture.GeneratesReceiveKeyThroughPackagedCore"; Pass =
        $validator -match 'modRoleEventWriter\.CreateSystemKey' },
    [pscustomobject]@{ Name = "Fixture.InboxSchemaCarriesSystemKey"; Pass =
        $validator -match '"System_Key", "SKU", "Qty"' },
    [pscustomobject]@{ Name = "Fixture.PersistsGeneratedKeyInReceiveEvent"; Pass =
        $validator -match 'function Seed-InboxReceiveRowOpen[\s\S]*?\$systemKey = \[string\]\(Run-WorkbookMacro[\s\S]*?modRoleEventWriter\.CreateSystemKey[\s\S]*?"System_Key" = \$systemKey' },
    [pscustomobject]@{ Name = "Aggregator.SumsDurableEntitiesByWarehouseAndSku"; Pass =
        $validator -match 'function Get-WarehouseSkuQtyOnHand' -and
        $validator -match '\$qtyA2 = Get-WarehouseSkuQtyOnHand' -and
        $validator -match '\$qtyB2 = Get-WarehouseSkuQtyOnHand' -and
        $validator -match '\$qtyB2 -eq 11' },
    [pscustomobject]@{ Name = "Processor.AcceptsAnyPositiveAppliedCountWithoutPoison"; Pass =
        $validator -match 'Applied=\[1-9\]\\d\*; SkipDup=0; Poison=0' },
    [pscustomobject]@{ Name = "Aggregator.PreservesExactSystemKeyInGlobalRows"; Pass =
        $aggregator -match 'snapHeaders = Array\("WarehouseId", "System_Key", "SKU"' -and
        $aggregator -match 'entry\("System_Key"\) = GetCellByColumnHq\(lo, rowIndex, "System_Key"\)' -and
        $aggregator -match 'systemKey = SafeTrimHq\(GetCellByColumnHq\(lo, i, "System_Key"\)\)' -and
        $aggregator -match 'WarehouseId"\)\) & "\|" & systemKey' },
    [pscustomobject]@{ Name = "Aggregator.PackagedProofRequiresVisibleDistinctKeys"; Pass =
        $validator -match 'function Get-WarehouseSkuEntityCount' -and
        $validator -match '\$globalIdentity1 = \(Get-ColumnIndexSafe -ListObject \$loGlobal1 -ColumnName "System_Key"\) -gt 0' -and
        $validator -match '\$globalIdentity2 = \(Get-ColumnIndexSafe -ListObject \$loGlobal2 -ColumnName "System_Key"\) -gt 0' -and
        $validator -match '\$entityCountB2 -eq 2' }
)

$failed = @($checks | Where-Object { -not $_.Pass })
foreach ($check in $checks) {
    "{0} {1}" -f $(if ($check.Pass) { "PASS" } else { "RED" }), $check.Name
}

"PHASE6_PACKAGED_WAN_HQ_SYSTEM_KEY_SOURCE passed=$($checks.Count - $failed.Count) red=$($failed.Count) total=$($checks.Count)"
if ($failed.Count -gt 0) {
    throw "Phase 6 packaged WAN/HQ System_Key fixture RED: $($failed.Name -join ', ')"
}
