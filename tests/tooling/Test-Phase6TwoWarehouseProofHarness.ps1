[CmdletBinding()]
param([string]$RepoRoot = ".")

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$harnessPath = Join-Path $repo "tools\validate_phase6_packaged_wan_hq.ps1"
$harness = Get-Content -LiteralPath $harnessPath -Raw

$checks = @(
    [pscustomobject]@{
        Check = "Phase6.TwoWarehouse.HarnessExceptionIsEvidence"
        Passed = $harness -match '(?s)\}\s*catch\s*\{.*?Add-ResultRow\s+-Rows\s+\$resultRows\s+-Check\s+"Harness\.Exception"\s+-Passed\s+\$false'
        Contract = "A packaged two-warehouse proof failure must write a failed Harness.Exception result row rather than leave a misleading all-pass report."
    },
    [pscustomobject]@{
        Check = "Phase6.TwoWarehouse.GlobalWorkbookReleasedBeforeCatchup"
        Passed = $harness -match '(?s)Release-ComObject \$loGlobal1.*?Release-ComObject \$loStatus1.*?Release-ComObject \$wsGlobal1.*?Release-ComObject \$wsStatus1.*?Release-ComObject \$wbGlobal1.*?RunHQAggregation'
        Contract = "The first read-only global-snapshot inspection must release all COM references before the same HQ Excel instance rebuilds the catch-up snapshot."
    }
)

$failed = @($checks | Where-Object { -not $_.Passed })
$checks | Format-Table -AutoSize | Out-String | Write-Output
if ($failed.Count -gt 0) {
    throw "Phase 6 two-warehouse proof harness: $($failed.Count) check(s) failed."
}

Write-Output "Phase 6 two-warehouse proof harness: $($checks.Count) passed, 0 failed"
