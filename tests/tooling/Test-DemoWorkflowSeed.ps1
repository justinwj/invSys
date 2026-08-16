[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

$repo = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path
$sourcePath = Join-Path $repo "src\Admin\Modules\modAdminInventorySeed.bas"
$resultPath = Join-Path $repo "tests\unit\demo_workflow_seed_results.md"
$text = Get-Content -Raw -LiteralPath $sourcePath
$seedRows = [regex]::Matches($text, '(?m)^\s*AddDemoInventoryItem\s+rows,').Count
$checks = @(
    [pscustomobject]@{
        Name = "Seed.CompleteKitCount"; Passed = $seedRows -eq 24
        Contract = "One Admin seed event carries exactly 24 R1 workflow inventory entities, including box-making consumables."
    },
    [pscustomobject]@{
        Name = "Seed.MaterialCoverage"
        Passed = ($text -match '"raw"') -and ($text -match '"wip"') -and
            ($text -match '"shippable"') -and ($text -match '"packaging\.ship"') -and
            ($text -match 'DEMO-PKG-SHIPPING-CARTON') -and ($text -match 'DEMO-PKG-PACKING-TAPE')
        Contract = "The kit covers raw inputs, WIP, shippable goods, and shipping packaging."
    },
    [pscustomobject]@{
        Name = "Seed.D14Identity"
        Passed = ($text -match 'CreateSystemKey\(\)') -and
            ($text -match '"GOOD"') -and ($text -notmatch 'item\("ROW"\)')
        Contract = "Every seed entity receives a new System_Key and GOOD condition without a ROW identity path."
    },
    [pscustomobject]@{
        Name = "Seed.CatalogMetadata"
        Passed = @('ITEM_CODE','ITEM','UOM','DESCRIPTION','CATEGORY') |
            ForEach-Object { $text -match ('item\("' + [regex]::Escape($_) + '"\)') } |
            Where-Object { -not $_ } | Measure-Object | Select-Object -ExpandProperty Count | ForEach-Object { $_ -eq 0 }
        Contract = "The event carries the catalog metadata needed by Receiving, Production, Shipping, and Viewer projections."
    }
)
$passed = @($checks | Where-Object Passed).Count
$failed = $checks.Count - $passed
$lines = @(
    "# Demo Workflow Seed Results", "", "- Passed: $passed", "- Failed: $failed", "",
    "| Check | Result | Contract |", "|---|---|---|"
)
foreach ($check in $checks) {
    $result = if ($check.Passed) { "PASS" } else { "FAIL" }
    $lines += "| $($check.Name) | $result | $($check.Contract) |"
}
[IO.File]::WriteAllText($resultPath, (($lines -join "`n") + "`n"), (New-Object Text.UTF8Encoding($false)))
Write-Host "DEMO_WORKFLOW_SEED_RESULTS=$resultPath"
Write-Host "PASSED=$passed FAILED=$failed TOTAL=$($checks.Count)"
if ($failed -gt 0) { exit 1 }
