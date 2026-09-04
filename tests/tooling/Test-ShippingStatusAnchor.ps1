[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

$repo = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path
$formPath = Join-Path $repo "src\Shipping\Forms\frmShipmentsTally.frm"
$modulePath = Join-Path $repo "src\Shipping\Modules\modTS_Shipments.bas"
$resultPath = Join-Path $repo "tests\unit\shipping_status_anchor_results.md"
$formText = Get-Content -Raw -LiteralPath $formPath
$moduleText = Get-Content -Raw -LiteralPath $modulePath
$checks = @(
    [pscustomobject]@{
        Name = "Shipping.PublicLauncher"
        Passed = $moduleText -match 'Public Sub BtnOpenShipmentsForm\(\)'
        Contract = "The protecting packaged path begins at the operator's public Shipping launcher."
    },
    [pscustomobject]@{
        Name = "Shipping.StatusFixedTopAnchor"
        Passed = $formText -match 'mAnchors\.Add\s+mTxtStatus,\s*ANCHOR_LEFT\s+Or\s+ANCHOR_TOP\s+Or\s+ANCHOR_RIGHT'
        Contract = "The message/status box remains at its established Top and fixed Height while stretching horizontally."
    },
    [pscustomobject]@{
        Name = "Shipping.StatusNotBottomAnchored"
        Passed = $formText -notmatch 'mAnchors\.Add\s+mTxtStatus[^\r\n]*ANCHOR_BOTTOM'
        Contract = "Height resize must not translate the status control below Search Boxes."
    },
    [pscustomobject]@{
        Name = "Shipping.PackagedGeometrySeam"
        Passed = ($formText -match 'Public Function TestStatusAnchorAfterResize') -and
            ($moduleText -match 'Public Function RunShippingStatusAnchorTest') -and
            ($moduleText -match 'BtnOpenShipmentsForm')
        Contract = "Packaged geometry proof exercises the same public launcher and the actual form anchor manager."
    }
)
$passed = @($checks | Where-Object Passed).Count
$failed = $checks.Count - $passed
$lines = @(
    "# Shipping Status Anchor Results", "",
    "- Passed: $passed", "- Failed: $failed", "",
    "| Check | Result | Contract |", "|---|---|---|"
)
foreach ($check in $checks) {
    $result = if ($check.Passed) { "PASS" } else { "FAIL" }
    $lines += "| $($check.Name) | $result | $($check.Contract) |"
}
[IO.File]::WriteAllText($resultPath, (($lines -join "`n") + "`n"), (New-Object Text.UTF8Encoding($false)))
Write-Host "SHIPPING_STATUS_ANCHOR_RESULTS=$resultPath"
Write-Host "PASSED=$passed FAILED=$failed TOTAL=$($checks.Count)"
if ($failed -gt 0) { exit 1 }
