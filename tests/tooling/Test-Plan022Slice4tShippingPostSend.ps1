[CmdletBinding()]
param(
    [string]$RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
)

$ErrorActionPreference = "Stop"
$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$moduleText = Get-Content -Raw -LiteralPath (Join-Path $repo "src/Shipping/Modules/modTS_Shipments.bas")
$formText = Get-Content -Raw -LiteralPath (Join-Path $repo "src/Shipping/Forms/frmShipmentsTally.frm")

function Procedure-Text {
    param(
        [string]$Text,
        [string]$Name
    )

    [regex]::Match(
        $Text,
        "(?ms)^(?:Public|Private) (?:Function|Sub) $([regex]::Escape($Name))\b.*?^End (?:Function|Sub)"
    ).Value
}

$runtimeBody = Procedure-Text -Text $moduleText -Name "RunShippingRuntimeQueueRefresh"
$sendBody = Procedure-Text -Text $formText -Name "mBtnSend_Click"
$formActionBody = Procedure-Text -Text $formText -Name "TestRunShipmentsSentActionForWorkbook"
$readModelCall = $runtimeBody.IndexOf("ShipmentsFormRefreshReadModelForWorkbook", [System.StringComparison]::OrdinalIgnoreCase)
$processorCall = $runtimeBody.IndexOf("modProcessor.RunBatch", [System.StringComparison]::OrdinalIgnoreCase)
$loadStateCall = $sendBody.IndexOf("LoadShipmentState", [System.StringComparison]::OrdinalIgnoreCase)
$loadShippablesCall = $sendBody.IndexOf("LoadShippables", [System.StringComparison]::OrdinalIgnoreCase)
$projectedCall = $sendBody.IndexOf("RefreshProjectedShippableInventory", [System.StringComparison]::OrdinalIgnoreCase)

$checks = @(
    [pscustomobject]@{
        Check = "Shipping.PostSend.CanonicalReadModel"
        Passed = ($processorCall -ge 0) -and ($readModelCall -gt $processorCall)
        Contract = "The Shipping runtime boundary refreshes the captured operator workbook from canonical inventory after the processor applies queued work."
    },
    [pscustomobject]@{
        Check = "Shipping.PostSend.ReloadsShippables"
        Passed = ($loadStateCall -ge 0) -and ($loadShippablesCall -gt $loadStateCall) -and ($projectedCall -gt $loadShippablesCall)
        Contract = "The real Shipments Sent callback reloads shippables after canonical refresh and only then derives visible projected inventory."
    },
    [pscustomobject]@{
        Check = "Shipping.PostSend.PublicFormEvidence"
        Passed = ($formActionBody -match 'mBtnSend_Click') -and
            ($formActionBody -match 'VisibleNas=') -and
            ($formActionBody -match 'VisibleProjected=') -and
            ($formActionBody -match 'VisibleLocked=')
        Contract = "The packaged form-action test reports the same NAS, Projected, and Locked values shown to the operator after Shipments Sent."
    },
    [pscustomobject]@{
        Check = "Shipping.PostSend.NoDuplicateStageClear"
        Passed = ([regex]::Matches((Procedure-Text -Text $moduleText -Name "BtnShipmentsSent"), 'ClearShipmentStageAfterRefresh').Count -le 1)
        Contract = "The legacy Shipments Sent callback does not repeat the same three-attempt stage cleanup after a successful runtime refresh."
    }
)

$passed = @($checks | Where-Object Passed).Count
$failed = $checks.Count - $passed
$resultPath = Join-Path $repo "tests/integration/plan022_slice4t_shipping_post_send_results.md"
$lines = @(
    "# Plan 022 Slice 4t Shipping Post-Send Results",
    "",
    "- Passed: $passed",
    "- Failed: $failed",
    "",
    "| Check | Result | Contract |",
    "|---|---|---|"
)
foreach ($check in $checks) {
    $lines += "| $($check.Check) | $(if ($check.Passed) { 'PASS' } else { 'FAIL' }) | $($check.Contract) |"
}
Set-Content -LiteralPath $resultPath -Value $lines -Encoding UTF8

$checks | Format-Table Check, Passed -AutoSize
Write-Host ("Plan 022 Slice 4t Shipping post-send contract: {0} passed, {1} failed" -f $passed, $failed)
if ($failed -gt 0) { exit 1 }
