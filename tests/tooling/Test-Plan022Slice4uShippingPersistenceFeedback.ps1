[CmdletBinding()]
param(
    [string]$RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
)

$ErrorActionPreference = "Stop"
$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$moduleText = Get-Content -Raw -LiteralPath (Join-Path $repo "src/Shipping/Modules/modTS_Shipments.bas")
$formText = Get-Content -Raw -LiteralPath (Join-Path $repo "src/Shipping/Forms/frmShipmentsTally.frm")
$processorText = Get-Content -Raw -LiteralPath (Join-Path $repo "src/Core/Modules/modProcessor.bas")

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

$commitBody = Procedure-Text -Text $moduleText -Name "ShipmentsFormCommitLine"
$sendBody = Procedure-Text -Text $moduleText -Name "ShipmentsFormRunShipmentsSentRows"
$batchUpsertBody = Procedure-Text -Text $moduleText -Name "UpsertShippingReservationRows"
$summaryBody = Procedure-Text -Text $moduleText -Name "AppendShippingPersistenceSummary"
$formCommitBody = Procedure-Text -Text $formText -Name "CommitCurrentLine"
$formSendBody = Procedure-Text -Text $formText -Name "mBtnSend_Click"

$checks = @(
    [pscustomobject]@{
        Check = "Shipping.Persistence.AddSummary"
        Passed = ($formText -match 'CommitCurrentLine\s+"ADD"') -and
            ($formCommitBody -match 'ShipmentsFormCommitLine') -and
            ($formCommitBody -match 'RefreshAfterAction\s+report') -and
            ($commitBody -match 'AppendShippingPersistenceSummary') -and
            ($summaryBody -match 'Persistence summary:') -and
            ($summaryBody -match 'warehouse inbox saved') -and
            ($summaryBody -match 'reservation ledger saved')
        Contract = "The real Shipping Add action reports its required durable writes once in the form status/message output."
    },
    [pscustomobject]@{
        Check = "Shipping.Persistence.SendSummary"
        Passed = ($formSendBody -match 'ExecuteShipmentsSent') -and
            ($formSendBody -match 'ShowStatus\s+report') -and
            ($sendBody -match 'AppendShippingPersistenceSummary') -and
            ($summaryBody -match 'processor durability saves=')
        Contract = "The real Shipments Sent action reports the queued event, reservation completion, and processor durability count once."
    },
    [pscustomobject]@{
        Check = "Shipping.Persistence.BatchedReservations"
        Passed = ($batchUpsertBody -match 'OpenCurrentShippingReservationsWorkbook') -and
            ([regex]::Matches($batchUpsertBody, '\bwb\.Save\b').Count -eq 1) -and
            ($batchUpsertBody -match 'For\s+i\s*=\s*LBound\(rowIndexes\)\s+To\s+UBound\(rowIndexes\)')
        Contract = "A multi-row Shipping action opens and saves the reservation ledger once rather than once per selected row."
    },
    [pscustomobject]@{
        Check = "Shipping.Persistence.RequiredProcessorDurability"
        Passed = ($processorText -match 'EventPersistenceSaves=') -and
            ($processorText -match 'eventPersistenceSaves\s*=\s*eventPersistenceSaves\s*\+\s*1')
        Contract = "Consolidated operator feedback does not remove the processor durability-save contract."
    }
)

$passed = @($checks | Where-Object Passed).Count
$failed = $checks.Count - $passed
$resultPath = Join-Path $repo "tests/integration/plan022_slice4u_shipping_persistence_feedback_results.md"
$lines = @(
    "# Plan 022 Slice 4u Shipping Persistence Feedback Results",
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
Write-Host ("Plan 022 Slice 4u Shipping persistence feedback contract: {0} passed, {1} failed" -f $passed, $failed)
if ($failed -gt 0) { exit 1 }
