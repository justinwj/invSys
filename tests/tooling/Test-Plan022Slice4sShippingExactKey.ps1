[CmdletBinding()]
param(
    [string]$RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
)

$ErrorActionPreference = "Stop"
$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$modulePath = Join-Path $repo "src/Shipping/Modules/modTS_Shipments.bas"
$formPath = Join-Path $repo "src/Shipping/Forms/frmShipmentsTally.frm"
$boxingServicePath = Join-Path $repo "src/Shipping/Modules/modBoxingService.bas"
$quietPath = Join-Path $repo "src/Core/Modules/modUiQuiet.bas"
$moduleText = Get-Content -Raw -LiteralPath $modulePath
$formText = Get-Content -Raw -LiteralPath $formPath
$boxingServiceText = Get-Content -Raw -LiteralPath $boxingServicePath
$quietText = Get-Content -Raw -LiteralPath $quietPath

function Procedure-Text([string]$name) {
    [regex]::Match(
        $moduleText,
        "(?ms)^(?:Public|Private) (?:Function|Sub) $([regex]::Escape($name))\b.*?^End (?:Function|Sub)"
    ).Value
}

$commitBody = Procedure-Text "ShipmentsFormCommitLine"
$applyBody = Procedure-Text "ApplyShipmentDeltasLocal"
$applyKeyBody = Procedure-Text "ApplyShipmentDeltasBySystemKey"
$tablePickerBody = Procedure-Text "BuildShippingInventoryPickerItems"
$canonicalPickerBody = Procedure-Text "BuildCanonicalRuntimeInventoryPickerItems"
$commitFormBody = [regex]::Match($formText, '(?ms)^Private Sub CommitCurrentLine\b.*?^End Sub').Value
$saveDesignBody = [regex]::Match($boxingServiceText, '(?ms)^Public Function SaveBoxDesign\b.*?^End Function').Value
$postBoxBody = [regex]::Match($boxingServiceText, '(?ms)^Public Function PostBoxMakerAction\b.*?^End Function').Value

$checks = @(
    [pscustomobject]@{
        Check = "Shipping.Add.PublicAction"
        Passed = ($formText -match 'CommitCurrentLine\s+"ADD"') -and
            ($commitBody -match 'BuildSelectedShipmentRowsDeltas') -and
            ($commitBody -match 'ApplyShipmentDeltasLocal')
        Contract = "The real Shipping Add callback reaches the public ShipmentsFormCommitLine action and its local reserve boundary."
    },
    [pscustomobject]@{
        Check = "Shipping.Add.SystemKeyApply"
        Passed = ($applyBody -match 'ColumnIndex\(invLo,\s*"ROW"\)') -and
            ($applyBody -match 'If\s+colRow\s*=\s*0\s+Then[\s\S]*ApplyShipmentDeltasBySystemKey') -and
            ($applyKeyBody -match 'delta\.Exists\("System_Key"\)') -and
            ($applyKeyBody -match 'FindInvListRowBySystemKey') -and
            ($applyKeyBody -notmatch 'delta\("ROW"\)')
        Contract = "Shipping Add reserves current-schema inventory by immutable System_Key when managed ROW is absent."
    },
    [pscustomobject]@{
        Check = "BoxDesigner.ActiveExactEntities"
        Passed = ($tablePickerBody -match 'seenSystemKeys') -and
            ($tablePickerBody -match 'availableQty\s*<=\s*0') -and
            ($canonicalPickerBody -match 'seenSystemKeys') -and
            ($canonicalPickerBody -match 'availableQty\s*<=\s*0')
        Contract = "Box Designer excludes nonpositive balances and removes only duplicate projections of the same exact System_Key."
    },
    [pscustomobject]@{
        Check = "BoxDesigner.PreservesDistinctIdentity"
        Passed = ($tablePickerBody -match 'result\(outRow,\s*1\)\s*=\s*systemKey') -and
            ($canonicalPickerBody -match 'result\(outRow,\s*1\)\s*=\s*systemKey') -and
            ($tablePickerBody -notmatch 'ITEM_CODE.*seen|itemName.*seen')
        Contract = "Positive entities with different System_Key values remain separate selectable component identities."
    },
    [pscustomobject]@{
        Check = "SavingUi.ActionBoundaries"
        Passed = ($commitFormBody -match 'modUiQuiet\.BeginQuietUi') -and
            ($commitFormBody -match 'modUiQuiet\.EndQuietUi') -and
            ($saveDesignBody -match 'modUiQuiet\.BeginQuietUi') -and
            ($saveDesignBody -match 'modUiQuiet\.EndQuietUi') -and
            ($postBoxBody -match 'modUiQuiet\.BeginQuietUi') -and
            ($postBoxBody -match 'modUiQuiet\.EndQuietUi')
        Contract = "Shipping Add, Box Designer save, and Box Maker post retain one nested quiet UI boundary across required persistence."
    },
    [pscustomobject]@{
        Check = "SavingUi.StatusBarRestored"
        Passed = ($quietText -match 'mPrevDisplayStatusBar') -and
            ($quietText -match 'Application\.DisplayStatusBar\s*=\s*False') -and
            ($quietText -match 'Application\.DisplayStatusBar\s*=\s*mPrevDisplayStatusBar')
        Contract = "Quiet UI hides Excel save-status churn and restores the operator's previous status-bar setting."
    }
)

$resultPath = Join-Path $repo "tests/unit/plan022_slice4s_shipping_exact_key_results.md"
$passed = @($checks | Where-Object Passed).Count
$failed = $checks.Count - $passed
$lines = @(
    "# Plan 022 Slice 4s Shipping Exact-Key Results",
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
Write-Host ("Plan 022 Slice 4s Shipping exact-key contract: {0} passed, {1} failed" -f $passed, $failed)
if ($failed -gt 0) { exit 1 }
