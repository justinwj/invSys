Param(
    [string]$RepoRoot = "."
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$validatorPath = Join-Path $repo "tools/validate_release1_full_chain.ps1"
$adminConsolePath = Join-Path $repo "src/Admin/Modules/modAdminConsole.bas"

$validatorText = if (Test-Path -LiteralPath $validatorPath) {
    Get-Content -LiteralPath $validatorPath -Raw
} else {
    ""
}
$liveValidatorPath = Join-Path $repo "tools/validate_phase6_live_role_workflows.ps1"
$packagedActionText = $validatorText
if (Test-Path -LiteralPath $liveValidatorPath) {
    $packagedActionText += Get-Content -LiteralPath $liveValidatorPath -Raw
}
$adminConsoleText = Get-Content -LiteralPath $adminConsolePath -Raw

$results = [System.Collections.Generic.List[object]]::new()
function Add-Check {
    Param(
        [string]$Name,
        [bool]$Passed,
        [string]$Detail
    )

    $results.Add([pscustomobject]@{
        Check = $Name
        Passed = $Passed
        Detail = $Detail
    })
}

Add-Check "Slice14.Validator.Exists" `
    (Test-Path -LiteralPath $validatorPath) `
    "A dedicated Release 1 validator must own the ordered chain and its combined evidence."

Add-Check "Slice14.Admin.PrimitiveEntryBoundary" `
    (($adminConsoleText -match '(?i)Public\s+Function\s+SeedDemoInventoryForAutomation') -and
     ($validatorText -match '(?i)invSys\.Admin\.xlam') -and
     ($validatorText -match '(?i)modAdminConsole\.BootstrapWarehouseLocalAdmin') -and
     ($validatorText -match '(?i)modAdminConsole\.SeedDemoInventoryForAutomation')) `
    "Fresh warehouse creation and fake inventory seeding must execute through packaged Admin primitive callbacks."

$orderedTokens = @(
    'GenerateFreshWarehouse',
    'SeedDemoInventoryThroughAdmin',
    'ReceiveInventory',
    'ProcessorApplyReceive',
    'RefreshAfterReceive',
    'ProductionTwoBatches',
    'ProductionConsumptionAndOutput',
    'BoxingVersionSelection',
    'ShipmentStagingAndSent',
    'ProcessorApplyShipment',
    'FinalRefresh',
    'RestartAndReconcile'
)
$lastIndex = -1
$ordered = $true
foreach ($token in $orderedTokens) {
    $nextIndex = $validatorText.IndexOf($token, $lastIndex + 1, [System.StringComparison]::Ordinal)
    if ($nextIndex -lt 0) {
        $ordered = $false
        break
    }
    $lastIndex = $nextIndex
}
Add-Check "Slice14.Chain.Ordered" $ordered `
    ("Required ordered phase tokens: " + ($orderedTokens -join " -> "))

Add-Check "Slice14.PackagedRoleActions" `
    (($packagedActionText -match '(?i)RunReceivingConfirmWritesFormActionForTest') -and
     ($packagedActionText -match '(?i)ProductionFormTwoBatchActionReportForTest') -and
     ($packagedActionText -match '(?i)RunRelease1BoxingActionForTest') -and
     ($packagedActionText -match '(?i)RunShipmentsSentFormActionForTest')) `
    "The chain must exercise the same packaged action boundaries used by operators."

$requiredAssertions = @(
    'UniqueSystemKeys',
    'NoRowHeaders',
    'HeaderPersistence',
    'ExactBalancesAndLocations',
    'EventIdentityStatusLogAndReplay',
    'NoNegativeInventory',
    'ProductionBatchState',
    'BoxingBomVersion',
    'LocksReleased',
    'OverlayPreserved',
    'RestartReconciliation',
    'CanonicalWorkbooksHidden',
    'NoDuplicatePackagesOrCallbacks',
    'RuntimeFivePackages',
    'StaticRetiredPathRatchet'
)
$missingAssertions = @($requiredAssertions | Where-Object {
    $validatorText -notmatch [regex]::Escape($_)
})
Add-Check "Slice14.RequiredAssertions" `
    ($missingAssertions.Count -eq 0) `
    $(if ($missingAssertions.Count -eq 0) {
        "All required invariant assertions are named."
    } else {
        "Missing: " + ($missingAssertions -join ", ")
    })

Add-Check "Slice14.RestartUsesSavedRuntime" `
    (($validatorText -match '(?i)\.Save\(') -and
     ($validatorText -match '(?i)\.Close\(') -and
     ($validatorText -match '(?i)Workbooks\.Open\(')) `
    "The validator must save, close, and reopen selected runtime boundaries before reconciliation."

Add-Check "Slice14.RuntimeAndStaticEvidence" `
    (($validatorText -match '(?i)export-invsys-runtime-state\.ps1') -and
     ($validatorText -match '(?i)inventory-vba-surface\.ps1')) `
    "The final gate must include read-only five-package runtime evidence and the static retired-path ratchet."

Add-Check "Slice14.Evidence.RedactsRunIds" `
    ($validatorText.Contains("'(?i)RunId=[^;|\s]+'") -and
     $validatorText -match "RunId=<redacted>") `
    "Committed full-chain evidence must redact generated processor run identifiers."

Add-Check "Slice14.Evidence.RecordsD13Trace" `
    (($validatorText -match "## D13 trace") -and
     ($validatorText -match "Focused RED") -and
     ($validatorText -match "Behavioral RED") -and
     ($validatorText -match "GREEN")) `
    "Committed Slice 14 evidence must record meaningful RED and final GREEN."

$failed = @($results | Where-Object { -not $_.Passed })
foreach ($row in $results) {
    $status = if ($row.Passed) { "PASS" } else { "FAIL" }
    Write-Host ("[{0}] {1} - {2}" -f $status, $row.Check, $row.Detail)
}
Write-Host ("Slice 14 full-chain contract: {0} passed, {1} failed" -f
    ($results.Count - $failed.Count), $failed.Count)

if ($failed.Count -gt 0) {
    exit 1
}
