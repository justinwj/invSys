[CmdletBinding()]
param(
    [string]$RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
)

$ErrorActionPreference = "Stop"
$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$reviewPath = Join-Path $repo "tests/tooling/slice12-reviewed-cleanup.json"
$manifestPath = Join-Path $repo "reports/static-baseline/implementation-manifest.json"
$candidatesPath = Join-Path $repo "reports/static-baseline/maintenance-candidates.json"
$shippingModulePath = Join-Path $repo "src/Shipping/Modules/modTS_Shipments.bas"
$boxingServicePath = Join-Path $repo "src/Shipping/Modules/modBoxingService.bas"
$dynamicRootsPath = Join-Path $repo "tools/contracts/vba-dynamic-roots.json"

$review = Get-Content -Raw -LiteralPath $reviewPath | ConvertFrom-Json
$manifest = Get-Content -Raw -LiteralPath $manifestPath | ConvertFrom-Json
$candidates = Get-Content -Raw -LiteralPath $candidatesPath | ConvertFrom-Json
$shippingText = Get-Content -Raw -LiteralPath $shippingModulePath
$boxingText = Get-Content -Raw -LiteralPath $boxingServicePath

$results = [System.Collections.Generic.List[object]]::new()
function Add-Check {
    param([string]$Name, [bool]$Passed, [string]$Detail)
    $results.Add([pscustomobject]@{
        Check = $Name
        Passed = $Passed
        Detail = $Detail
    }) | Out-Null
}

$allReviewed = $true
$allAbsent = $true
foreach ($deletion in $review.highConfidenceDeletions) {
    if ($deletion.scannerConfidence -ne "HIGH" -or
        $deletion.procedures.Count -eq 0 -or
        $review.protectingTests.Count -eq 0) {
        $allReviewed = $false
    }
    foreach ($procedureName in $deletion.procedures) {
        $pattern = "(?im)^\s*(Public|Private|Friend)?\s*(Sub|Function|Property\s+(Get|Let|Set))\s+$([regex]::Escape($procedureName))\b"
        if ($shippingText -match $pattern) { $allAbsent = $false }
    }
}
Add-Check "Slice12.Deletions.ReviewedAndProtected" $allReviewed `
    "Every deletion must have scanner confidence and protecting-test evidence."
Add-Check "Slice12.Deletions.HighConfidenceProceduresAbsent" $allAbsent `
    "The reviewed HIGH-confidence Shipping procedures must be absent."

$relocationsPass = $true
foreach ($relocation in $review.relocations) {
    $source = Join-Path $repo $relocation.source
    $destination = Join-Path $repo $relocation.destination
    if ((Test-Path -LiteralPath $source) -or
        -not (Test-Path -LiteralPath $destination -PathType Leaf)) {
        $relocationsPass = $false
    }
}
Add-Check "Slice12.RuntimeDiagnostics.Relocated" $relocationsPass `
    "Developer export/test/diagnostic modules must not remain in runtime Core."

$retiredPass = $true
foreach ($component in $review.retiredComponents) {
    if (Test-Path -LiteralPath (Join-Path $repo $component.source)) {
        $retiredPass = $false
    }
}
Add-Check "Slice12.Shipping.RetiredFormsRemoved" $retiredPass `
    "Separate Box Builder and Box Maker forms must be removed after tab cutover."

Add-Check "Slice12.Shipping.FormTestsUseBoxingService" `
    (($shippingText -notmatch '(?i)\bfrmShippingBoxBuilder\b') -and
     ($shippingText -notmatch '(?i)\bfrmShippingBoxMaker\b') -and
     ($shippingText -match '(?i)modBoxingService\.ProjectedComponentInventoryTextForTest') -and
     ($shippingText -match '(?i)modBoxingService\.RenderedComponentInventoryAfterPendingActionForTest') -and
     ($boxingText -match '(?i)Public\s+Function\s+ProjectedComponentInventoryTextForTest') -and
     ($boxingText -match '(?i)Public\s+Function\s+RenderedComponentInventoryAfterPendingActionForTest')) `
    "Behavior tests formerly hosted by the retired form must use the typed Boxing service."

$currentRootHash = (Get-FileHash -LiteralPath $dynamicRootsPath -Algorithm SHA256).Hash.ToLowerInvariant()
Add-Check "Slice12.DynamicRoots.NoSilencingException" `
    ($currentRootHash -eq $review.before.dynamicRootRegistrySha256) `
    "The dynamic-root registry must not change merely to silence cleanup findings."

$duplicateCount = @($candidates.candidates |
    Where-Object { $_.candidateType -eq "REPLACE_DUPLICATE" }).Count

function Test-ApprovedGrowth {
    param(
        [object]$Approval,
        [int]$CurrentCount
    )

    return (($null -ne $Approval) -and
        ($CurrentCount -ge [int]$Approval.baseline) -and
        ($CurrentCount -le [int]$Approval.ceiling) -and
        (-not [string]::IsNullOrWhiteSpace([string]$Approval.slice)) -and
        (-not [string]::IsNullOrWhiteSpace([string]$Approval.rationale)) -and
        (@($Approval.protectingTests).Count -gt 0))
}

$componentCount = @($manifest.components).Count
$procedureCount = @($manifest.procedures).Count
$candidateCount = @($candidates.candidates).Count
Add-Check "Slice12.Metrics.ComponentsImprove" `
    (($componentCount -lt [int]$review.before.components) -and
     (($componentCount -le [int]$review.approvedComponentGrowth.baseline) -or
      (Test-ApprovedGrowth $review.approvedComponentGrowth $componentCount))) `
    "Runtime component growth must remain within the explicit protected Slice 4x exception."
Add-Check "Slice12.Metrics.ProceduresImprove" `
    (($procedureCount -lt [int]$review.before.procedures) -or
     (Test-ApprovedGrowth $review.approvedProcedureGrowth $procedureCount)) `
    "Runtime procedure count must improve from Slice 11 or remain within an explicit protected slice exception."
Add-Check "Slice12.Metrics.CandidatesImprove" `
    (($candidateCount -lt [int]$review.before.maintenanceCandidates) -or
     (Test-ApprovedGrowth $review.approvedCandidateGrowth $candidateCount)) `
    "Maintenance candidates must improve or remain within an explicit protected Slice 4x exception."
Add-Check "Slice12.Metrics.DuplicatesImprove" `
    (($duplicateCount -le [int]$review.approvedDuplicateGrowth.baseline) -or
     (Test-ApprovedGrowth $review.approvedDuplicateGrowth $duplicateCount)) `
    "Duplicate-body groups must not exceed the reviewed Slice 4x wrapper exception."

$declaredDuplicateIds = @($review.approvedDuplicateGrowth.candidateIds)
$currentDuplicateIds = @($candidates.candidates |
    Where-Object { $_.candidateType -eq "REPLACE_DUPLICATE" } |
    ForEach-Object { $_.id })
$duplicateExceptionPass = ($declaredDuplicateIds.Count -eq
    ([int]$review.approvedDuplicateGrowth.ceiling - [int]$review.approvedDuplicateGrowth.baseline))
foreach ($candidateId in $declaredDuplicateIds) {
    if ($candidateId -notin $currentDuplicateIds) { $duplicateExceptionPass = $false }
}
Add-Check "Slice12.Metrics.DuplicateExceptionReviewed" $duplicateExceptionPass `
    "Every duplicate-body increase must be named explicitly in the Slice 4x exception."

$moduleGrowthPass = @($review.approvedModuleGrowth).Count -gt 0 -and
    @($review.approvedModuleGrowthProtectingTests).Count -gt 0
foreach ($approval in @($review.approvedModuleGrowth)) {
    $component = @($manifest.components | Where-Object { $_.sourcePath -eq $approval.source })
    if ($component.Count -ne 1 -or
        [int]$component[0].lineCount -gt [int]$approval.ceiling -or
        [string]::IsNullOrWhiteSpace([string]$approval.rationale)) {
        $moduleGrowthPass = $false
    }
}
Add-Check "Slice12.Metrics.ModuleGrowthReviewed" $moduleGrowthPass `
    "Every oversized or over-1,000-line Slice 4x module growth path must have a bounded reviewed exception."

$literalTargets = @($manifest.procedures |
    ForEach-Object { $_.literalApplicationRunTargets } |
    Where-Object { $_ }).Count
$unresolvedCalls = @($manifest.procedures |
    ForEach-Object { $_.unresolvedApplicationRunExpressions } |
    Where-Object { $_ }).Count
Add-Check "Slice12.Metrics.NoLateBindingRegression" `
    (($literalTargets -le [int]$review.before.literalApplicationRunTargets) -and
     ($unresolvedCalls -le [int]$review.before.unresolvedDynamicCalls)) `
    "Literal Application.Run and unresolved dynamic-call counts must not increase."

$results | Format-Table -AutoSize
$failed = @($results | Where-Object { -not $_.Passed })
Write-Host ("Slice 12 reviewed cleanup contract: {0} passed, {1} failed" -f
    ($results.Count - $failed.Count), $failed.Count)
if ($failed.Count -gt 0) { exit 1 }
