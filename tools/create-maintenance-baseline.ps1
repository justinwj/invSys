[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$OutputDirectory,

    [string]$ReportTimestampUtc = ""
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = (Resolve-Path (Join-Path $scriptRoot "..")).Path
if ([string]::IsNullOrWhiteSpace($ReportTimestampUtc)) {
    $ReportTimestampUtc = [DateTime]::UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ")
}

function Write-Utf8NoBom {
    param(
        [string]$Path,
        [string]$Content
    )
    $normalized = $Content -replace "`r`n", "`n"
    if (-not $normalized.EndsWith("`n")) {
        $normalized += "`n"
    }
    [IO.File]::WriteAllText(
        $Path,
        $normalized,
        (New-Object Text.UTF8Encoding($false))
    )
}

function Read-JsonFile {
    param([string]$Path)
    return (Get-Content -Raw -LiteralPath $Path | ConvertFrom-Json)
}

function Get-FileSha256Lower {
    param([string]$Path)
    return (
        Get-FileHash -LiteralPath $Path -Algorithm SHA256
    ).Hash.ToLowerInvariant()
}

function Get-PackageKeyFromPath {
    param([string]$Path)
    if ($Path -match '(?i)^src/([^/]+)/') {
        return $matches[1]
    }
    if ($Path -match '(?i)^tests/') {
        return "Tests"
    }
    if ($Path -match '(?i)^tools/') {
        return "DeveloperTooling"
    }
    return "Other"
}

function Get-Workstream {
    param([string[]]$SourcePaths)

    $packages = @(
        $SourcePaths |
            ForEach-Object { Get-PackageKeyFromPath $_ } |
            Sort-Object -Unique
    )
    $rolePackages = @(
        $packages | Where-Object {
            $_ -in @("Receiving", "Production", "Shipping")
        }
    )
    if ($rolePackages.Count -gt 1 -or
        ($rolePackages.Count -eq 1 -and $packages.Count -gt 1)) {
        return "SHARED_OPERATIONS"
    }
    if ($rolePackages.Count -eq 1) {
        return $rolePackages[0].ToUpperInvariant()
    }
    if ("Core" -in $packages) {
        return "CORE"
    }
    if (@($packages | Where-Object {
        $_ -in @("InventoryDomain", "DesignsDomain")
    }).Count -gt 0) {
        return "DOMAINS"
    }
    if ("Admin" -in $packages) {
        return "ADMIN"
    }
    if ("Tests" -in $packages) {
        return "TESTING"
    }
    return "DEVELOPER_TOOLING"
}

function Get-NextAction {
    param([string]$Disposition)
    switch ($Disposition) {
        "REMOVE" {
            return "Write a focused test for the nearest public entry point, prove compile and regression GREEN, then request reviewed deletion."
        }
        "MOVE_TO_TESTS" {
            return "Move the diagnostic or fixture helper out of runtime packaging, then prove the packaged surface and developer test still pass."
        }
        "SPLIT_MODULE" {
            return "Extract one coherent service behind the current public contract and ratchet the original module below its baseline."
        }
        "REPLACE_DUPLICATE" {
            return "Choose one typed implementation boundary and protect every affected caller before consolidating duplicate bodies."
        }
        "REPLACE_SAME_PROJECT_LATE_BINDING" {
            return "Protect the public action, replace the same-project string dispatch with a direct typed call, and rescan."
        }
        "RETAIN_DYNAMIC_ROOT" {
            return "Retain the procedure and keep its registry or discovered root evidence current."
        }
        "ISOLATE_LEGACY_IMPORT" {
            return "Write the greenfield no-import acceptance test, then remove or isolate any old-business-inventory path it exposes."
        }
        default {
            return "Investigate callers, callbacks, package ownership, and protecting tests before selecting a source change."
        }
    }
}

function ConvertTo-MarkdownCell {
    param([string]$Value)
    return (($Value -replace '\|', '\|') -replace "`r?`n", " ")
}

$scannerPath = Join-Path $scriptRoot "inventory-vba-surface.ps1"
$rootRegistryPath = Join-Path $scriptRoot "contracts\vba-dynamic-roots.json"
$backlogSchemaPath = Join-Path $scriptRoot `
    "contracts\reviewed-cleanup-backlog.schema.json"
foreach ($requiredPath in @($scannerPath, $rootRegistryPath, $backlogSchemaPath)) {
    if (-not (Test-Path -LiteralPath $requiredPath -PathType Leaf)) {
        throw "Required baseline input is missing: $requiredPath"
    }
}

if (-not (Test-Path -LiteralPath $OutputDirectory -PathType Container)) {
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
}
$resolvedOutput = (Resolve-Path -LiteralPath $OutputDirectory).Path

& $scannerPath `
    -SourceRoot (Join-Path $repoRoot "src") `
    -BuildMapPath (Join-Path $scriptRoot "build-xlam.ps1") `
    -RibbonRoot (Join-Path $scriptRoot "build-xlam.ps1") `
    -TestRoot (Join-Path $repoRoot "tests") `
    -RootRegistryPath $rootRegistryPath `
    -OutputDirectory $resolvedOutput `
    -ReportTimestampUtc $ReportTimestampUtc
if (-not $?) {
    throw "Static scanner failed while creating the maintenance baseline."
}

$manifestPath = Join-Path $resolvedOutput "implementation-manifest.json"
$maintenancePath = Join-Path $resolvedOutput "maintenance-candidates.json"
$manifest = Read-JsonFile $manifestPath
$maintenance = Read-JsonFile $maintenancePath
$reviewed = New-Object System.Collections.Generic.List[object]

foreach ($candidate in @($maintenance.candidates)) {
    $sourcePaths = @($candidate.sourcePaths | ForEach-Object { [string]$_ })
    $protectingTests = @(
        $candidate.protectingTests | ForEach-Object { [string]$_ } | Sort-Object -Unique
    )
    $disposition = [string]$candidate.candidateType
    $reviewReason = [string]$candidate.reason
    if ([string]$candidate.componentName -eq "modExportImportAll") {
        $disposition = "MOVE_TO_TESTS"
        $reviewReason += (
            " This source is a developer export/import diagnostic embedded in " +
            "a runtime package and must be isolated from the packaged product."
        )
    }

    $reviewedConfidence = [string]$candidate.confidence
    if ($reviewedConfidence -eq "HIGH" -and $protectingTests.Count -eq 0) {
        $reviewedConfidence = "MEDIUM"
        $reviewReason += (
            " Scanner confidence alone is not deletion authority; no protecting " +
            "test is currently associated with this candidate."
        )
    }

    $reviewStatus = "PLANNED"
    if ($disposition -eq "RETAIN_DYNAMIC_ROOT") {
        $reviewStatus = "RETAIN"
    }
    elseif ($disposition -eq "UNRESOLVED") {
        $reviewStatus = "MANUAL_INVESTIGATION"
    }
    elseif ($protectingTests.Count -eq 0) {
        $reviewStatus = "REQUIRES_PROTECTING_TEST"
    }

    $reviewed.Add([ordered]@{
        id = [string]$candidate.id
        origin = "SCANNER"
        scannerCandidateType = [string]$candidate.candidateType
        disposition = $disposition
        scannerConfidence = [string]$candidate.confidence
        reviewedConfidence = $reviewedConfidence
        reviewStatus = $reviewStatus
        workstream = Get-Workstream $sourcePaths
        componentName = [string]$candidate.componentName
        procedureNames = @($candidate.procedureNames | ForEach-Object { [string]$_ })
        sourcePaths = $sourcePaths
        reason = $reviewReason
        protectingTests = $protectingTests
        deletionApproved = $false
        nextAction = Get-NextAction $disposition
    })
}

$reviewed.Add([ordered]@{
    id = "manual:greenfield-old-inventory-import-boundary"
    origin = "MANUAL"
    scannerCandidateType = "MANUAL"
    disposition = "ISOLATE_LEGACY_IMPORT"
    scannerConfidence = "MANUAL"
    reviewedConfidence = "MEDIUM"
    reviewStatus = "REQUIRES_PROTECTING_TEST"
    workstream = "ADMIN"
    componentName = "GreenfieldWarehouseGeneration"
    procedureNames = @(
        "CreateWarehouseFromForm",
        "SeedDemoInventoryForWarehouse"
    )
    sourcePaths = @(
        "src/Admin/Modules/modAdmin.bas",
        "src/Core/Modules/modWarehouseBootstrap.bas"
    )
    reason = (
        "D14 prohibits old business inventory import or ROW-to-System_Key " +
        "mapping from the supported Generate Warehouse and demo-seed paths. " +
        "Slice 4 must prove the boundary before selecting any legacy-path removal."
    )
    protectingTests = @()
    deletionApproved = $false
    nextAction = Get-NextAction "ISOLATE_LEGACY_IMPORT"
})

$reviewed.Add([ordered]@{
    id = "manual:operations-shared-package-boundary"
    origin = "MANUAL"
    scannerCandidateType = "MANUAL"
    disposition = "SPLIT_MODULE"
    scannerConfidence = "MANUAL"
    reviewedConfidence = "MEDIUM"
    reviewStatus = "REQUIRES_PROTECTING_TEST"
    workstream = "SHARED_OPERATIONS"
    componentName = "invSys.Operations"
    procedureNames = @()
    sourcePaths = @(
        "tools/build-xlam.ps1",
        "src/Receiving",
        "src/Production",
        "src/Shipping"
    )
    reason = (
        "D12 requires the three role sources to become one Operations package " +
        "without merging them into a new monolithic module."
    )
    protectingTests = @()
    deletionApproved = $false
    nextAction = (
        "In Slice 5, add the five-package manifest and coexistence RED before " +
        "changing the build map or role package composition."
    )
})

$reviewedCandidates = @(
    $reviewed.ToArray() |
        Sort-Object { ([string]$_.workstream) + "|" + ([string]$_.id) }
)

$workstreamReasons = [ordered]@{
    RECEIVING = "Receiving-owned forms, services, and role package source."
    PRODUCTION = "Production-owned forms, services, and role package source."
    SHIPPING = "Shipping and Boxing forms, services, and role package source."
    SHARED_OPERATIONS = "Cross-role or future invSys.Operations packaging work."
    CORE = "Headless shared runtime and developer-support source in Core."
    DOMAINS = "Inventory and Designs Domain authority source."
    ADMIN = "Administrative setup, lifecycle, and developer-support source."
    DEVELOPER_TOOLING = "Build, scan, report, and other developer-only tooling."
    TESTING = "Test harness and fixture source that must remain outside runtime packages."
}
$workstreams = New-Object System.Collections.Generic.List[object]
foreach ($name in @($workstreamReasons.Keys)) {
    $workstreams.Add([ordered]@{
        name = $name
        candidateCount = @($reviewedCandidates | Where-Object {
            $_.workstream -eq $name
        }).Count
        reason = [string]$workstreamReasons[$name]
    })
}

$oversizedModules = New-Object System.Collections.Generic.List[object]
foreach ($component in @($manifest.components | Where-Object {
    [int]$_.lineCount -gt [int]$maintenance.ratchets.maxNewModuleLines
} | Sort-Object sourcePath)) {
    $oversizedModules.Add([ordered]@{
        sourcePath = [string]$component.sourcePath
        componentName = [string]$component.name
        packageKey = Get-PackageKeyFromPath ([string]$component.sourcePath)
        baselineLineCount = [int]$component.lineCount
        growthAllowedWithoutException = $false
    })
}

$deletionDispositions = @(
    "REMOVE", "REPLACE_DUPLICATE", "ISOLATE_LEGACY_IMPORT"
)
$highConfidenceDeletionCount = @($reviewedCandidates | Where-Object {
    $_.reviewedConfidence -eq "HIGH" -and
    $_.disposition -in $deletionDispositions
}).Count

$backlog = [ordered]@{
    schemaVersion = "1.0.0"
    reportType = "reviewed-cleanup-backlog"
    generatedAtUtc = $ReportTimestampUtc
    sourceReports = [ordered]@{
        implementationManifestSha256 = Get-FileSha256Lower $manifestPath
        maintenanceCandidatesSha256 = Get-FileSha256Lower $maintenancePath
    }
    policy = [ordered]@{
        automaticDeletionAllowed = $false
        compileRequiredBeforeDeletion = $true
        protectingRegressionRequiredBeforeDeletion = $true
        scannerConfidenceIsDeletionAuthority = $false
    }
    summary = [ordered]@{
        scannerCandidateCount = @($maintenance.candidates).Count
        reviewedCandidateCount = $reviewedCandidates.Count
        manualCandidateCount = @($reviewedCandidates | Where-Object {
            $_.origin -eq "MANUAL"
        }).Count
        approvedDeletionCount = 0
        highConfidenceDeletionCount = $highConfidenceDeletionCount
    }
    workstreams = $workstreams.ToArray()
    ratchets = [ordered]@{
        maxNewModuleLines = [int]$maintenance.ratchets.maxNewModuleLines
        maxNewProcedureLines = [int]$maintenance.ratchets.maxNewProcedureLines
        allowSameProjectApplicationRunGrowth =
            [bool]$maintenance.ratchets.allowSameProjectApplicationRunGrowth
        allowUnresolvedDynamicCallGrowth =
            [bool]$maintenance.ratchets.allowUnresolvedDynamicCallGrowth
        allowDuplicateBodyGrowth =
            [bool]$maintenance.ratchets.allowDuplicateBodyGrowth
        oversizedModules = $oversizedModules.ToArray()
    }
    candidates = $reviewedCandidates
}

$backlogJsonPath = Join-Path $resolvedOutput "reviewed-cleanup-backlog.json"
Write-Utf8NoBom -Path $backlogJsonPath -Content (
    $backlog | ConvertTo-Json -Depth 100
)

& (Join-Path $scriptRoot "validate-json-contract.ps1") `
    -JsonPath $backlogJsonPath `
    -SchemaPath $backlogSchemaPath
if (-not $?) {
    throw "Reviewed cleanup backlog did not satisfy its schema."
}

$markdown = New-Object System.Collections.Generic.List[string]
$markdown.Add("# invSys Reviewed Cleanup Backlog")
$markdown.Add("")
$markdown.Add("- Schema: 1.0.0")
$markdown.Add("- Baseline: " + $ReportTimestampUtc)
$markdown.Add("- Scanner candidates: " + $backlog.summary.scannerCandidateCount)
$markdown.Add("- Reviewed candidates: " + $backlog.summary.reviewedCandidateCount)
$markdown.Add("- Approved deletions: 0")
$markdown.Add("- Automatic deletion allowed: False")
$markdown.Add("")
$markdown.Add("## Workstreams")
$markdown.Add("")
$markdown.Add("| Workstream | Candidates | Scope |")
$markdown.Add("|---|---:|---|")
foreach ($workstream in @($backlog.workstreams)) {
    $markdown.Add(
        "| $($workstream.name) | $($workstream.candidateCount) | " +
        "$(ConvertTo-MarkdownCell ([string]$workstream.reason)) |"
    )
}
$markdown.Add("")
$markdown.Add("## Module-growth ratchets")
$markdown.Add("")
$markdown.Add(
    "- New modules: at most $($backlog.ratchets.maxNewModuleLines) lines."
)
$markdown.Add(
    "- New procedures: at most $($backlog.ratchets.maxNewProcedureLines) lines."
)
$markdown.Add("- Existing oversized modules may not grow without an explicit exception.")
$markdown.Add("- Same-project Application.Run, unresolved dynamic calls, and duplicate bodies may not grow.")
$markdown.Add("")
$markdown.Add("| Source | Package | Baseline lines |")
$markdown.Add("|---|---|---:|")
foreach ($module in @($backlog.ratchets.oversizedModules)) {
    $markdown.Add(
        "| $($module.sourcePath) | $($module.packageKey) | " +
        "$($module.baselineLineCount) |"
    )
}
$markdown.Add("")
$markdown.Add("## Reviewed candidates")
$markdown.Add("")
$markdown.Add("| ID | Workstream | Disposition | Confidence | Status |")
$markdown.Add("|---|---|---|---|---|")
foreach ($candidate in @($backlog.candidates)) {
    $markdown.Add(
        "| $(ConvertTo-MarkdownCell ([string]$candidate.id)) | " +
        "$($candidate.workstream) | $($candidate.disposition) | " +
        "$($candidate.reviewedConfidence) | $($candidate.reviewStatus) |"
    )
}

Write-Utf8NoBom `
    -Path (Join-Path $resolvedOutput "reviewed-cleanup-backlog.md") `
    -Content ($markdown -join "`n")

Write-Host "invSys maintenance baseline complete."
Write-Host ("Reviewed candidates: " + $reviewedCandidates.Count)
Write-Host ("Oversized module ratchets: " + $oversizedModules.Count)
Write-Host ("Output: " + $resolvedOutput)
