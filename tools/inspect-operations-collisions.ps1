[CmdletBinding()]
param(
    [string]$RepoRoot = ".",
    [string]$OutputDirectory = "reports/operations-shadow",
    [string]$ReportTimestampUtc = "",
    [string]$ResolutionsPath = "",
    [switch]$FailOnUnresolved
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

$repo = (Resolve-Path $RepoRoot).Path
if ([string]::IsNullOrWhiteSpace($ReportTimestampUtc)) {
    $ReportTimestampUtc = [DateTime]::UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ")
}
if ([string]::IsNullOrWhiteSpace($ResolutionsPath)) {
    $ResolutionsPath = Join-Path $repo `
        "tools\contracts\operations-shadow-collision-resolutions.json"
}
elseif (-not [IO.Path]::IsPathRooted($ResolutionsPath)) {
    $ResolutionsPath = Join-Path $repo $ResolutionsPath
}
if (-not (Test-Path -LiteralPath $ResolutionsPath -PathType Leaf)) {
    throw "Collision resolution contract not found: $ResolutionsPath"
}
if ([IO.Path]::IsPathRooted($OutputDirectory)) {
    $outputRoot = [IO.Path]::GetFullPath($OutputDirectory)
}
else {
    $outputRoot = Join-Path $repo $OutputDirectory
}
if (-not (Test-Path -LiteralPath $outputRoot)) {
    New-Item -ItemType Directory -Path $outputRoot -Force | Out-Null
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

function Get-RepoRelativePath {
    param([string]$Path)

    $fullPath = [IO.Path]::GetFullPath($Path)
    $prefix = $repo.TrimEnd("\") + "\"
    if (-not $fullPath.StartsWith(
        $prefix,
        [StringComparison]::OrdinalIgnoreCase
    )) {
        throw "Source path is outside the repository: $fullPath"
    }
    return $fullPath.Substring($prefix.Length).Replace("\", "/")
}

function Get-ComponentName {
    param([string]$Text)

    $match = [regex]::Match(
        $Text,
        '(?m)^Attribute VB_Name = "([^"]+)"'
    )
    if (-not $match.Success) {
        return ""
    }
    return $match.Groups[1].Value
}

function Get-RoleFromRelativePath {
    param([string]$RelativePath)

    $parts = $RelativePath.Split("/")
    if ($parts.Count -ge 2 -and $parts[0] -eq "src") {
        return $parts[1]
    }
    return "Unknown"
}

function New-CollisionRows {
    param(
        [object[]]$Rows,
        [string]$Kind
    )

    $collisions = @()
    foreach ($group in @(
        $Rows |
            Group-Object { ([string]$_.Name).ToUpperInvariant() } |
            Where-Object { $_.Count -gt 1 } |
            Sort-Object Name
    )) {
        $firstName = [string]$group.Group[0].Name
        $occurrences = @(
            $group.Group |
                Sort-Object Path, Component |
                ForEach-Object {
                    [ordered]@{
                        role = [string]$_.Role
                        component = [string]$_.Component
                        path = [string]$_.Path
                    }
                }
        )
        $collisions += [ordered]@{
            collisionId = "$Kind`:$firstName"
            name = $firstName
            occurrences = $occurrences
        }
    }
    return @($collisions)
}

$resolutionContract = Get-Content -Raw -LiteralPath $ResolutionsPath |
    ConvertFrom-Json
$excludedNames = @(
    $resolutionContract.shadowExcludedFileNames |
        ForEach-Object { [string]$_ }
)
$resolutionById = @{}
foreach ($resolution in @($resolutionContract.resolutions)) {
    $resolutionById[[string]$resolution.collisionId] = $resolution
}

$sourceRoots = @(
    "src/Operations",
    "src/Receiving",
    "src/Production",
    "src/Shipping"
)
$sourceFiles = @()
foreach ($relativeRoot in $sourceRoots) {
    $sourceRoot = Join-Path $repo $relativeRoot
    if (-not (Test-Path -LiteralPath $sourceRoot -PathType Container)) {
        throw "Operations shadow source root is missing: $relativeRoot"
    }
    $sourceFiles += Get-ChildItem -LiteralPath $sourceRoot -Recurse -File |
        Where-Object {
            $_.Extension -in @(".bas", ".cls", ".frm") -and
            $_.Name -notin $excludedNames
        }
}
$sourceFiles = @($sourceFiles | Sort-Object FullName -Unique)

$componentRows = @()
$procedureRows = @()
foreach ($file in $sourceFiles) {
    $relativePath = Get-RepoRelativePath $file.FullName
    $role = Get-RoleFromRelativePath $relativePath
    $text = Get-Content -Raw -LiteralPath $file.FullName
    $componentName = Get-ComponentName $text
    if ([string]::IsNullOrWhiteSpace($componentName)) {
        throw "VBA component has no Attribute VB_Name: $relativePath"
    }
    $componentRows += [pscustomobject]@{
        Name = $componentName
        Role = $role
        Component = $componentName
        Path = $relativePath
    }

    if ($file.Extension -ne ".bas") {
        continue
    }
    $procedurePattern = (
        '(?im)^(?!\s*(?:Private|Friend)\s+)' +
        '(?:Public\s+)?(?:Static\s+)?' +
        '(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+' +
        '([A-Za-z_][A-Za-z0-9_]*)\b'
    )
    foreach ($match in [regex]::Matches($text, $procedurePattern)) {
        $procedureRows += [pscustomobject]@{
            Name = $match.Groups[1].Value
            Role = $role
            Component = $componentName
            Path = $relativePath
        }
    }
}

$buildText = Get-Content -Raw -LiteralPath (
    Join-Path $repo "tools\build-xlam.ps1"
)
$ribbonRows = @()
foreach ($match in [regex]::Matches(
    $buildText,
    '(?m)^\s*(?:Enabled)?CallbackName\s*=\s*"([^"]+)"'
)) {
    $ribbonRows += [pscustomobject]@{
        Name = $match.Groups[1].Value
        Role = "Build"
        Component = "modRibbonGenerated"
        Path = "tools/build-xlam.ps1"
    }
}

$componentCollisions = @(New-CollisionRows $componentRows "COMPONENT")
$procedureCollisions = @(
    New-CollisionRows $procedureRows "PUBLIC_PROCEDURE"
)
$ribbonCollisions = @(
    New-CollisionRows $ribbonRows "RIBBON_CALLBACK"
)
$allCollisions = @(
    $componentCollisions +
    $procedureCollisions +
    $ribbonCollisions
)
$resolved = @()
$unresolved = @()
foreach ($collision in $allCollisions) {
    if ($resolutionById.ContainsKey([string]$collision.collisionId)) {
        $resolution = $resolutionById[[string]$collision.collisionId]
        $resolved += [ordered]@{
            collisionId = [string]$collision.collisionId
            disposition = [string]$resolution.disposition
            reason = [string]$resolution.reason
        }
    }
    else {
        $unresolved += $collision
    }
}

$report = [ordered]@{
    schemaVersion = "1.0.0"
    reportType = "operations-shadow-collisions"
    generatedAtUtc = $ReportTimestampUtc
    sourceRoots = $sourceRoots
    excludedFileNames = $excludedNames
    summary = [ordered]@{
        componentCount = $componentRows.Count
        publicProcedureCount = $procedureRows.Count
        ribbonCallbackCount = $ribbonRows.Count
        componentCollisionCount = $componentCollisions.Count
        publicProcedureCollisionCount = $procedureCollisions.Count
        ribbonCallbackCollisionCount = $ribbonCollisions.Count
        resolvedCollisionCount = $resolved.Count
        unresolvedCollisionCount = $unresolved.Count
    }
    collisions = [ordered]@{
        components = $componentCollisions
        publicProcedures = $procedureCollisions
        ribbonCallbacks = $ribbonCollisions
    }
    resolvedCollisions = $resolved
    unresolvedCollisions = $unresolved
}

$jsonPath = Join-Path $outputRoot "collision-report.json"
$markdownPath = Join-Path $outputRoot "collision-report.md"
Write-Utf8NoBom -Path $jsonPath -Content (
    $report | ConvertTo-Json -Depth 100
)

$lines = @(
    "# Operations Shadow Collision Report",
    "",
    "- Generated: $ReportTimestampUtc",
    "- Components: $($componentRows.Count)",
    "- Public standard-module procedures: $($procedureRows.Count)",
    "- Ribbon callbacks inspected: $($ribbonRows.Count)",
    "- Component collision groups: $($componentCollisions.Count)",
    "- Public-procedure collision groups: $($procedureCollisions.Count)",
    "- Ribbon callback collision groups: $($ribbonCollisions.Count)",
    "- Explicitly resolved current collisions: $($resolved.Count)",
    "- Unresolved collisions: $($unresolved.Count)",
    "",
    "The shadow import excludes only the reviewed standalone startup wrappers",
    "and the unreferenced legacy search-form template recorded in the",
    "machine-readable resolution contract.",
    ""
)
if ($allCollisions.Count -eq 0) {
    $lines += "No collisions remain in the Operations shadow import set."
}
else {
    $lines += "| Collision | State |"
    $lines += "|---|---|"
    foreach ($collision in $allCollisions) {
        $state = if ($resolutionById.ContainsKey(
            [string]$collision.collisionId
        )) { "RESOLVED" } else { "UNRESOLVED" }
        $lines += "| $($collision.collisionId) | $state |"
    }
}
Write-Utf8NoBom -Path $markdownPath -Content ($lines -join "`n")

Write-Host "Operations shadow collision inspection complete."
Write-Host "Components: $($componentRows.Count)"
Write-Host "Public procedures: $($procedureRows.Count)"
Write-Host "Ribbon callbacks: $($ribbonRows.Count)"
Write-Host "Unresolved collisions: $($unresolved.Count)"
Write-Host "Output: $outputRoot"

if ($FailOnUnresolved -and $unresolved.Count -gt 0) {
    throw "Operations shadow collision report contains $($unresolved.Count) unresolved collision(s)."
}
