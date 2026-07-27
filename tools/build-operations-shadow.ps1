[CmdletBinding()]
param(
    [string]$RepoRoot = ".",
    [string]$OutputDirectory = "",
    [string]$EvidenceDirectory = "reports/operations-shadow",
    [string]$ReportTimestampUtc = ""
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

$repo = (Resolve-Path $RepoRoot).Path
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = Join-Path ([IO.Path]::GetTempPath()) (
        "invsys-operations-shadow-" + [Guid]::NewGuid().ToString("N")
    )
}
elseif (-not [IO.Path]::IsPathRooted($OutputDirectory)) {
    $OutputDirectory = Join-Path $repo $OutputDirectory
}
$shadowRoot = [IO.Path]::GetFullPath($OutputDirectory)
$deployRoot = [IO.Path]::GetFullPath(
    (Join-Path $repo "deploy\current")
).TrimEnd("\")
if ([string]::Equals(
    $shadowRoot.TrimEnd("\"),
    $deployRoot,
    [StringComparison]::OrdinalIgnoreCase
)) {
    throw "The Operations shadow cannot target deploy/current."
}
if ([string]::IsNullOrWhiteSpace($ReportTimestampUtc)) {
    $ReportTimestampUtc = [DateTime]::UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ")
}

$legacyNames = @(
    "invSys.Receiving.xlam",
    "invSys.Production.xlam",
    "invSys.Shipping.xlam"
)
$beforeHashes = @{}
foreach ($legacyName in $legacyNames) {
    $legacyPath = Join-Path $deployRoot $legacyName
    if (-not (Test-Path -LiteralPath $legacyPath -PathType Leaf)) {
        throw "Active legacy role package is missing: $legacyName"
    }
    $beforeHashes[$legacyName] = (
        Get-FileHash -LiteralPath $legacyPath -Algorithm SHA256
    ).Hash
}

& (Join-Path $repo "tools\inspect-operations-collisions.ps1") `
    -RepoRoot $repo `
    -OutputDirectory $EvidenceDirectory `
    -ReportTimestampUtc $ReportTimestampUtc `
    -FailOnUnresolved

& (Join-Path $repo "tools\build-xlam.ps1") `
    -RepoRoot $repo `
    -OutputRoot $shadowRoot `
    -IncludeOperationsShadow `
    -Projects @(
        "InventoryDomain",
        "DesignsDomain",
        "OperationsShadow"
    ) `
    -Apply

$expectedShadowFiles = @(
    "invSys.Core.xlam",
    "invSys.Inventory.Domain.xlam",
    "invSys.Designs.Domain.xlam",
    "invSys.Operations.xlam"
)
$missingShadowFiles = @($expectedShadowFiles | Where-Object {
    -not (Test-Path -LiteralPath (Join-Path $shadowRoot $_) -PathType Leaf)
})
if ($missingShadowFiles.Count -gt 0) {
    throw "Shadow build omitted: $($missingShadowFiles -join ', ')"
}

$changedLegacyFiles = @()
foreach ($legacyName in $legacyNames) {
    $legacyPath = Join-Path $deployRoot $legacyName
    $afterHash = (
        Get-FileHash -LiteralPath $legacyPath -Algorithm SHA256
    ).Hash
    if ($afterHash -ne $beforeHashes[$legacyName]) {
        $changedLegacyFiles += $legacyName
    }
}
if ($changedLegacyFiles.Count -gt 0) {
    throw "Shadow build changed active legacy package(s): $($changedLegacyFiles -join ', ')"
}
if (Test-Path -LiteralPath (
    Join-Path $deployRoot "invSys.Operations.xlam"
) -PathType Leaf) {
    throw "Shadow build incorrectly published invSys.Operations.xlam to deploy/current."
}

$resultPath = Join-Path $repo `
    "tests\unit\slice6_shadow_build_results.md"
$lines = @(
    "# Slice 6 Operations Shadow Build Results",
    "",
    "- Result: PASS",
    "- Shadow packages built: $($expectedShadowFiles.Count)",
    "- Unresolved collisions: 0",
    "- Active legacy role packages changed: 0",
    "- Operations packages published to deploy/current: 0",
    "",
    "The disposable shadow output contains Core, both Domain packages, and",
    "invSys.Operations.xlam. Its machine-specific output path is intentionally",
    "omitted from committed evidence."
)
[IO.File]::WriteAllText(
    $resultPath,
    (($lines -join "`n") + "`n"),
    (New-Object Text.UTF8Encoding($false))
)

Write-Host "OPERATIONS_SHADOW_BUILD_OK"
Write-Host "SHADOW_OUTPUT=$shadowRoot"
Write-Host "RESULTS=$resultPath"
