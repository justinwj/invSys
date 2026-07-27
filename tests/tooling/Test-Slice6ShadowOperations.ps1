[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path
$buildPath = Join-Path $repoRoot "tools\build-xlam.ps1"
$shadowBuildPath = Join-Path $repoRoot "tools\build-operations-shadow.ps1"
$collisionToolPath = Join-Path $repoRoot "tools\inspect-operations-collisions.ps1"
$resolutionPath = Join-Path $repoRoot `
    "tools\contracts\operations-shadow-collision-resolutions.json"
$validatorPath = Join-Path $repoRoot "tools\validate-operations-shadow.ps1"
$inventoryPath = Join-Path $repoRoot "tools\inventory-vba-surface.ps1"
$manifestPath = Join-Path $repoRoot `
    "reports\static-baseline\implementation-manifest.json"
$resultPath = Join-Path $repoRoot `
    "tests\unit\slice6_shadow_operations_results.md"
$deployRoot = Join-Path $repoRoot "deploy\current"
$rows = New-Object System.Collections.Generic.List[object]

function Add-Check {
    param(
        [string]$Name,
        [bool]$Passed,
        [string]$Contract
    )

    $rows.Add([pscustomobject]@{
        Name = $Name
        Passed = $Passed
        Contract = $Contract
    })
}

$buildText = Get-Content -Raw -LiteralPath $buildPath
$inventoryText = Get-Content -Raw -LiteralPath $inventoryPath
$hasShadowProject = ($buildText -match 'Key\s*=\s*"OperationsShadow"') -and
                    ($buildText -match 'OutputFile\s*=\s*"invSys\.Operations\.xlam"') -and
                    ($buildText -match 'Deployable\s*=\s*\$false')
$hasAllRoleSources = ($buildText -match 'src/Receiving') -and
                     ($buildText -match 'src/Production') -and
                     ($buildText -match 'src/Shipping')
$supportsSelection = ($buildText -match '\[string\[\]\]\$Projects') -and
                     ($buildText -match '\[switch\]\$IncludeOperationsShadow')

Add-Check "Shadow.ProjectDefinition" ($hasShadowProject -and $hasAllRoleSources) `
    "The build map must define one explicitly non-deployable Operations shadow project containing all three role source sets."
Add-Check "Shadow.ProjectSelection" $supportsSelection `
    "The build entry point must select complete projects so the shadow build can avoid publishing unrelated packages."
Add-Check "Shadow.StaticInventoryMultiSource" `
    (($inventoryText -match 'sourceMatches') -and
     ($inventoryText -match 'sourceMatches\.Count\s+-gt\s+0')) `
    "Static maintenance tooling must inventory every source directory in a combined build project."
$manifestPackages = @()
if (Test-Path -LiteralPath $manifestPath -PathType Leaf) {
    $manifest = Get-Content -Raw -LiteralPath $manifestPath | ConvertFrom-Json
    $manifestPackages = @($manifest.packages | ForEach-Object { $_.key })
}
Add-Check "Shadow.StaticInventoryRetainsPackageSet" `
    (("Admin" -in $manifestPackages) -and
     ("OperationsShadow" -in $manifestPackages)) `
    "Adding a combined project must not make the static manifest drop Admin or the shadow package."
Add-Check "Shadow.BuildEntryPoint" `
    (Test-Path -LiteralPath $shadowBuildPath -PathType Leaf) `
    "A dedicated entry point must build the Operations shadow outside deploy/current."
Add-Check "Shadow.CollisionHarness" `
    (Test-Path -LiteralPath $collisionToolPath -PathType Leaf) `
    "A deterministic harness must report component, public-procedure, and Ribbon callback collisions."
Add-Check "Shadow.CollisionResolutions" `
    (Test-Path -LiteralPath $resolutionPath -PathType Leaf) `
    "Every accepted shadow collision must have a reviewed machine-readable resolution."
Add-Check "Shadow.PackagedValidator" `
    (Test-Path -LiteralPath $validatorPath -PathType Leaf) `
    "Packaged validation must compile/load the shadow and initialize each role form in isolation."
Add-Check "Shadow.NotDeployed" `
    (-not (Test-Path -LiteralPath `
        (Join-Path $deployRoot "invSys.Operations.xlam") -PathType Leaf)) `
    "Slice 6 must not publish invSys.Operations.xlam to deploy/current."

$legacyRoleFiles = @(
    "invSys.Receiving.xlam",
    "invSys.Production.xlam",
    "invSys.Shipping.xlam"
)
$legacyPresent = @($legacyRoleFiles | Where-Object {
    -not (Test-Path -LiteralPath (Join-Path $deployRoot $_) -PathType Leaf)
}).Count -eq 0
Add-Check "Shadow.LegacyPackagesRemainActive" $legacyPresent `
    "The three standalone role XLAMs must remain the active deploy/current packages during Slice 6."

$passed = @($rows | Where-Object Passed).Count
$failed = $rows.Count - $passed
$lines = @(
    "# Slice 6 Shadow Operations Contract Results",
    "",
    "- Passed: $passed",
    "- Failed: $failed",
    "",
    "| Check | Result | Contract |",
    "|---|---|---|"
)
foreach ($row in $rows) {
    $result = if ($row.Passed) { "PASS" } else { "FAIL" }
    $lines += "| $($row.Name) | $result | $($row.Contract) |"
}
[IO.File]::WriteAllText(
    $resultPath,
    (($lines -join "`n") + "`n"),
    (New-Object Text.UTF8Encoding($false))
)

Write-Host "SLICE6_SHADOW_CONTRACT_RESULTS=$resultPath"
Write-Host "PASSED=$passed FAILED=$failed TOTAL=$($rows.Count)"
if ($failed -gt 0) {
    exit 1
}

exit 0
