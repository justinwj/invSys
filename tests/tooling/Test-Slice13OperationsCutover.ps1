[CmdletBinding()]
param(
    [string]$RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
)

$ErrorActionPreference = "Stop"
$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$buildPath = Join-Path $repo "tools/build-xlam.ps1"
$registerPath = Join-Path $repo "tools/register_current_addins.ps1"
$deployPath = Join-Path $repo "tools/deploy_current_xlams_to_nas.ps1"
$packagedPath = Join-Path $repo "tools/validate_phase6_packaged_xlams.ps1"
$ribbonPath = Join-Path $repo "tools/validate_phase6_packaged_ribbon.ps1"
$livePath = Join-Path $repo "tools/validate_phase6_live_role_workflows.ps1"
$surfacePath = Join-Path $repo "src/Core/Modules/modRoleWorkbookSurfaces.bas"
$publishPath = Join-Path $repo "src/Admin/Modules/modAddinsPublish.bas"
$registrationModulePath = Join-Path $repo "src/Admin/Modules/modLocalAddinsRegistration.bas"
$diagnosticsPath = Join-Path $repo "src/Admin/Modules/modPackageDiagnostics.bas"
$testerBundlePath = Join-Path $repo "src/Admin/Modules/modTesterBundle.bas"
$testerSetupPath = Join-Path $repo "src/Admin/Modules/modTesterSetup.bas"
$operationsInitPath = Join-Path $repo "src/Operations/Modules/modOperationsInit.bas"
$deployRoot = Join-Path $repo "deploy/current"
$packageManifestPath = Join-Path $deployRoot "addins-manifest.json"

$results = [System.Collections.Generic.List[object]]::new()
function Add-Check {
    param([string]$Name, [bool]$Passed, [string]$Detail)
    $results.Add([pscustomobject]@{
        Check = $Name
        Passed = $Passed
        Detail = $Detail
    }) | Out-Null
}

function Read-Text([string]$Path) {
    return Get-Content -Raw -LiteralPath $Path
}

$buildText = Read-Text $buildPath
$registerText = Read-Text $registerPath
$deployText = Read-Text $deployPath
$packagedText = Read-Text $packagedPath
$ribbonText = Read-Text $ribbonPath
$liveText = Read-Text $livePath
$surfaceText = Read-Text $surfacePath
$operationsInitText = Read-Text $operationsInitPath
$adminTexts = @(
    (Read-Text $publishPath),
    (Read-Text $registrationModulePath),
    (Read-Text $diagnosticsPath),
    (Read-Text $testerBundlePath),
    (Read-Text $testerSetupPath)
) -join "`n"

$operationsBlock = [regex]::Match(
    $buildText,
    '(?is)Key\s*=\s*"Operations".*?(?=\r?\n\s*}\s*\r?\n\s*@\{\s*\r?\n\s*Key\s*=)'
).Value
Add-Check "Slice13.Build.DeployableOperationsProject" `
    (($operationsBlock -match 'OutputFile\s*=\s*"invSys\.Operations\.xlam"') -and
     ($operationsBlock -match 'src/Receiving') -and
     ($operationsBlock -match 'src/Production') -and
     ($operationsBlock -match 'src/Shipping') -and
     ($operationsBlock -notmatch 'Deployable\s*=\s*\$false')) `
    "The selectable Operations target must build the complete combined role project as a deployable XLAM."

Add-Check "Slice13.Build.NoStandaloneRoleProjects" `
    (($buildText -notmatch 'Key\s*=\s*"Receiving"') -and
     ($buildText -notmatch 'Key\s*=\s*"Production"') -and
     ($buildText -notmatch 'Key\s*=\s*"Shipping"')) `
    "Standalone role XLAMs must no longer be build targets."

Add-Check "Slice13.Build.LegacyRetirement" `
    (($operationsBlock -match 'LegacyOutputFiles\s*=\s*@\([^\)]*invSys\.Receiving\.xlam') -and
     ($operationsBlock -match 'LegacyOutputFiles\s*=\s*@\([^\)]*invSys\.Production\.xlam') -and
     ($operationsBlock -match 'LegacyOutputFiles\s*=\s*@\([^\)]*invSys\.Shipping\.xlam')) `
    "The Operations build must archive all three superseded role binaries."

Add-Check "Slice13.Ribbon.OneOperationsTab" `
    (($operationsBlock -match 'TabId\s*=\s*"tabInvSysOperations"') -and
     ($operationsBlock -match 'Label\s*=\s*"Operations"') -and
     ($operationsBlock -match 'Id\s*=\s*"grpOperationsSession"') -and
     ($operationsBlock -match 'Id\s*=\s*"grpOperationsReceiving"') -and
     ($operationsBlock -match 'Id\s*=\s*"grpOperationsProduction"') -and
     ($operationsBlock -match 'Id\s*=\s*"grpOperationsShipping"')) `
    "One Operations tab must contain shared session controls and independently gated role groups."

Add-Check "Slice13.Ribbon.RoleLaunchersOnly" `
    (($operationsBlock -match 'btnOperationsReceivingForm') -and
     ($operationsBlock -match 'btnOperationsProductionForm') -and
     ($operationsBlock -match 'btnOperationsShippingForm') -and
     ($operationsBlock -notmatch 'btnShippingBoxBuilderForm') -and
     ($operationsBlock -notmatch 'btnShippingBoxMakerForm') -and
     ($operationsBlock -notmatch 'Purchasing.*Button')) `
    "The ribbon must expose one launcher per main role and no separate Purchasing, Box Builder, or Box Maker launcher."

Add-Check "Slice13.Tooling.FivePackageLists" `
    (($registerText -match 'invSys\.Operations\.xlam') -and
     ($deployText -match 'invSys\.Operations\.xlam') -and
     ($packagedText -match 'invSys\.Operations\.xlam') -and
     ($ribbonText -match 'invSys\.Operations\.xlam') -and
     ($liveText -match 'invSys\.Operations\.xlam')) `
    "Registration, publication, and packaged validators must target the consolidated package."

$toolingLegacyRoleRefs = @(
    $registerText, $deployText, $packagedText, $ribbonText, $liveText
) -join "`n"
Add-Check "Slice13.Tooling.NoLegacyRolePackageRefs" `
    (($toolingLegacyRoleRefs -notmatch 'invSys\.Receiving\.xlam') -and
     ($toolingLegacyRoleRefs -notmatch 'invSys\.Production\.xlam') -and
     ($toolingLegacyRoleRefs -notmatch 'invSys\.Shipping\.xlam')) `
    "Active registration, publication, and packaged validation paths must not require standalone role binaries."

Add-Check "Slice13.Core.WorkbookButtonsRouteToOperations" `
    (($surfaceText -match "'invSys\.Operations\.xlam'!modTS_Received\.ConfirmWrites") -and
     ($surfaceText -notmatch "'invSys\.Receiving\.xlam'!")) `
    "Generated operator-workbook buttons must route to Operations."

Add-Check "Slice13.Admin.FivePackageContract" `
    (($adminTexts -match 'invSys\.Operations\.xlam') -and
     ($adminTexts -notmatch 'invSys\.Receiving\.xlam') -and
     ($adminTexts -notmatch 'invSys\.Production\.xlam') -and
     ($adminTexts -notmatch 'invSys\.Shipping\.xlam')) `
    "Admin publish, registration, diagnostics, bundle, and tester setup paths must use the five-package contract."

$expectedFiles = @(
    "invSys.Core.xlam",
    "invSys.Inventory.Domain.xlam",
    "invSys.Designs.Domain.xlam",
    "invSys.Operations.xlam",
    "invSys.Admin.xlam"
)
$actualXlams = @(
    Get-ChildItem -LiteralPath $deployRoot -Filter "*.xlam" -File |
        Select-Object -ExpandProperty Name |
        Sort-Object
)
Add-Check "Slice13.Deploy.ExactFivePackages" `
    (($actualXlams.Count -eq 5) -and
     (@(Compare-Object ($expectedFiles | Sort-Object) $actualXlams).Count -eq 0)) `
    "deploy/current must contain exactly the five normative XLAM filenames."

Add-Check "Slice13.Deploy.LegacyAbsent" `
    (-not (Test-Path (Join-Path $deployRoot "invSys.Receiving.xlam")) -and
     -not (Test-Path (Join-Path $deployRoot "invSys.Production.xlam")) -and
     -not (Test-Path (Join-Path $deployRoot "invSys.Shipping.xlam"))) `
    "Standalone role binaries must be absent after cutover."

$packageManifestGreen = $false
if (Test-Path -LiteralPath $packageManifestPath -PathType Leaf) {
    $packageManifest = Get-Content -Raw -LiteralPath $packageManifestPath | ConvertFrom-Json
    $manifestNames = @($packageManifest.packages | ForEach-Object { $_.name } | Sort-Object)
    $hashesMatch = $true
    foreach ($package in @($packageManifest.packages)) {
        $packagePath = Join-Path $deployRoot ([string]$package.name)
        if (-not (Test-Path -LiteralPath $packagePath -PathType Leaf) -or
            (Get-FileHash -LiteralPath $packagePath -Algorithm SHA256).Hash.ToLowerInvariant() -ne
                ([string]$package.sha256).ToLowerInvariant()) {
            $hashesMatch = $false
        }
    }
    $packageManifestGreen = (
        $packageManifest.packageSetVersion -eq "R1-5" -and
        $manifestNames.Count -eq 5 -and
        @(Compare-Object ($expectedFiles | Sort-Object) $manifestNames).Count -eq 0 -and
        $hashesMatch
    )
}
Add-Check "Slice13.Deploy.VersionCoherentManifest" $packageManifestGreen `
    "The deployed five-package set must have a hash-verified R1-5 manifest."

Add-Check "Slice13.Build.SameProjectRibbonCallsTyped" `
    ($buildText -notmatch 'Application\.Run\s+""''""\s*&\s*ThisWorkbook\.Name') `
    "Generated same-project Ribbon callbacks must use direct typed procedure calls."

Add-Check "Slice13.Upgrade.CoexistenceDiagnostic" `
    (($operationsInitText -match 'LegacyRoleAddinCoexistenceReport') -and
     ($operationsInitText -match 'invSys\.Receiving\.xlam') -and
     ($operationsInitText -match 'invSys\.Production\.xlam') -and
     ($operationsInitText -match 'invSys\.Shipping\.xlam') -and
     ($operationsInitText -match 'register_current_addins\.ps1')) `
    "Operations startup must detect stale role add-ins and provide an exact remediation path."

$results | Format-Table -AutoSize
$failed = @($results | Where-Object { -not $_.Passed })
Write-Host ("Slice 13 Operations cutover contract: {0} passed, {1} failed" -f
    ($results.Count - $failed.Count), $failed.Count)
if ($failed.Count -gt 0) { exit 1 }
