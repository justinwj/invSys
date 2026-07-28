[CmdletBinding()]
param(
    [string]$RepoRoot = "."
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$modulePath = Join-Path $repo "src/Production/Modules/mProduction.bas"
$eventCreatorPath = Join-Path $repo "src/Production/Modules/modProductionEventCreator.bas"
$formPath = Join-Path $repo "src/Production/Forms/frmProduction.frm"
$buildScriptPath = Join-Path $repo "tools/build-xlam.ps1"
$resultPath = Join-Path $repo "tests/unit/slice8_production_retirement_results.md"

$moduleText = Get-Content -LiteralPath $modulePath -Raw
$eventCreatorText = Get-Content -LiteralPath $eventCreatorPath -Raw
$formText = Get-Content -LiteralPath $formPath -Raw
$buildScriptText = Get-Content -LiteralPath $buildScriptPath -Raw
$identitySurface = $moduleText + "`n" + $eventCreatorText + "`n" + $formText
$controllerSurface = $moduleText + "`n" + $formText
$productionSourceText = (
    Get-ChildItem -LiteralPath (Join-Path $repo "src/Production") -Recurse -File |
        Where-Object { $_.Extension -in @(".bas", ".cls", ".frm") } |
        Sort-Object FullName |
        ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw }
) -join "`n"

$rows = [System.Collections.Generic.List[object]]::new()
function Add-Check {
    param([string]$Name, [bool]$Passed, [string]$Detail)
    $rows.Add([pscustomobject]@{
        Name = $Name
        Passed = $Passed
        Detail = $Detail
    }) | Out-Null
}

Add-Check "Production.Identity.NoManagedRowLiteral" `
    ($identitySurface -notmatch '(?i)"ROW"') `
    "Production runtime code must not declare, resolve, serialize, display, or restore the retired ROW identity."
Add-Check "Production.Identity.NoRowAliases" `
    ($identitySurface -notmatch '(?i)"ROWID"|"ROW#"') `
    "Production runtime code must not retain compatibility aliases for retired ROW identity."
$legacyRowAuthorityNames = @(
    "ResolveInvSysDetailsByRow",
    "BuildInvSysOutputLookup",
    "ResolveProductionOutputInventoryRow",
    "ResolveProductionOutputIdentityFromPicker",
    "LookupOutputRowLooseFromPicker",
    "LookupOutputRowLoose",
    "InventoryRowMatchesOutputName",
    "BuildInvSysRowIndex",
    "EnrichOutputDeltaFromPicker",
    "ResolveInvSysLocationByRow"
)
$legacyRowAuthorityPattern = "(?i)\b(" + (($legacyRowAuthorityNames | ForEach-Object {
    [regex]::Escape($_)
}) -join "|") + ")\b"
Add-Check "Production.Identity.NoNumericRowAuthorityHelpers" `
    ($moduleText -notmatch $legacyRowAuthorityPattern) `
    "Production must not preserve numeric-row lookup, resolution, picker, or index helpers under renamed System_Key headers."
$legacyIdentityNames = @(
    "NormalizeRowKey",
    "NormalizeRunRowKey",
    "BuildInvSysRowMap",
    "AddInvSysTableRowsToRowMap",
    "AddInventoryPickerRowsToRowMap",
    "BuildRowKeySetFromDeltas",
    "BuildUsedSnapshotForRows",
    "GetAllowedInvRowsForIngredient"
)
$legacyIdentityNamePattern = "(?i)\b(" + (($legacyIdentityNames | ForEach-Object {
    [regex]::Escape($_)
}) -join "|") + ")\b"
Add-Check "Production.Identity.NoLegacyRowIdentityNames" `
    ($identitySurface -notmatch $legacyIdentityNamePattern) `
    "Production identity helpers and form lookups must name and preserve System_Key rather than normalize legacy numeric row keys."
Add-Check "Production.Identity.SystemKeySurface" `
    (($moduleText -match '(?i)"System_Key"') -and
     ($eventCreatorText -match '(?i)"System_Key"') -and
     ($formText -match '(?i)"System_Key"')) `
    "The controller, event creator, and form must all carry immutable System_Key identity."

$sameProjectRun = '(?i)Application\.Run[^\r\n]*"mProduction\.'
Add-Check "Production.InternalCalls.NoSameProjectApplicationRun" `
    ($formText -notmatch $sameProjectRun) `
    "Production form-to-controller calls inside Operations must be direct typed procedure calls."
Add-Check "Production.InternalCalls.NoControllerApplicationRun" `
    ($controllerSurface -notmatch '(?i)\bApplication\.Run\b') `
    "Production controller and form must use typed calls; dynamic dispatch belongs only in declared cross-XLAM bridge modules."
Add-Check "Production.Bridges.PrimitiveJsonOnly" `
    (($productionSourceText -notmatch '(?i)modRoleEventWriter\.BuildPayloadJsonFromCollection') -and
     ($productionSourceText -notmatch '(?i)\bWarehouseTarget\b') -and
     ($productionSourceText -notmatch '(?i)\bmodNasConnection\.GetCurrentTarget\b') -and
     ($productionSourceText -notmatch '(?i)\bmodRoleEventWriter\.Create(?:InventoryEntity)?PayloadItem\b') -and
     ($productionSourceText -notmatch '(?i)\bmodRoleWorkbookSurfaces\.') -and
     ($productionSourceText -notmatch '(?i)\bmodOperatorReadModel\.') -and
     ($productionSourceText -notmatch '(?i)\bmodUiQuiet\.BeginQuietUi\b') -and
     ($productionSourceText -notmatch '(?i)\bmodRoleUiAccess\.ApplyShapeCapability\b') -and
     ($productionSourceText -notmatch '(?i)\bmodDesignsDomainBridge\.(?:ListDesignsBridge|GetDesignBOMBridge|GetDesignBOMForStatusBridge)\b') -and
     ($productionSourceText -match '(?i)\bmodOperationsPrimitiveBridge\.') -and
     ($controllerSurface -notmatch '(?i)\bmodUR_Snapshot\.') -and
     ($controllerSurface -notmatch '(?i)\bmodUserFormResizeWin\.')) `
    "Production must serialize payloads and create payload objects locally, consume primitive target/workbook/shape values through the declared bridge, and never pass Collections, forms, workbooks, worksheets, or Core class instances across the Core XLAM boundary."
Add-Check "Production.InternalCalls.NoDynamicControllerWrappers" `
    ($formText -notmatch '(?i)\bRunProduction(Object)?[012]\b') `
    "Dynamic RunProduction wrapper dispatch must be retired."
Add-Check "Production.Form.ModelessLauncher" `
    ($moduleText -match '(?i)frmProduction\.Show\s+vbModeless') `
    "The Production form must open modelessly while retaining its captured workbook."
Add-Check "Production.Form.CapturedContextAuthority" `
    (($formText -notmatch '(?i)\bApplication\.ActiveWorkbook\b') -and
     ($formText -notmatch '(?i)\bActivateOperatorWorkbookForRun\b') -and
     ($formText -notmatch '(?i)\bwb\.Activate\b') -and
     ($formText -match '(?i)\bmProduction\.BindProductionOperatorWorkbook\b') -and
     ($moduleText -match '(?i)\bPrivate\s+mProductionOperatorWorkbook\s+As\s+Workbook\b')) `
    "A modeless Production form must route through its captured workbook binding without activating or recapturing Application.ActiveWorkbook."

$legacyMutationNames = @(
    "ApplyUsedDeltasLocal",
    "ApplyMadeDeltasLocal",
    "ApplyMadeToInventoryDeltasLocal"
)
$legacyMutationPattern = "(?i)\b(" + (($legacyMutationNames | ForEach-Object {
    [regex]::Escape($_)
}) -join "|") + ")\b"
Add-Check "Production.Domain.NoLegacyLocalInventoryMutation" `
    ($moduleText -notmatch $legacyMutationPattern) `
    "Legacy local inventory mutation procedures must be removed after completion-service cutover."

$designsBranchPattern = '(?s)Public Function LoadIngredientListForRecipe.*?' +
                        'If ProductionDesignsEnabled\(\) Then.*?' +
                        'LoadIngredientListFromDesigns.*?' +
                        'Exit Function'
Add-Check "Production.Designs.ReleasedOnlyWhenEnabled" `
    ($moduleText -match $designsBranchPattern) `
    "Designs-enabled Production must return released Designs Domain ingredients without falling through to legacy recipes."

$testOnlyRegionPattern = '(?ms)^\s*''@TestOnlyBegin\s*$.*?^\s*''@TestOnlyEnd\s*$'
$deployedModuleText = [regex]::Replace($moduleText, $testOnlyRegionPattern, "")
Add-Check "Production.Runtime.NoEmbeddedTestFixtures" `
    (($deployedModuleText -notmatch '(?im)^\s*Public\s+(Function|Sub)\s+Test') -and
     ($buildScriptText -match '(?i)Remove-VbaTestOnlyRegions') -and
     ($buildScriptText -match '@TestOnlyBegin')) `
    "Production controller test fixtures must be explicitly marked and stripped from the deployed runtime module while action adapters remain available to packaged form tests."

$passed = @($rows | Where-Object Passed).Count
$failed = $rows.Count - $passed
$lines = @(
    "# Slice 8 Production Retirement Results",
    "",
    "- Passed: $passed",
    "- Failed: $failed",
    "",
    "| Check | Result | Contract |",
    "|---|---|---|"
)
foreach ($row in $rows) {
    $state = if ($row.Passed) { "PASS" } else { "FAIL" }
    $detail = ([string]$row.Detail).Replace("|", "/")
    $lines += "| $($row.Name) | $state | $detail |"
}
[IO.File]::WriteAllText(
    $resultPath,
    (($lines -join "`n") + "`n"),
    [Text.UTF8Encoding]::new($false)
)

Write-Host "RESULT passed=$passed failed=$failed"
if ($failed -gt 0) {
    foreach ($row in $rows | Where-Object { -not $_.Passed }) {
        Write-Host ("  " + $row.Name + ": " + $row.Detail)
    }
    exit 1
}
