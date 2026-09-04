[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

$repo = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path

function Read-Source([string]$RelativePath) {
    Get-Content -Raw -LiteralPath (Join-Path $repo $RelativePath)
}

$productionModule = Read-Source "src\Production\Modules\mProduction.bas"
$productionForm = Read-Source "src\Production\Forms\frmProduction.frm"
$reusableDesigns = Read-Source "src\Production\Modules\modProductionReusableDesigns.bas"
$reusableRun = Read-Source "src\Production\Modules\modProductionReusableRun.bas"
$designsApply = Read-Source "src\DesignsDomain\Modules\modDesignsApply.bas"
$designsSchema = Read-Source "src\DesignsDomain\Modules\modDesignsSchema.bas"
$viewerData = Read-Source "src\Core\Modules\modInventoryViewerData.bas"

$pageCaptions = @(
    '"Process Designer"',
    '"Recipe Designer"',
    '"Ingredients Assignment"',
    '"Production Run - List"',
    '"Production Run - Tree"'
)
$hasFiveTargetPages = (@($pageCaptions | Where-Object {
    -not $productionForm.Contains($_)
}).Count -eq 0) -and ($productionForm -notmatch 'Caption\s*=\s*"Recipe Builder"')

$checks = @(
    [pscustomobject]@{
        Name = "Production.PublicLauncherPreserved"
        Passed = ($productionModule -match 'Public Sub BtnOpenProductionForm\(\)') -and
            ($productionModule -match 'RequireCurrentUserCapabilityCached\("PROD_POST"\)') -and
            ($productionModule -match 'OpenOrCreateCurrentRoleOperatorWorkbook') -and
            ($productionModule -match 'ShowProductionForm wb')
        Contract = "Slice 4x preserves the packaged Production launcher, PROD_POST gate, and captured-workbook form boundary."
    },
    [pscustomobject]@{
        Name = "Production.Form.FiveTargetPages"
        Passed = $hasFiveTargetPages
        Contract = "The public Production form exposes Process Designer, Recipe Designer, Ingredients Assignment, Run List, and experimental Run Tree, with no Recipe Builder page."
    },
    [pscustomobject]@{
        Name = "Production.ProcessDesigner.LifecycleHandlers"
        Passed = ($productionForm -match 'New Process') -and
            ($productionForm -match 'Reuse as New Version') -and
            ($productionForm -match 'Save Draft') -and
            ($productionForm -match 'Obsolete') -and
            ($reusableDesigns -match 'PROCESS_SAVE') -and
            ($reusableDesigns -match 'PROCESS_RELEASE') -and
            ($reusableDesigns -match 'PROCESS_OBSOLETE')
        Contract = "Process Designer uses operator handlers for new/reuse, draft save, release, and obsolete lifecycle events."
    },
    [pscustomobject]@{
        Name = "Production.ProcessDesigner.RequiresOutput"
        Passed = ($productionForm -match 'Process.*at least one output') -and
            ($designsApply -match 'Process.*at least one output')
        Contract = "The form and Designs Domain both reject a Process definition with no output."
    },
    [pscustomobject]@{
        Name = "Production.RecipeDesigner.GraphHandlers"
        Passed = ($productionForm -match 'Validate Recipe') -and
            ($productionForm -match 'Auto Order') -and
            ($productionForm -match 'Disconnect') -and
            ($productionForm -match 'ValidateRecipeDraft') -and
            ($designsApply -match 'circular|cycle') -and
            ($reusableRun -match 'unresolved') -and
            ($designsApply -match 'over-allocated|over-allocation|overalloc')
        Contract = "Recipe Designer and Designs Domain validate connections, execution order, unresolved inputs, quantities, and circular dependencies."
    },
    [pscustomobject]@{
        Name = "Production.IngredientsAssignment.ProcessRequirements"
        Passed = ($productionForm -match 'Ingredient Requirements') -and
            ($productionForm -match 'Acceptable Items') -and
            ($productionForm -match 'ProcessVersion') -and
            ($reusableDesigns -match 'PROCESS_SAVE')
        Contract = "Ingredients Assignment maps each exact Process-version requirement to acceptable managed item/SKU alternatives."
    },
    [pscustomobject]@{
        Name = "Production.RunSession.MultiOutput"
        Passed = ($reusableRun -match 'Private mOutputs As Collection') -and
            ($reusableRun -match 'Private mCompletedNodes As Object') -and
            ($reusableRun -match 'AssignFreshOutputKeysForNode') -and
            ($reusableRun -notmatch 'Private mOutputSystemKey As String')
        Contract = "The typed run session carries correlated Process executions and multiple output allocations rather than one singular output key."
    },
    [pscustomobject]@{
        Name = "Production.Completion.ExactKeysAndCoProducts"
        Passed = ($reusableRun -match 'BuildNodeCompleteItems') -and
            ($reusableRun -match 'BuildNodeConsumeItems') -and
            ($reusableRun -match 'OutputIdentityKey') -and
            ($reusableRun -match 'OutgoingQtyForOutput') -and
            ($reusableRun -match 'AllocationTotalForEntity')
        Contract = "Completion serializes every fresh output key, exact routed-intermediate input keys, and finished/co-product balances."
    },
    [pscustomobject]@{
        Name = "DesignsDomain.ProcessRecipeSchemaAndEvents"
        Passed = ($designsSchema -match 'tblProcesses') -and
            ($designsSchema -match 'tblProcessRequirements') -and
            ($designsSchema -match 'tblProcessOutputs') -and
            ($designsSchema -match 'tblProcessIngredientAlternatives') -and
            ($designsSchema -match 'tblRecipes') -and
            ($designsSchema -match 'tblRecipeProcesses') -and
            ($designsSchema -match 'tblRecipeConnections') -and
            ($designsApply -match 'PROCESS_SAVE') -and
            ($designsApply -match 'RECIPE_SAVE')
        Contract = "The headless Designs Domain owns reusable Process/Recipe lifecycle events and rebuildable projections."
    },
    [pscustomobject]@{
        Name = "Viewer.ProductionOperatorEvents"
        Passed = ($viewerData -match 'Production Input Consumed') -and
            ($viewerData -match 'Production Output Created')
        Contract = "Published Viewer Events expose correlated Production input consumption and output creation as operator actions."
    }
)

$passed = @($checks | Where-Object Passed).Count
$failed = $checks.Count - $passed
$resultPath = Join-Path $repo "tests\integration\plan022_slice4x_reusable_production_red_results.md"
$lines = @(
    "# Plan 022 Slice 4x Reusable Production Results", "",
    "- Passed: $passed", "- Failed: $failed", "",
    "| Check | Result | Contract |", "|---|---|---|"
)
foreach ($check in $checks) {
    $result = if ($check.Passed) { "PASS" } else { "FAIL" }
    $lines += "| $($check.Name) | $result | $($check.Contract) |"
}
[IO.File]::WriteAllText(
    $resultPath,
    (($lines -join "`n") + "`n"),
    (New-Object Text.UTF8Encoding($false))
)

Write-Host "PLAN022_SLICE4X_RESULTS=$resultPath"
Write-Host "PASSED=$passed FAILED=$failed TOTAL=$($checks.Count)"
if ($failed -gt 0) { exit 1 }
