[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
$repo = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path

function Read-Source([string]$RelativePath) {
    Get-Content -Raw -LiteralPath (Join-Path $repo $RelativePath)
}

function Add-Check([string]$Name, [bool]$Passed, [string]$Contract) {
    $script:checks.Add([pscustomobject]@{
        Name = $Name
        Passed = $Passed
        Contract = $Contract
    }) | Out-Null
}

$productionModule = Read-Source "src\Production\Modules\mProduction.bas"
$productionForm = Read-Source "src\Production\Forms\frmProduction.frm"
$designsRuntime = Read-Source "src\DesignsDomain\Modules\modDesignsRuntime.bas"
$designsSchema = Read-Source "src\DesignsDomain\Modules\modDesignsSchema.bas"
$validator = Read-Source "tools\validate_plan022_packaged_launchers.ps1"
$checks = New-Object 'System.Collections.Generic.List[object]'

Add-Check "Production.Restart.PublicProbe" `
    (($productionModule -match 'Public Function RunReusableProductionRestartActionContractTest') -and
     ($productionModule -match 'TestReusableProductionRestartActionContract')) `
    "Packaged Operations exposes a bounded restart probe through the already-open public Production form."

Add-Check "Production.Restart.UsesRunListLoadHandler" `
    (($productionForm -match 'Public Function TestReusableProductionRestartActionContract') -and
     ($productionForm -match '(?s)TestReusableProductionRestartActionContract.*?FindIdentityListRow\(mLstLoaderRecipes') -and
     ($productionForm -match '(?s)TestReusableProductionRestartActionContract.*?mBtnLoaderLoad_Click') -and
     ($productionForm -match '(?s)TestReusableProductionRestartActionContract.*?ReusableRunIsLoaded')) `
    "After restart, the probe selects the persisted exact Recipe version and invokes the operator Run List Load handler."

Add-Check "Production.Restart.TwoExcelProcesses" `
    (($validator -match '(?s)ProductionReusable.*?reusableRecipeId.*?Quit\(\).*?New-Object -ComObject Excel\.Application') -and
     ($validator -match 'RunReusableProductionRestartActionContractTest')) `
    "The ProductionReusable validator crosses a real Excel process boundary before exercising persisted design state."

Add-Check "Production.Restart.ReusesSavedWorkbook" `
    (($validator -match 'PRODUCTION_RESTART') -and
     ($validator -match 'RestartSameOperatorWorkbook') -and
     ($validator -match 'RestartNewWorkbooks')) `
    "The second session proves the public launcher reuses the same saved station-local Production workbook without creating another one."

Add-Check "Production.Launch.ReadsDoNotSaveDesignsAuthority" `
    (($designsRuntime -match 'If Not wb\.ReadOnly And Not wb\.Saved Then wb\.Save') -and
     ($designsRuntime -notmatch 'If Not wb\.ReadOnly Then wb\.Save')) `
    "Resolving an already-valid Designs workbook for Production lists must not save or mutate canonical authority merely from form launch."

Add-Check "Production.Launch.SchemaFormattingIsIdempotent" `
    (($designsSchema -match 'currentFormat\s*=\s*lc\.Range\.NumberFormat') -and
     ($designsSchema -match 'If IsNull\(currentFormat\) Then') -and
     ($designsSchema -match 'ElseIf StrComp\(CStr\(currentFormat\), "@", vbBinaryCompare\) <> 0 Then')) `
    "Repeated schema assurance must not dirty the Designs workbook by reassigning an already-correct text number format."

$passed = @($checks | Where-Object Passed).Count
$failed = $checks.Count - $passed
foreach ($check in $checks) {
    $result = if ($check.Passed) { "PASS" } else { "FAIL" }
    Write-Host "$result $($check.Name) - $($check.Contract)"
}
Write-Host "RESULT passed=$passed failed=$failed"
if ($failed -gt 0) { exit 1 }
