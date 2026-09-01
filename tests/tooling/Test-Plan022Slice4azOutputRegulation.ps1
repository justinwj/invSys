$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$form = Get-Content (Join-Path $root 'src\Production\Forms\frmProduction.frm') -Raw
$run = Get-Content (Join-Path $root 'src\Production\Modules\modProductionReusableRun.bas') -Raw
$production = Get-Content (Join-Path $root 'src\Production\Modules\mProduction.bas') -Raw
$schema = Get-Content (Join-Path $root 'src\DesignsDomain\Modules\modDesignsSchema.bas') -Raw
$apply = Get-Content (Join-Path $root 'src\DesignsDomain\Modules\modDesignsApply.bas') -Raw
$queries = Get-Content (Join-Path $root 'src\DesignsDomain\Modules\modDesignsQueries.bas') -Raw

$checks = @(
    @{ Name = 'Settings page'; Passed = $form -match 'Production Settings' -and $form -match 'mCmbOutputRegulationScope' -and $form -match 'mBtnOutputRegulationApply' },
    @{ Name = 'Settings workflow instructions'; Passed = $form -match 'How to use:' -and $form -match 'Production Run only reads the released Recipe version' -and $form -match 'selectedNodeId = ComboText\(mCmbOutputRegulationNode\)' },
    @{ Name = 'Versioned process defaults'; Passed = $form -match 'OutputRegulationEnabled' -and $schema -match 'OutputRegulationEnabled' -and $apply -match 'OutputRegulationEnabled' -and $queries -match 'OutputRegulationEnabled' },
    @{ Name = 'Versioned recipe overrides'; Passed = $form -match 'OUTPUT_REGULATION' -and $schema -match 'tblRecipeOutputRegulations' -and $apply -match 'OUTPUT_REGULATION' -and $queries -match 'OUTPUT_REGULATION' },
    @{ Name = 'Headless release validation'; Passed = $apply -match 'ValidateOutputRegulation' -and $apply -match 'ROUTED_COMMITMENT_EXCEEDS_CEILING' },
    @{ Name = 'Route-safe actual validation'; Passed = $run -match 'ValidateReusableActualOutput' -and $run -match 'EffectiveOutputRegulationFloor' -and $run -match 'ScaledOutputRegulationCeiling' -and $run -match 'below its routed downstream commitment' },
    @{ Name = 'Packaged public action test'; Passed = $production -match 'RunProductionOutputRegulationActionContractTest' -and $production -match 'BtnOpenProductionForm' -and $form -match 'TestProductionOutputRegulationActionContract' }
)

$failed = @($checks | Where-Object { -not $_.Passed })
$checks | ForEach-Object { "{0}: {1}" -f $_.Name, $(if ($_.Passed) { 'PASS' } else { 'FAIL' }) }
if ($failed.Count) { throw "Plan 022 Slice 4az output regulation static contract failed: $($failed.Name -join ', ')" }
