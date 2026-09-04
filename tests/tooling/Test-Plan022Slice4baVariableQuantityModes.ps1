$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$form = Get-Content (Join-Path $root 'src\Production\Forms\frmProduction.frm') -Raw
$worksheet = Get-Content (Join-Path $root 'src\Production\Modules\modProductionProcessWorksheet.bas') -Raw
$production = Get-Content (Join-Path $root 'src\Production\Modules\mProduction.bas') -Raw
$schema = Get-Content (Join-Path $root 'src\DesignsDomain\Modules\modDesignsSchema.bas') -Raw
$apply = Get-Content (Join-Path $root 'src\DesignsDomain\Modules\modDesignsApply.bas') -Raw
$queries = Get-Content (Join-Path $root 'src\DesignsDomain\Modules\modDesignsQueries.bas') -Raw
$run = Get-Content (Join-Path $root 'src\Production\Modules\modProductionReusableRun.bas') -Raw

$checks = @(
    @{ Name = 'Process Designer modes'; Passed = $form -match 'mCmbRequirementQtyMode' -and $form -match 'mCmbProcessOutputQtyMode' -and $form -match 'Variable -- determined at Check In' -and $form -match 'Variable -- determined by Actual Output' },
    @{ Name = 'Versioned quantity modes'; Passed = $schema -match 'RequirementQtyMode' -and $schema -match 'OutputQtyMode' -and $apply -match 'RequirementQtyMode' -and $apply -match 'OutputQtyMode' -and $queries -match 'RequirementQtyMode' -and $queries -match 'OutputQtyMode' },
    @{ Name = 'Worksheet Qty Mode round trip'; Passed = $worksheet -match 'COL_QTY_MODE' -and $worksheet -match 'RequirementQtyMode' -and $worksheet -match 'OutputQtyMode' -and $worksheet -match 'Variable -- determined at Check In' },
    @{ Name = 'ACTUAL worksheet input skips formula-total validation'; Passed = $worksheet -match 'If qtyMode = "ACTUAL" And \(hasQty Or hasPercent Or basisQty > 0\) Then' -and $worksheet -match 'If qtyMode = "FIXED" Then inputUoms\(uom\) = True' -and $worksheet -match 'If qtyMode = "FIXED" Then\s*\r?\n\s*If Not inputPercentTotals.Exists\(uom\) Then' },
    @{ Name = 'Route-safe variable input'; Passed = $apply -match 'ACTUAL requirement cannot receive a Recipe connection' -and $run -match 'Actual requirement requires a measured external allocation' },
    @{ Name = 'Packaged public action test'; Passed = $production -match 'RunProductionVariableQuantityModeActionContractTest' -and $production -match 'BtnOpenProductionForm' -and $form -match 'TestProductionVariableQuantityModeActionContract' }
)

$failed = @($checks | Where-Object { -not $_.Passed })
$checks | ForEach-Object { "{0}: {1}" -f $_.Name, $(if ($_.Passed) { 'PASS' } else { 'FAIL' }) }
if ($failed.Count) { throw "Plan 022 Slice 4ba variable quantity mode static contract failed: $($failed.Name -join ', ')" }
