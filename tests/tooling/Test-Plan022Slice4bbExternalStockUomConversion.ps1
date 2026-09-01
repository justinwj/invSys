$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$uom = Get-Content (Join-Path $root 'src\Core\Modules\modUomSettings.bas') -Raw
$configDefaults = Get-Content (Join-Path $root 'src\Core\Modules\modConfigDefaults.bas') -Raw
$config = Get-Content (Join-Path $root 'src\Core\Modules\modConfig.bas') -Raw
$worksheet = Get-Content (Join-Path $root 'src\Production\Modules\modProductionUomCatalogWorksheet.bas') -Raw -ErrorAction SilentlyContinue
$form = Get-Content (Join-Path $root 'src\Production\Forms\frmProduction.frm') -Raw
$run = Get-Content (Join-Path $root 'src\Production\Modules\modProductionReusableRun.bas') -Raw
$production = Get-Content (Join-Path $root 'src\Production\Modules\mProduction.bas') -Raw
$schemaCapacity = [regex]::Match($configDefaults, 'ReDim defs\(1 To (\d+)\)')
$schemaKeys = [regex]::Matches($configDefaults, 'AddConfigKey defs, idx,').Count
$warehouseHeaders = [regex]::Match($config, '(?s)whHeaders\s*=\s*Array\(\s*_(.*?)\)\s*stHeaders')

$checks = @(
    @{ Name = 'Versioned dimension catalog'; Passed = $uom -match 'GetUomConversion' -and $uom -match 'Units Per Base UOM' -and $uom -match 'UOM_CATALOG' },
    @{ Name = 'Config schema capacity'; Passed = $schemaCapacity.Success -and ([int]$schemaCapacity.Groups[1].Value -ge $schemaKeys) },
    @{ Name = 'Persisted config schema'; Passed = $warehouseHeaders.Success -and $warehouseHeaders.Groups[1].Value -match '"UomCatalog"' -and $warehouseHeaders.Groups[1].Value -match '"UomConversionCatalog"' -and $warehouseHeaders.Groups[1].Value -match '"UomConversionCatalogVersion"' },
    @{ Name = 'Captured worksheet workbench'; Passed = $worksheet -match 'Attribute VB_Name = "modProductionUomCatalog"' -and $worksheet -match 'SendUomCatalogToWorksheet' -and $worksheet -match 'RetrieveUomCatalogFromWorksheet' -and $worksheet -match 'Units Per Base UOM' },
    @{ Name = 'Production Settings actions'; Passed = $form -match 'Edit UOM Catalog on Sheet' -and $form -match 'Retrieve UOM Catalog' },
    @{ Name = 'Native exact-key allocation'; Passed = $run -match 'NativeAllocationQty' -and $run -match 'GetUomConversion' -and $run -match 'ConversionCatalogVersion' -and $run -match '\(availableQty - reservedOther\) \* conversionFactor' -and $run -match 'RequirementUOM & " / " & stockUom' },
    @{ Name = 'Packaged public action'; Passed = $form -match 'mBtnUomCatalogSend_Click' -and $form -match 'mBtnUomCatalogRetrieve_Click' -and $form -match 'SheetAction=' -and $production -match 'RunProductionExternalStockUomWorksheetHandlerContractTest' -and $production -match 'BtnOpenProductionForm' }
)

$failed = @($checks | Where-Object { -not $_.Passed })
$checks | ForEach-Object { "{0}: {1}" -f $_.Name, $(if ($_.Passed) { 'PASS' } else { 'FAIL' }) }
if ($failed.Count) { throw "Plan 022 Slice 4bb external-stock UOM conversion contract failed: $($failed.Name -join ', ')" }
