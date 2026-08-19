[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

$repo = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path

function Read-Source([string]$relativePath) {
    Get-Content -Raw -LiteralPath (Join-Path $repo $relativePath)
}

$bootstrap = Read-Source "src\Core\Modules\modWarehouseBootstrap.bas"
$stationIdentityPath = Join-Path $repo "src\Core\Modules\modStationIdentity.bas"
$stationIdentity = if (Test-Path -LiteralPath $stationIdentityPath) {
    Get-Content -Raw -LiteralPath $stationIdentityPath
} else { "" }
$seedForm = Read-Source "src\Admin\Forms\frmSeedInventory.frm"
$connectionForm = Read-Source "src\Core\Forms\frmWarehouseConnection.frm"
$seedService = Read-Source "src\Admin\Modules\modAdminInventorySeed.bas"
$adminRibbon = Read-Source "tools\build-xlam.ps1"
$adminForm = Read-Source "src\Admin\Forms\frmAddInventoryItem.frm"
$shippingForm = Read-Source "src\Shipping\Forms\frmShipmentsTally.frm"
$choiceForm = Read-Source "src\Shipping\Forms\frmBoxVersionSaveChoice.frm"
$shippingModule = Read-Source "src\Shipping\Modules\modTS_Shipments.bas"
$boxingService = Read-Source "src\Shipping\Modules\modBoxingService.bas"
$receivingForm = Read-Source "src\Receiving\Forms\frmReceiving.frm"
$receivingService = Read-Source "src\Receiving\Modules\modTS_Received.bas"
$receivingPosting = Read-Source "src\Receiving\Modules\modReceivingPostingService.bas"
$receivingSurface = Read-Source "src\Core\Modules\modRoleWorkbookSurfaces.bas"
$productionForm = Read-Source "src\Production\Forms\frmProduction.frm"
$productionModule = Read-Source "src\Production\Modules\mProduction.bas"
$receivingEnsureBlock = [regex]::Match(
    $receivingSurface,
    '(?ms)^Public Function EnsureReceivingWorkbookSurface.*?^End Function'
).Value
$receivingArrangeBlock = [regex]::Match(
    $receivingSurface,
    '(?ms)^Private Sub ArrangeReceivingTablesSurface.*?^End Sub'
).Value
$receivingRebuildBlock = [regex]::Match(
    $receivingSurface,
    '(?ms)^Private Sub RebuildTableAtSurface.*?^End Sub'
).Value

$shippingPickerBlock = @(
    [regex]::Match($shippingModule, '(?ms)^Private Function BuildCanonicalRuntimeInventoryPickerItems.*?^End Function').Value
    [regex]::Match($shippingModule, '(?ms)^Private Function ShippingInventoryPickerTableHasRows.*?^End Function').Value
    [regex]::Match($shippingModule, '(?ms)^Private Function BuildShippingInventoryPickerItems.*?^End Function').Value
) -join "`n"
$shippingBomIdentityBlock = (@(
    'QueueBoxMakerFormPayload'
    'QueueBoxBuildEventFromBuilder'
    'QueueBoxUnboxEventFromBuilder'
    'AddBoxBuildComponentPayloadItems'
    'AddBoxUnboxComponentPayloadItems'
    'AddBoxBuildPayloadItem'
    'ResolveItemCodeForBoxBuildPayload'
    'CommitBoxBuilderFormState'
    'CollectBomComponents'
    'EnsureBoxBomEntryColumns'
    'ResolveCanonicalComponentInfoShipping'
    'SaveShippingBomToRuntime'
    'ShippingBomComponentSignatureFromCollection'
    'WriteShippingBomPackageTable'
    'LoadBoxMakerBomForPackage'
    'LoadBoxBomForPackageVersion'
    'LoadShippableVersionInventoryCore'
    'BoxMakerFormLoadVersionComponents'
    'BoxBuilderFormCurrentComponents'
    'BoxMakerFormLoadShippableInventory'
    'BuildBoxVersionInventoryCache'
    'ExtractBoxPackageSystemKeyFromNoteShipping'
    'EvictCompletedShipmentInventoryOverlaysForShippables'
    'ResolveBoxMakerUnboxAvailableQty'
    'ResolveCurrentInventoryBySystemKey'
) | ForEach-Object {
    [regex]::Match($shippingModule, "(?ms)^(?:Public|Private) (?:Function|Sub) $_.*?^End (?:Function|Sub)").Value
}) -join "`n"
$shippingBomIdentityBlock = $shippingBomIdentityBlock -replace `
    '(?m)^.*RemoveColumnIfExistsShipping[^\r\n]*"ROW"[^\r\n]*\r?$', ''

$seedRows = [regex]::Matches($seedService, '(?m)^\s*AddDemoInventoryItem\s+rows,').Count
$shippingCodes = @(
    "DEMO-PKG-SHIPPING-CARTON",
    "DEMO-PKG-CASE-DIVIDER",
    "DEMO-PKG-SHIPPING-LABEL",
    "DEMO-PKG-PACKING-TAPE",
    "DEMO-PKG-VOID-FILL"
)
$missingStationIdentityHarnesses = @(
    Get-ChildItem -LiteralPath (Join-Path $repo "tools") -Filter "*.ps1" -File |
        Where-Object {
            $content = Get-Content -Raw -LiteralPath $_.FullName
            $content -match 'src[\\/]Core[\\/]Modules[\\/]modConfig\.bas' -and
                $content -notmatch 'src[\\/]Core[\\/]Modules[\\/]modStationIdentity\.bas'
        }
)

$checks = @(
    [pscustomobject]@{
        Name = "Workbook.ExactRoleNames"
        Passed = @(
            '.Receiving.Operator.xlsm',
            '.Production.Operator.xlsm',
            '.Shipping.Operator.xlsm'
        ) | ForEach-Object { $bootstrap.Contains($_) } |
            Where-Object { -not $_ } | Measure-Object |
            Select-Object -ExpandProperty Count | ForEach-Object { $_ -eq 0 }
        Contract = "Warehouse bootstrap uses the three exact WarehouseId.Role.Operator.xlsm filenames."
    },
    [pscustomobject]@{
        Name = "Station.ComputerIdentity"
        Passed = ($stationIdentity -match 'Public Function CurrentComputerStationId') -and
            ($stationIdentity -match 'Environ\$\("COMPUTERNAME"\)') -and
            ($seedForm -match 'modStationIdentity\.CurrentComputerStationId') -and
            ($connectionForm -match 'modStationIdentity\.CurrentComputerStationId') -and
            ($seedForm -notmatch '"S1"') -and
            ($connectionForm -notmatch 'mCboStation|mChkRequireStation')
        Contract = "Seed and connection forms derive station identity from the Windows computer name without an S1 or selector path."
    },
    [pscustomobject]@{
        Name = "Station.HarnessDependency"
        Passed = ($missingStationIdentityHarnesses.Count -eq 0)
        Contract = "Every standalone harness that imports modConfig also imports its computer-station identity dependency."
    },
    [pscustomobject]@{
        Name = "Admin.OneDesignLifecycleLauncher"
        Passed = ([regex]::Matches($adminRibbon, 'btnAdminDesignLifecycle').Count -eq 1) -and
            ($adminRibbon -match 'Label = "Design Lifecycle"') -and
            ($adminRibbon -notmatch 'btnAdminReleaseDesign|btnAdminObsoleteDesign')
        Contract = "Admin exposes one Design Lifecycle launcher; release and obsolete remain actions inside its form."
    },
    [pscustomobject]@{
        Name = "Admin.ClearAddModeCaption"
        Passed = ($adminForm -match 'AddButton\("btnAddMode"[^\r\n]*"Add Item Mode"\)')
        Contract = "The mode selector cannot be mistaken for the button that commits an item."
    },
    [pscustomobject]@{
        Name = "Admin.TestEnvironmentWording"
        Passed = ($adminRibbon -match 'Label = "Test Environment Setup"') -and
            ($adminRibbon -notmatch 'Label = "Setup Tester Station"')
        Contract = "The retained admin utility is identified as isolated test-environment provisioning."
    },
    [pscustomobject]@{
        Name = "Seed.BoxMakingMaterials"
        Passed = ($seedRows -eq 24) -and (@($shippingCodes | Where-Object { -not $seedService.Contains($_) }).Count -eq 0)
        Contract = "The 24-item demo kit includes five explicit consumables needed to build shipping boxes."
    },
    [pscustomobject]@{
        Name = "Boxing.DesignerTerminology"
        Passed = ($shippingForm -match 'Tabs\.Add "tabBoxBuilder", "Box Designer"') -and
            ($shippingForm -match '"Update Alternative"') -and
            ($shippingForm -match '"New Alternative"') -and
            ($shippingForm -match '"Delete Alternative"') -and
            ($choiceForm -match 'Alternative')
        Contract = "Operator wording uses Box Designer and alternative, while durable internal version keys may remain compatible."
    },
    [pscustomobject]@{
        Name = "Boxing.FullWidthResponsiveLayout"
        Passed = ($shippingForm -match 'Private Sub LayoutBoxDesignerPage') -and
            ($shippingForm -match 'Private Sub LayoutBoxMakerPage') -and
            ($shippingForm -match 'BoxDesignerOverlaps=False') -and
            ($shippingForm -match 'BoxMakerOverlaps=False')
        Contract = "Designer and Maker lists receive explicit full-width responsive layouts with overlap evidence."
    },
    [pscustomobject]@{
        Name = "Boxing.HeadersTrackColumns"
        Passed = ($shippingForm -match 'ApplyBoxingHeaderLayout') -and
            ($shippingForm -match 'HeaderColumnsAligned=True') -and
            ($shippingForm -match 'Array\("", "Box / assembly", "", "UOM", "Location", "", ""\)') -and
            ($shippingForm -match 'If i <= UBound\(headings\) Then')
        Contract = "Boxing list headers are recalculated and tested against their list columns after resizing."
    },
    [pscustomobject]@{
        Name = "Boxing.ComponentIdentityIsSystemKey"
        Passed = ($shippingPickerBlock -match 'tblInventoryEntities') -and
            ($shippingPickerBlock -match 'ColumnIndex\([^\r\n]+"System_Key"\)') -and
            ($shippingPickerBlock -notmatch '"ROW"')
        Contract = "Box Designer component choices carry immutable System_Key and never a managed ROW surrogate."
    },
    [pscustomobject]@{
        Name = "Boxing.BomPersistenceIdentity"
        Passed = ($shippingBomIdentityBlock -match 'ComponentSystemKey') -and
            ($shippingBomIdentityBlock -match 'System_Key') -and
            ($shippingModule -match 'EnsureColumnExists loBom, "System_Key"') -and
            ($shippingModule -match 'RemoveColumnIfExistsShipping loBom, "ROW"') -and
            ($shippingBomIdentityBlock -cnotmatch '"ROW"|invSys ROW|Package ROW|component ROW') -and
            ($boxingService -match 'ResolveBoxingInventorySystemKeyForSku') -and
            ($boxingService -match 'modRoleEventWriter\.CreateSystemKey') -and
            ($boxingService -notmatch 'componentRows\(1, 4\) = 1|PostBoxMakerAction\(operatorWb, 1')
        Contract = "Box Designer save, alternative load, runtime BOM persistence, matching, and Box Maker events preserve string System_Key identity."
    },
    [pscustomobject]@{
        Name = "Boxing.RuntimeQueueUsesCoreApi"
        Passed = ($shippingModule -match 'modRoleEventWriter\.SyncLocalStagedInboxRows') -and
            ($shippingModule -notmatch 'SyncLocalStagedInboxSystemKeys')
        Contract = "The public Box Maker action calls the existing Core staging-sync API before processor execution."
    },
    [pscustomobject]@{
        Name = "Receiving.SearchSelectionSurface"
        Passed = ($receivingForm -match 'mTxtItemSearch') -and
            ($receivingForm -match 'mLstReceiveItems') -and
            ($receivingForm -match 'Receive item search') -and
            ($receivingForm -match 'Receive Item Results')
        Contract = "Receiving has a dedicated searchable result list with visible item details."
    },
    [pscustomobject]@{
        Name = "Receiving.LocationAndLot"
        Passed = ($receivingForm -match 'mTxtReceiveLocation') -and
            ($receivingForm -match 'mTxtLotNumber') -and
            ($receivingService -match 'locationOverride') -and
            ($receivingService -match 'lotNumber') -and
            ($receivingSurface -match 'LOT_NUMBER') -and
            ($receivingPosting -match 'BuildReceivingAttributesJson')
        Contract = "A receiving entry requires a location and carries an optional lot through staging, log, and durable inventory attributes."
    },
    [pscustomobject]@{
        Name = "Receiving.SchemaExpansionAvoidsOverlap"
        Passed = ($receivingEnsureBlock.IndexOf('ArrangeReceivingTablesSurface wb') -ge 0) -and
            ($receivingEnsureBlock.IndexOf('ArrangeReceivingTablesSurface wb') -lt $receivingEnsureBlock.IndexOf('EnsureTableSurface wb')) -and
            ($receivingEnsureBlock -match 'AggregateReceived.+NextFreeReceivingTableAddressSurface') -and
            ($receivingEnsureBlock -match 'invSysData_Receiving.+NextFreeReceivingTableAddressSurface') -and
            ($receivingArrangeBlock -match 'ProjectedTableColumnCountSurface') -and
            ($receivingArrangeBlock -match 'inventoryStartColumn') -and
            ($receivingArrangeBlock -match 'aggregateStartColumn\s*=\s*3\s*\+\s*receivedWidth\s*\+\s*1') -and
            ($receivingArrangeBlock.IndexOf('MoveTableTopLeftAtCellSurface ws, "invSysData_Receiving"') -lt $receivingArrangeBlock.IndexOf('MoveTableTopLeftAtCellSurface ws, "AggregateReceived"')) -and
            ($receivingArrangeBlock.IndexOf('MoveTableTopLeftAtCellSurface ws, "AggregateReceived"') -lt $receivingArrangeBlock.IndexOf('MoveTableTopLeftAtCellSurface ws, "ReceivedTally"')) -and
            ($receivingRebuildBlock -match 'Err\.Raise')
        Contract = "Existing Receiving tables reserve required and unknown user-column widths, then move right-to-left before Location/Lot columns expand their schemas."
    },
    [pscustomobject]@{
        Name = "Receiving.HeadersTrackColumns"
        Passed = ($receivingForm -match 'ApplyReceivingHeaderLayout') -and
            ($receivingForm -match 'ReceivingHeaderColumnsAligned=True') -and
            ($receivingService -match 'Public Function RunReceivingSearchAndHeaderContractTest')
        Contract = "Receiving item, history, tally, and aggregate headers track their list columns."
    },
    [pscustomobject]@{
        Name = "Production.ListBatchScaling"
        Passed = ($productionForm -match 'mTxtBatchScalePercent') -and
            ($productionForm -match 'BATCH_SCALE_MIN_PERCENT As Double = 0\.001') -and
            ($productionForm -match 'BATCH_SCALE_MAX_PERCENT As Double = 1000') -and
            ($productionForm -match 'ApplyBatchScaleToRunList') -and
            ($productionForm -match 'Private Function IsBatchScaleRunTable') -and
            ($productionForm -match 'If IsBatchScaleRunTable\(lo\) Then') -and
            ($productionForm -match 'TestBatchScaleContract') -and
            ($productionModule -match 'Public Function RunProductionBatchScaleContractTest')
        Contract = "Production Run - List exposes and applies a 0.001% through 1000% batch scale."
    }
)

$passed = @($checks | Where-Object Passed).Count
$failed = $checks.Count - $passed
$resultPath = Join-Path $repo "tests\unit\plan022_slice4j_workflow_readiness_results.md"
$lines = @(
    "# Plan 022 Slice 4j Workflow Readiness Results", "",
    "- Passed: $passed", "- Failed: $failed", "",
    "| Check | Result | Contract |", "|---|---|---|"
)
foreach ($check in $checks) {
    $result = if ($check.Passed) { "PASS" } else { "FAIL" }
    $lines += "| $($check.Name) | $result | $($check.Contract) |"
}
[IO.File]::WriteAllText($resultPath, (($lines -join "`n") + "`n"), (New-Object Text.UTF8Encoding($false)))

Write-Host "PLAN022_SLICE4J_RESULTS=$resultPath"
Write-Host "PASSED=$passed FAILED=$failed TOTAL=$($checks.Count)"
if ($failed -gt 0) { exit 1 }
