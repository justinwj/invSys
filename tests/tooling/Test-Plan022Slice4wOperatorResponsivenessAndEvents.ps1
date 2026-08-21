param(
    [string]$RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
)

$ErrorActionPreference = "Stop"
$repo = (Resolve-Path $RepoRoot).Path

function Read-Source([string]$relativePath) {
    Get-Content -Raw -LiteralPath (Join-Path $repo $relativePath)
}

function Procedure-Text([string]$text, [string]$name) {
    [regex]::Match(
        $text,
        "(?ms)^(?:Public|Private) (?:Function|Sub) $([regex]::Escape($name))\b.*?^End (?:Function|Sub)"
    ).Value
}

$connectionForm = Read-Source "src/Core/Forms/frmWarehouseConnection.frm"
$roleWriter = Read-Source "src/Core/Modules/modRoleEventWriter.bas"
$warehouseSync = Read-Source "src/Core/Modules/modWarehouseSync.bas"
$viewerData = Read-Source "src/Core/Modules/modInventoryViewerData.bas"
$viewerForm = Read-Source "src/Operations/Forms/frmInventoryViewer.frm"
$operationsAnchors = Read-Source "src/Operations/ClassModules/cOperationsAnchorManager.cls"
$viewerController = Read-Source "src/Operations/Modules/modInventoryViewer.bas"
$receivingForm = Read-Source "src/Receiving/Forms/frmReceiving.frm"
$shippingForm = Read-Source "src/Shipping/Forms/frmShipmentsTally.frm"
$shippingService = Read-Source "src/Shipping/Modules/modTS_Shipments.bas"

$connectClick = Procedure-Text $connectionForm "mBtnConnect_Click"
$connectAction = Procedure-Text $roleWriter "ConnectWarehouseStorageForCapability"
$receivingConfirm = Procedure-Text $receivingForm "mBtnConfirm_Click"
$receivingAggregateClick = Procedure-Text $receivingForm "mLstAggregate_Click"
$viewerBuild = Procedure-Text $viewerForm "BuildLayout"
$viewerTab = Procedure-Text $viewerForm "ApplyViewerTab"
$viewerRefreshEvents = Procedure-Text $viewerForm "RefreshEvents"
$viewerEventsTest = Procedure-Text $viewerForm "TestEventsReport"
$shippingCommit = Procedure-Text $shippingForm "CommitCurrentLine"
$shippingSend = Procedure-Text $shippingForm "mBtnSend_Click"
$snapshotAction = Procedure-Text $warehouseSync "GenerateWarehouseSnapshot"

$checks = @(
    [pscustomobject]@{
        Check = "ServerConnection.ProgressBeforeBlockingIO"
        Passed = ($connectClick -match 'Connecting to warehouse storage') -and
            ($connectClick -match 'Me\.Repaint') -and ($connectClick -match 'DoEvents') -and
            ($connectAction -match 'BeginServerConnectionProgressRole') -and
            ($connectAction -match 'EndServerConnectionProgressRole')
        Contract = "Manual and ribbon Server Sign In render progress before the synchronous Windows SMB call and restore Excel UI afterward."
    },
    [pscustomobject]@{
        Check = "Receiving.AggregateReferenceDetail"
        Passed = ($receivingForm -match 'txtAggregateReferences') -and
            ($receivingForm -match 'Selected references') -and
            ($receivingForm -match 'mTxtAggregateReferences\.MultiLine\s*=\s*True') -and
            ($receivingAggregateClick -match 'ShowSelectedAggregateReferences') -and
            ($receivingForm -match 'ClearAggregateReferenceDetail')
        Contract = "The fixed-height aggregate list retains one-line rows while a dedicated multiline detail surface shows every concatenated reference and clears with staging."
    },
    [pscustomobject]@{
        Check = "Viewer.Events.ReadOnlyTab"
        Passed = ($viewerBuild -match 'Tabs\(0\)\.Caption\s*=\s*"Inventory"') -and
            ($viewerBuild -match 'Tabs\(1\)\.Caption\s*=\s*"Events"') -and
            ($viewerTab -match 'RefreshEvents') -and
            ($viewerRefreshEvents -match 'LoadCurrentInventoryEventViewerData') -and
            ($viewerController -match 'LoadShippingViewerSupplementEvents') -and
            ($shippingService -match 'BOX_DESIGNED') -and
            ($shippingService -match 'SHIP_HELD')
        Contract = "Inventory Viewer exposes a read-only Events tab covering canonical inventory events plus current box-design and held-shipment activity."
    },
    [pscustomobject]@{
        Check = "Viewer.Tabs.ExactlyInventoryAndEvents"
        Passed = ($viewerBuild -notmatch 'Tabs\.Add\s+"tabInventory"') -and
            ($viewerBuild -notmatch 'Tabs\.Add\s+"tabEvents"') -and
            ($viewerEventsTest -match 'TabCount=') -and
            ($viewerEventsTest -match 'TabCaptions=') -and
            ($viewerEventsTest -match 'SelectedTab=')
        Contract = "The runtime Viewer does not append duplicate placeholder pages, exposes exactly Inventory and Events, and its public Events action selects the operator-visible Events tab."
    },
    [pscustomobject]@{
        Check = "Viewer.Layout.GuardsNativeWindowState"
        Passed = ($operationsAnchors -match 'GetUserFormWindowHandle') -and
            ($operationsAnchors -match 'IsIconic') -and ($operationsAnchors -match 'IsZoomed') -and
            ($operationsAnchors -match 'ApplyMinimumFormSize') -and
            ($operationsAnchors -match 'Err\.Number\s*=\s*384')
        Contract = "Operations anchoring skips native form-size enforcement while minimized or maximized and contains residual run-time error 384 without disabling restored-state layout."
    },
    [pscustomobject]@{
        Check = "Viewer.Events.ReadableTimestampRefresh"
        Passed = ($viewerData -match 'Format\$\(CDate\(CDbl\(eventDateText\)\),\s*"yyyy-mm-dd hh:nn:ss"\)') -and
            ($viewerEventsTest -match 'ReadableDates=') -and
            ($viewerEventsTest -match 'FirstReference=')
        Contract = "Events renders readable timestamps and the public Events refresh reports the newly published first event rather than retaining stale rows."
    },
    [pscustomobject]@{
        Check = "Viewer.Events.PublishedProjection"
        Passed = ($snapshotAction -match 'WriteSnapshotEventRows') -and
            ($warehouseSync -match 'tblInventoryEvents') -and
            ($viewerData -match 'Public Function LoadCurrentInventoryEventViewerData') -and
            ($viewerData -match 'tblInventoryEvents')
        Contract = "Viewer event history is read from the published snapshot projection rather than making the form a canonical writer or authority."
    },
    [pscustomobject]@{
        Check = "Viewer.Events.RemoveRelease"
        Passed = ($viewerData -match 'Case\s+"SHIP_RELEASE"\s*:\s*ViewerFriendlyEventType\s*=\s*"Remove"') -and
            ($shippingService -match 'EVENT_TYPE_SHIP_RELEASE')
        Contract = "Shipping Remove releases locked inventory through SHIP_RELEASE and the operator-facing Events view labels that event Remove."
    },
    [pscustomobject]@{
        Check = "OperatorPersistence.PendingStatus"
        Passed = ($receivingConfirm -match 'ShowPersistencePending') -and
            ($receivingConfirm -match 'ExecuteConfirmWrites') -and
            ($shippingCommit -match 'ShowPersistencePending') -and
            ($shippingSend -match 'ShowPersistencePending') -and
            ($shippingForm -match 'Me\.Repaint') -and ($shippingForm -match 'DoEvents')
        Contract = "Receiving/Returns and Shipping render their own saving-to-server status before required persistence begins; Office-native progress UI remains separate."
    }
)

$passed = @($checks | Where-Object Passed).Count
$failed = $checks.Count - $passed
$resultPath = Join-Path $repo "tests/integration/plan022_slice4w_operator_responsiveness_and_events_results.md"
$lines = @(
    "# Plan 022 Slice 4w Operator Responsiveness and Events Results",
    "",
    "- Passed: $passed",
    "- Failed: $failed",
    "",
    "| Check | Result | Contract |",
    "|---|---|---|"
)
foreach ($check in $checks) {
    $lines += "| $($check.Check) | $(if ($check.Passed) { 'PASS' } else { 'FAIL' }) | $($check.Contract) |"
}
Set-Content -LiteralPath $resultPath -Value $lines -Encoding utf8
$checks | Format-Table -AutoSize
Write-Host "Plan 022 Slice 4w operator responsiveness/events: $passed passed, $failed failed"
if ($failed -gt 0) { exit 1 }
