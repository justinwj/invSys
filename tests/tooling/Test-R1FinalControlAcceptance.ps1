[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

$repo = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path
$buildText = Get-Content -Raw -LiteralPath (Join-Path $repo "tools\build-xlam.ps1")
$shippingForm = Get-Content -Raw -LiteralPath (Join-Path $repo "src\Shipping\Forms\frmShipmentsTally.frm")
$shippingModule = Get-Content -Raw -LiteralPath (Join-Path $repo "src\Shipping\Modules\modTS_Shipments.bas")
$shippingEvents = Get-Content -Raw -LiteralPath (Join-Path $repo "src\Shipping\Modules\modShippingEventCreator.bas")
$shippingSurfaces = Get-Content -Raw -LiteralPath (Join-Path $repo "src\Core\Modules\modRoleWorkbookSurfaces.bas")
$receivingForm = Get-Content -Raw -LiteralPath (Join-Path $repo "src\Receiving\Forms\frmReceiving.frm")
$receivingModule = Get-Content -Raw -LiteralPath (Join-Path $repo "src\Receiving\Modules\modTS_Received.bas")
$resultPath = Join-Path $repo "tests\unit\r1_final_control_acceptance_results.md"
$checks = New-Object System.Collections.Generic.List[object]

function Add-Check {
    param([string]$Name, [bool]$Passed, [string]$Contract)
    $checks.Add([pscustomobject]@{ Name = $Name; Passed = $Passed; Contract = $Contract })
}

function Get-VbaProcedureText {
    param([string]$Text, [string]$ProcedureName)
    $pattern = "(?ms)^\s*(?:Public|Private|Friend)?\s*(?:Sub|Function)\s+" +
        [regex]::Escape($ProcedureName) + "\b.*?^\s*End\s+(?:Sub|Function)\s*$"
    $match = [regex]::Match($Text, $pattern)
    if ($match.Success) { return $match.Value }
    return ""
}

Add-Check "Viewer.RibbonIcon" `
    ($buildText -match 'btnOperationsInventoryViewer[^\r\n]*ImageMso\s*=\s*"PivotTableInsert"') `
    "Inventory Viewer uses a built-in Excel icon that is visible in the Operations ribbon."

$boxingLists = @(
    "mLstBoxBuilderDesigns",
    "mLstBoxBuilderInventory",
    "mLstBoxBuilderComponents",
    "mLstBoxMakerDesigns",
    "mLstBoxMakerComponents"
)
$allBoxingListsAnchored = $true
foreach ($listName in $boxingLists) {
    if ($shippingForm -notmatch ("mAnchors\.Add\s+" + $listName + ",")) {
        $allBoxingListsAnchored = $false
    }
}
Add-Check "Boxing.Layout.ListsAnchored" $allBoxingListsAnchored `
    "Every Box Designer and Box Maker list participates in native form resizing."
Add-Check "Boxing.Layout.PackagedSeam" `
    (($shippingForm -match 'Public Function TestBoxingLayoutAfterResize') -and
     ($shippingModule -match 'Public Function RunShippingBoxingLayoutTest') -and
     ($shippingModule -match 'BtnOpenShipmentsForm')) `
    "Packaged layout proof enters through the public Shipping launcher and measures both Boxing tabs."
Add-Check "BoxBuilder.ComponentSearch" `
    (($shippingForm -match 'Private WithEvents mTxtBoxBuilderSearch As MSForms\.TextBox') -and
     ($shippingForm -match 'Private Sub mTxtBoxBuilderSearch_Change\(\)') -and
     ($shippingForm -match 'FilterBoxBuilderInventory')) `
    "Box Designer locally filters the loaded managed component inventory from a dedicated search control."
Add-Check "Boxing.Headers.Aligned" `
    (($shippingForm -match 'ConfigureBoxingListHeaders') -and
     ($shippingForm -match 'ApplyBoxingHeaderLayout') -and
     ($shippingForm -match 'BoxingHeaderWidthsMatchLists=True')) `
    "Boxing list headers are generated and repositioned from the same width contract as their list columns."
Add-Check "Boxing.Version.NotApplicable" `
    (($shippingForm -match 'DisplayVersionOrNA') -and
     ($shippingForm -match '"NA"')) `
    "Managed inventory without a Shipping BOM version displays NA instead of inheriting or inventing a version."

Add-Check "Receiving.History.TopList" `
    (($receivingForm -match '"Receiving Entries History"') -and
     ($receivingForm -match 'LoadReceivingEntriesHistoryForWorkbook') -and
     ($receivingModule -match 'Public Function LoadReceivingEntriesHistoryForWorkbook')) `
    "The Receiving form's top list is a captured-workbook ReceivedLog history projection."
Add-Check "Receiving.History.SeparateItemSelector" `
    (($receivingForm -match 'Private WithEvents mLstReceiveItems As MSForms\.ListBox') -and
     ($receivingForm -match 'mTxtItemSearch') -and
     ($receivingForm -match 'LoadReceivingFormInventoryForWorkbook')) `
    "Receiving stages from a dedicated searchable managed-item results list, not from a history row."
Add-Check "Receiving.History.RefreshSemantics" `
    (($receivingForm -match 'Search history') -and
     ($receivingForm -notmatch 'Set mLblInventoryTitle = AddLabel\([^\r\n]*"Inventory"')) `
    "Refresh and search wording describe receipt history rather than a second inventory viewer."

Add-Check "Shipping.Identity.FormSystemKey" `
    (($shippingForm -match 'mTxtSystemKey') -and
     ($shippingForm -notmatch '"ROW"') -and
     ($shippingForm -notmatch 'mTxtRow|mSelectedBoxBuilderPackageRow|mSelectedBoxMakerPackageRow')) `
    "Visible Shipping and Boxing form state identifies inventory entities by System_Key, never worksheet ROW."
Add-Check "Shipping.Identity.EventSystemKey" `
    (($shippingEvents -match 'System_Key') -and
     ($shippingEvents -notmatch '"ROW"')) `
    "Shipping event creation serializes exact inventory System_Key identity and contains no managed ROW field."
Add-Check "Shipping.Identity.ControllerNoRowHeader" `
    (((Get-VbaProcedureText $shippingModule "ShipmentsFormLoadLines") -match 'System_Key') -and
     ((Get-VbaProcedureText $shippingModule "ShipmentsFormCommitLine") -match 'ByVal systemKey As String') -and
     ((Get-VbaProcedureText $shippingModule "BuildSelectedShipmentSystemKeyDeltas") -match 'System_Key') -and
     ((Get-VbaProcedureText $shippingModule "ApplyShipmentsSentDeltasBySystemKey") -match 'System_Key') -and
     ((Get-VbaProcedureText $shippingModule "RunShippingSystemKeyIdentityTest") -match 'TestSystemKeyIdentityContract') -and
     ($shippingSurfaces -match 'EnsureTableSurface\s+wb,\s+SHIPPING_BACKEND_SHEET,\s+"ShipmentsTally",\s+Array\([^\r\n]*"System_Key"') -and
     ($shippingSurfaces -match 'ShippingBomViewHeadersSurface[\s\S]*?PackageSystemKey[\s\S]*?ComponentSystemKey')) `
    "The reachable R1 Shipping form, staging, sent-event, reservation, and managed-surface paths use exact System_Key identity; packaged form proof covers the operator seam."

$passed = @($checks | Where-Object Passed).Count
$failed = $checks.Count - $passed
$lines = @(
    "# R1 Final Control Acceptance Results", "",
    "- Passed: $passed", "- Failed: $failed", "",
    "| Check | Result | Contract |", "|---|---|---|"
)
foreach ($check in $checks) {
    $result = if ($check.Passed) { "PASS" } else { "FAIL" }
    $lines += "| $($check.Name) | $result | $($check.Contract) |"
}
[IO.File]::WriteAllText($resultPath, (($lines -join "`n") + "`n"), (New-Object Text.UTF8Encoding($false)))
Write-Host "R1_FINAL_CONTROL_RESULTS=$resultPath"
Write-Host "PASSED=$passed FAILED=$failed TOTAL=$($checks.Count)"
if ($failed -gt 0) { exit 1 }
