[CmdletBinding()]
param(
    [string]$RepoRoot = "."
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$resultPath = Join-Path $repo "tests/unit/slice5_behavior_lock_results.md"
$receivingModule = Get-Content -LiteralPath (Join-Path $repo "src/Receiving/Modules/modTS_Received.bas") -Raw
$receivingForm = Get-Content -LiteralPath (Join-Path $repo "src/Receiving/Forms/frmReceiving.frm") -Raw
$productionModule = Get-Content -LiteralPath (Join-Path $repo "src/Production/Modules/mProduction.bas") -Raw
$productionForm = Get-Content -LiteralPath (Join-Path $repo "src/Production/Forms/frmProduction.frm") -Raw
$shippingModule = Get-Content -LiteralPath (Join-Path $repo "src/Shipping/Modules/modTS_Shipments.bas") -Raw
$shippingForm = Get-Content -LiteralPath (Join-Path $repo "src/Shipping/Forms/frmShipmentsTally.frm") -Raw
$buildScript = Get-Content -LiteralPath (Join-Path $repo "tools/build-xlam.ps1") -Raw

$rows = [System.Collections.Generic.List[object]]::new()
function Add-Check {
    param([string]$Name, [bool]$Passed, [string]$Detail)
    $rows.Add([pscustomobject]@{ Name = $Name; Passed = $Passed; Detail = $Detail }) | Out-Null
}

Add-Check "Receiving.FormAction.ConfirmWrites.Handler" `
    ($receivingForm -match '(?s)Private Sub mBtnConfirm_Click\(\).*?modTS_Received\.ConfirmWrites') `
    "The form button must call the operator Confirm Writes handler."
Add-Check "Production.FormActions.RequiredHandlers" `
    (($productionForm -match '(?s)Private Sub mBtnRunApplyPalette_Click\(\).*?ApplySelectedRunPaletteSplit') -and
     ($productionForm -match '(?s)Private Sub mBtnManagerCheckIn_Click\(\).*?CheckInProductionRun') -and
     ($productionForm -match '(?s)Private Sub mBtnManagerApplyOutput_Click\(\).*?CompleteProductionRun') -and
     ($productionForm -match '(?s)Private Sub mBtnManagerNext_Click\(\).*?BtnNextBatch')) `
    "Selection/Apply, Check In, Complete Run, and Next Batch must remain wired to the operator handlers."
Add-Check "Shipping.FormActions.RequiredHandlers" `
    (($shippingForm -match '(?s)Private Sub mBtnStage_Click\(\).*?RunShippingAction True') -and
     ($shippingForm -match '(?s)Private Sub mBtnSend_Click\(\).*?RunShippingAction False')) `
    "To Shipments and Shipments Sent must remain wired to the operator handlers."

Add-Check "Receiving.Form.ModelessLauncher" `
    ($receivingModule -match 'frm\.Show\s+vbModeless') `
    "Receiving launcher must open the main form modelessly."
Add-Check "Production.Form.ModelessLauncher" `
    ($productionModule -match 'frmProduction\.Show\s+vbModeless') `
    "Production launcher must open the main form modelessly."
Add-Check "Shipping.Form.ModelessLauncher" `
    ($shippingModule -match 'frmShipmentsTally\.Show\s+vbModeless') `
    "Shipping launcher must open the main form modelessly."

$purchasingStubPresent = ($receivingForm -match 'Forms\.MultiPage\.1') -and
                         ($receivingForm -match '(?i)Caption\s*=\s*"Purchasing"') -and
                         ($receivingForm -match '(?i)not yet operational|not operational')
Add-Check "Receiving.Navigation.PurchasingStub" $purchasingStubPresent `
    "Receiving must expose a selectable, visibly non-operational Purchasing tab."

$shippingTabbedShell = ($shippingForm -match 'Forms\.MultiPage\.1') -and
                       ($shippingForm -match '(?i)Caption\s*=\s*"Box Builder"') -and
                       ($shippingForm -match '(?i)Caption\s*=\s*"Box Maker"')
Add-Check "Shipping.Navigation.SingleTabbedShell" $shippingTabbedShell `
    "The main Shipping form must contain Box Builder and Box Maker tabs."

$operationsPackage = ($buildScript -match 'invSys\.Operations\.xlam')
Add-Check "Operations.Package.Exists" $operationsPackage `
    "The build map must define invSys.Operations.xlam."

$noSeparateShippingLaunchers = $operationsPackage -and
                               ($buildScript -notmatch '(?i)Box Builder.*Ribbon') -and
                               ($buildScript -notmatch '(?i)Box Maker.*Ribbon')
Add-Check "Operations.Ribbon.SingleShippingLauncher" $noSeparateShippingLaunchers `
    "Operations Ribbon must expose one Shipping launcher and no separate Box Builder or Box Maker buttons."

Add-Check "Receiving.Form.CapturedWorkbookState" `
    (($receivingForm -match 'Private mOperatorWorkbook As Workbook') -and
     ($receivingForm -match 'BoundWorkbook="\s*&\s*mOperatorWorkbook\.Name')) `
    "Receiving form must retain explicit operator-workbook state."
Add-Check "Production.Form.CapturedWorkbookState" `
    (($productionForm -match 'Private mOperatorWorkbook As Workbook') -and
     ($productionForm -match 'BoundWorkbook="\s*&\s*mOperatorWorkbook\.Name')) `
    "Production form must retain explicit operator-workbook state."
Add-Check "Shipping.Form.CapturedWorkbookState" `
    (($shippingForm -match 'Private mOperatorWorkbook As Workbook') -and
     ($shippingForm -match 'BoundWorkbook="\s*&\s*mOperatorWorkbook\.Name')) `
    "Shipping form must retain explicit operator-workbook state."

$passed = @($rows | Where-Object Passed).Count
$failed = $rows.Count - $passed
$lines = @(
    "# Slice 5 Packaged Behavior Lock Results",
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
[System.IO.File]::WriteAllLines($resultPath, $lines, [System.Text.UTF8Encoding]::new($false))

Write-Output "SLICE5_BEHAVIOR_LOCK_RESULTS=$resultPath"
Write-Output "PASSED=$passed FAILED=$failed TOTAL=$($rows.Count)"
if ($failed -gt 0) { exit 1 }
