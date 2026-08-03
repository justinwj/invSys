[CmdletBinding()]
param(
    [string]$RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
)

$ErrorActionPreference = "Stop"
$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$shippingRoot = Join-Path $repo "src/Shipping"
$modulePath = Join-Path $shippingRoot "Modules/modTS_Shipments.bas"
$postingServicePath = Join-Path $shippingRoot "Modules/modShippingPostingService.bas"
$boxingServicePath = Join-Path $shippingRoot "Modules/modBoxingService.bas"
$statePath = Join-Path $shippingRoot "ClassModules/cShippingWorkflowState.cls"
$formPath = Join-Path $shippingRoot "Forms/frmShipmentsTally.frm"
$buildPath = Join-Path $repo "tools/build-xlam.ps1"

$moduleText = Get-Content -Raw -LiteralPath $modulePath
$formText = Get-Content -Raw -LiteralPath $formPath
$buildText = Get-Content -Raw -LiteralPath $buildPath
$postingText = if (Test-Path -LiteralPath $postingServicePath) {
    Get-Content -Raw -LiteralPath $postingServicePath
} else { "" }
$boxingText = if (Test-Path -LiteralPath $boxingServicePath) {
    Get-Content -Raw -LiteralPath $boxingServicePath
} else { "" }
$stateText = if (Test-Path -LiteralPath $statePath) {
    Get-Content -Raw -LiteralPath $statePath
} else { "" }

$results = [System.Collections.Generic.List[object]]::new()
function Add-Check {
    param([string]$Name, [bool]$Passed, [string]$Detail)
    $results.Add([pscustomobject]@{
        Check = $Name
        Passed = $Passed
        Detail = $Detail
    }) | Out-Null
}

Add-Check "Shipping.Services.SeparatedStateAndActions" `
    ((Test-Path -LiteralPath $statePath -PathType Leaf) -and
     (Test-Path -LiteralPath $postingServicePath -PathType Leaf) -and
     (Test-Path -LiteralPath $boxingServicePath -PathType Leaf) -and
     ($stateText -match '(?i)STAGED') -and
     ($stateText -match '(?i)RESERVED') -and
     ($stateText -match '(?i)SUBMITTED') -and
     ($stateText -match '(?i)APPLIED') -and
     ($stateText -match '(?i)COMPLETED')) `
    "Shipping needs typed state plus separate shipment-posting and Boxing services."

Add-Check "Shipping.Form.SingleTabbedShell" `
    (($formText -match '(?i)MSForms\.(TabStrip|MultiPage)') -and
     ($formText -match '(?i)Box Builder') -and
     ($formText -match '(?i)Box Maker') -and
     ($formText -match '(?i)SelectShippingPageForTest')) `
    "The main Shipping form must own selectable Shipping, Box Builder, and Box Maker pages."

Add-Check "Shipping.Form.TabOperatorActionsReachable" `
    (($formText -match '(?i)mBtnBoxBuilderNew') -and
     ($formText -match '(?i)mBtnBoxBuilderAddComponent') -and
     ($formText -match '(?i)mBtnBoxBuilderRemoveComponent') -and
     ($formText -match '(?i)mBtnBoxBuilderSave') -and
     ($formText -match '(?i)mBtnBoxBuilderUpdateVersion') -and
     ($formText -match '(?i)mBtnBoxBuilderNewVersion') -and
     ($formText -match '(?i)mBtnBoxBuilderDeleteVersion') -and
     ($formText -match '(?i)mBtnBoxBuilderArchive') -and
     ($formText -match '(?i)mBtnBoxBuilderDelete') -and
     ($formText -match '(?i)mBtnBoxMakerMake') -and
     ($formText -match '(?i)mBtnBoxMakerUnmake') -and
     ($formText -match '(?is)Private\s+Sub\s+mBtnBoxBuilderSave_Click\s*\(\s*\).*?modBoxingService\.SaveBoxDesign') -and
     ($formText -match '(?is)Private\s+Sub\s+mBtnBoxBuilderDeleteVersion_Click\s*\(\s*\).*?modBoxingService\.DeleteBoxDesignVersion') -and
     ($formText -match '(?is)Private\s+Sub\s+mBtnBoxBuilderArchive_Click\s*\(\s*\).*?modBoxingService\.ArchiveBoxDesign') -and
     ($formText -match '(?is)Private\s+Sub\s+mBtnBoxBuilderDelete_Click\s*\(\s*\).*?modBoxingService\.DeleteBoxDesign') -and
     ($formText -match '(?is)Private\s+Sub\s+mBtnBoxMakerMake_Click\s*\(\s*\).*?modBoxingService\.PostBoxMakerAction') -and
     ($formText -match '(?is)Private\s+Sub\s+mBtnBoxMakerUnmake_Click\s*\(\s*\).*?modBoxingService\.PostBoxMakerAction') -and
     ($boxingText -match '(?i)Public\s+Function\s+DeleteBoxDesignVersion') -and
     ($boxingText -match '(?i)Public\s+Function\s+ArchiveBoxDesign') -and
     ($boxingText -match '(?i)Public\s+Function\s+DeleteBoxDesign') -and
     ($boxingText -match '(?i)ByVal\s+operatorWb\s+As\s+Workbook')) `
    "Every legacy Box Builder action and Box Maker make/unmake must be reachable as a real tab callback routed through the captured-workbook Boxing service."

$sendBody = ""
$sendMatch = [regex]::Match(
    $formText,
    '(?is)Private\s+Sub\s+mBtnSend_Click\s*\(\s*\)(?<body>.*?)End\s+Sub'
)
if ($sendMatch.Success) { $sendBody = $sendMatch.Groups["body"].Value }
Add-Check "Shipping.Form.RealSendUsesTypedService" `
    (($sendBody -match '(?i)modShippingPostingService\.ExecuteShipmentsSent') -and
     ($sendBody -match '(?i)mOperatorWorkbook')) `
    "The real Shipments Sent handler must delegate to the typed service with captured workbook context."

Add-Check "Shipping.Form.ModelessCapturedContext" `
    (($moduleText -match '(?i)(?:frm|mShipmentsLauncherForm)\.Show\s+vbModeless') -and
     ($formText -match '(?i)Private\s+mOperatorWorkbook\s+As\s+Workbook') -and
     ($formText -notmatch '(?i)\bApplication\.ActiveWorkbook\b')) `
    "The modeless Shipping shell must remain bound to its captured operator workbook."

$shippingProjectMatch = [regex]::Match(
    $buildText,
    '(?is)Key\s*=\s*"Operations".*?(?=\n\s*@\{\s*\n\s*Key\s*=|\n\)\s*$)'
)
$shippingProjectText = if ($shippingProjectMatch.Success) {
    $shippingProjectMatch.Value
} else { "" }
Add-Check "Shipping.Ribbon.SingleLauncher" `
    (($shippingProjectText -match '(?i)btnOperationsShippingForm') -and
     ($shippingProjectText -notmatch '(?i)btnShippingBoxBuilderForm') -and
     ($shippingProjectText -notmatch '(?i)btnShippingBoxMakerForm') -and
     ($shippingProjectText -notmatch '(?i)BtnOpenBoxBuilder') -and
     ($shippingProjectText -notmatch '(?i)BtnOpenBoxMaker')) `
    "The Shipping ribbon must launch only the main shell, with no direct Box Builder or Box Maker callbacks."

Add-Check "Shipping.Inventory.NasReadOnlyProjectedDerived" `
    (($postingText -match '(?i)Projected') -and
     ($postingText -match '(?i)NAS') -and
     ($postingText -notmatch '(?i)SetCell.*NAS\s*Inv') -and
     ($boxingText -notmatch '(?i)SetCell.*NAS\s*Inv')) `
    "NAS Inv is read-only; Projected Inv must be derived from canonical and declared overlay inputs."

Add-Check "Shipping.Lock.ExactReservationIdentity" `
    (($postingText -match '(?i)(ShipmentLineId|ReservationEventId)') -and
     ($postingText -match '(?i)SHIP_RELEASE') -and
     ($postingText -match '(?i)ReleaseExact')) `
    "Remove and completion must release the exact active reservation identity."

Add-Check "Shipping.Events.IdempotentStableIdentity" `
    (($stateText -match '(?i)EnsureEventIdentit') -and
     ($postingText -match '(?i)(SKIP_DUP|Idempotent|EventId)')) `
    "Shipment replay must preserve one stable event identity and remain idempotent."

Add-Check "Shipping.Restart.CompletedStateDoesNotResurrect" `
    (($postingText -match '(?i)(Tombstone|CompletedLine)') -and
     ($postingText -match '(?i)(Restart|Reload|Restore)') -and
     ($postingText -match '(?i)Clear.*Staging')) `
    "Restart/reopen must retain completion tombstones and must not resurrect staged or locked lines."

$builderLauncher = [regex]::Match(
    $moduleText,
    '(?is)Public\s+Sub\s+BtnOpenBoxBuilder\s*\(\s*\)(?<body>.*?)End\s+Sub'
).Groups["body"].Value
$makerLauncher = [regex]::Match(
    $moduleText,
    '(?is)Public\s+Sub\s+BtnOpenBoxMaker\s*\(\s*\)(?<body>.*?)End\s+Sub'
).Groups["body"].Value
Add-Check "Shipping.Legacy.SeparateLaunchersRetired" `
    (($builderLauncher -notmatch '(?i)\.Show') -and
     ($makerLauncher -notmatch '(?i)\.Show') -and
     ($builderLauncher -match '(?i)BtnOpenShipmentsForm') -and
     ($makerLauncher -match '(?i)BtnOpenShipmentsForm')) `
    "Legacy callbacks must not open separate Box Builder or Box Maker forms."

$results | Format-Table -AutoSize
$failed = @($results | Where-Object { -not $_.Passed })
Write-Host ("Slice 11 Shipping/Boxing stabilization contract: {0} passed, {1} failed" -f
    ($results.Count - $failed.Count), $failed.Count)
if ($failed.Count -gt 0) { exit 1 }
