[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

$repo = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path
$buildPath = Join-Path $repo "tools\build-xlam.ps1"
$specPath = Join-Path $repo "..\invSys_docs\0 plan docs\xlam_invSys\invSys-Design-v4.11.md"
$resultPath = Join-Path $repo "tests\unit\r1_control_surface_results.md"
$rows = New-Object System.Collections.Generic.List[object]

function Add-Check {
    param([string]$Name, [bool]$Passed, [string]$Contract)
    $rows.Add([pscustomobject]@{ Name = $Name; Passed = $Passed; Contract = $Contract })
}

$obsoleteForms = @(
    "src/Admin/Forms/frmAdminControls.frm",
    "src/Admin/Forms/frmAdminEmail.frm",
    "src/Admin/Forms/frmEditUser.frm",
    "src/Admin/Forms/ufAdminItemSearch.frm",
    "src/Admin/Forms/ufDynItemSearchTemplate.frm",
    "src/Production/Forms/frmCreateRecipeTable.frm",
    "src/Production/Forms/frmCreateSubstitutionList.frm",
    "src/Production/Forms/frmIngredientPalette.frm",
    "src/Production/Forms/frmSubstitution.frm",
    "src/Production/Forms/ufProductionItemSearch.frm",
    "src/Production/Forms/ufDynItemSearchTemplate.frm",
    "src/Receiving/Forms/frmReceivingSavedList.frm",
    "src/Receiving/Forms/ufReceivingItemSearch.frm",
    "src/Receiving/Forms/ufDynItemSearchTemplate.frm",
    "src/Shipping/Forms/frmShippingCreateList.frm",
    "src/Shipping/Forms/frmShippingSavedList.frm",
    "src/Shipping/Forms/ufShippingItemSearch.frm",
    "src/Shipping/Forms/ufDynItemSearchTemplate.frm"
)
$presentObsolete = @($obsoleteForms | Where-Object {
    Test-Path -LiteralPath (Join-Path $repo $_) -PathType Leaf
})
Add-Check "Forms.ObsoleteShellsRemoved" ($presentObsolete.Count -eq 0) `
    "The reviewed Release 1 package must not retain empty or unreachable form shells."

$requiredForms = @(
    "src/Admin/Forms/frmAddInventoryItem.frm",
    "src/Admin/Forms/frmAdminDesignLifecycle.frm",
    "src/Admin/Forms/frmAdminSettings.frm",
    "src/Admin/Forms/frmCreateDeleteUser.frm",
    "src/Admin/Forms/frmCreateWarehouse.frm",
    "src/Admin/Forms/frmReAuthGate.frm",
    "src/Admin/Forms/frmRetireMigrateWarehouse.frm",
    "src/Admin/Forms/frmSeedInventory.frm",
    "src/Admin/Forms/frmSetupTesterStation.frm",
    "src/Core/Forms/frmItemSearch.frm",
    "src/Core/Forms/frmSignIn.frm",
    "src/Core/Forms/frmWarehouseConnection.frm",
    "src/Receiving/Forms/frmReceiving.frm",
    "src/Production/Forms/frmProduction.frm",
    "src/Shipping/Forms/frmBoxVersionSaveChoice.frm",
    "src/Shipping/Forms/frmShipmentsTally.frm",
    "src/Operations/Forms/frmInventoryViewer.frm"
)
$missingRequired = @($requiredForms | Where-Object {
    -not (Test-Path -LiteralPath (Join-Path $repo $_) -PathType Leaf)
})
Add-Check "Forms.ReviewedSetPresent" ($missingRequired.Count -eq 0) `
    "Every reviewed active form, the Purchasing-bearing Receiving form, and the Inventory Viewer must remain present."

$receivingText = Get-Content -Raw -LiteralPath (Join-Path $repo "src\Receiving\Forms\frmReceiving.frm")
Add-Check "Receiving.PurchasingStubRetained" `
    (($receivingText -match 'Purchasing') -and ($receivingText -match 'stub|future|placeholder')) `
    "The reviewed Purchasing stub remains visible in the Receiving form."

$buildText = Get-Content -Raw -LiteralPath $buildPath
$viewerModulePath = Join-Path $repo "src\Operations\Modules\modInventoryViewer.bas"
$viewerModuleText = if (Test-Path -LiteralPath $viewerModulePath) {
    Get-Content -Raw -LiteralPath $viewerModulePath
} else { "" }
Add-Check "Viewer.RibbonVisibleForSignedInUsers" `
    (($buildText -match 'Label\s*=\s*"Inventory Viewer"') -and
     ($buildText -match 'Macro\s*=\s*"modInventoryViewer\.OpenInventoryViewer"') -and
     ($buildText -notmatch 'Inventory Viewer"[^\r\n]*RequiredCapability')) `
    "Operations exposes Inventory Viewer without a role capability restriction; the action itself requires sign-in."
Add-Check "Viewer.PublicAction" `
    (($viewerModuleText -match 'Public Sub OpenInventoryViewer') -and
     ($viewerModuleText -match 'RequireCurrentUser|CurrentUser|modAuth\.IsSignedIn')) `
    "The Viewer ribbon action is public and enforces a signed-in invSys session."

$specText = Get-Content -Raw -LiteralPath $specPath
Add-Check "D4.SharedSearchForm" `
    (($specText -match 'one runtime-built Core form') -and
     ($specText -match 'Empty role-named form copies are prohibited')) `
    "D4 names the real shared search-form boundary before the obsolete shells are removed."

$passed = @($rows | Where-Object Passed).Count
$failed = $rows.Count - $passed
$lines = @(
    "# Release 1 Control Surface Results", "",
    "- Passed: $passed", "- Failed: $failed", "",
    "| Check | Result | Contract |", "|---|---|---|"
)
foreach ($row in $rows) {
    $outcome = if ($row.Passed) { "PASS" } else { "FAIL" }
    $lines += "| $($row.Name) | $outcome | $($row.Contract) |"
}
if ($presentObsolete.Count -gt 0) {
    $lines += ""
    $lines += "Obsolete source components still present: $($presentObsolete -join ', ')"
}
if ($missingRequired.Count -gt 0) {
    $lines += ""
    $lines += "Required source components missing: $($missingRequired -join ', ')"
}
[IO.File]::WriteAllText($resultPath, (($lines -join "`n") + "`n"), (New-Object Text.UTF8Encoding($false)))

Write-Host "R1_CONTROL_SURFACE_RESULTS=$resultPath"
Write-Host "PASSED=$passed FAILED=$failed TOTAL=$($rows.Count)"
if ($failed -gt 0) { exit 1 }
