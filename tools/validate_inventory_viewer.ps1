[CmdletBinding()]
param(
    [string]$RepoRoot = ".",
    [string]$DeployRoot = "deploy/current"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Import-FunctionDefinitions {
    param([string]$ScriptPath)
    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile(
        $ScriptPath, [ref]$tokens, [ref]$errors
    )
    if ($errors.Count -gt 0) { throw "Unable to parse helper source: $($errors[0].Message)" }
    foreach ($definition in $ast.FindAll({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst]
    }, $true)) {
        $scriptDefinition = $definition.Extent.Text -replace (
            '^(?i)function\s+' + [regex]::Escape($definition.Name)
        ), ('function script:' + $definition.Name)
        . ([scriptblock]::Create($scriptDefinition))
    }
}

$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$deploy = (Resolve-Path -LiteralPath (Join-Path $repo $DeployRoot)).Path
Import-FunctionDefinitions -ScriptPath (Join-Path $repo "tools\validate_phase6_live_role_workflows.ps1")

$runtimeRoot = Join-Path ([IO.Path]::GetTempPath()) (
    "invsys-inventory-viewer-" + [guid]::NewGuid().ToString("N")
)
$warehouseId = "WHV" + [guid]::NewGuid().ToString("N").Substring(0, 6).ToUpperInvariant()
$stationId = "S1"
$testUser = if ([string]::IsNullOrWhiteSpace($env:USERNAME)) { "user1" } else { $env:USERNAME }
$testPin = [guid]::NewGuid().ToString("N")
$testPinHash = Get-InvSysCredentialHash -Credential $testPin
$configPath = Join-Path $runtimeRoot ($warehouseId + ".invSys.Config.xlsb")
$authPath = Join-Path $runtimeRoot ($warehouseId + ".invSys.Auth.xlsb")
$inventoryPath = Join-Path $runtimeRoot ($warehouseId + ".invSys.Data.Inventory.xlsb")
$snapshotPath = Join-Path $runtimeRoot ($warehouseId + ".invSys.Snapshot.Inventory.xlsb")
$resultPath = Join-Path $repo "tests\integration\inventory_viewer_results.md"
$excel = $null
$opened = New-Object System.Collections.Generic.List[object]
$facts = [ordered]@{}
$passed = $false
$detail = ""
$step = "startup"

try {
    New-Item -ItemType Directory -Path $runtimeRoot -Force | Out-Null
    $step = "start isolated Excel"
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $true
    $excel.AutomationSecurity = 1

    $step = "create isolated runtime"
    $configWb = New-ConfigWorkbook -Excel $excel -Path $configPath `
        -WarehouseId $warehouseId -StationId $stationId -RuntimeRoot $runtimeRoot
    $authWb = New-AuthWorkbook -Excel $excel -Path $authPath `
        -WarehouseId $warehouseId -StationId $stationId `
        -CurrentUserIds @($testUser) -CredentialHash $testPinHash
    $inventoryWb = New-InventoryWorkbook -Excel $excel -Path $inventoryPath `
        -WarehouseId $warehouseId -SkuRows @("SKU-SHIP", "SKU-SUGAR", "SKU-COMP")
    $inventoryLog = $inventoryWb.Worksheets.Item("InventoryLog").ListObjects.Item("tblInventoryLog")
    Add-ListObjectRow -ListObject $inventoryLog -Values @{
        "EventID" = "EVT-VIEWER-RECEIVE"
        "AppliedSeq" = 4
        "EventType" = "RECEIVE"
        "OccurredAtUTC" = [datetime]::UtcNow
        "AppliedAtUTC" = [datetime]::UtcNow
        "WarehouseId" = $warehouseId
        "StationId" = $stationId
        "UserId" = $testUser
        "System_Key" = "SYS-VIEWER-RECEIVE"
        "SKU" = "SKU-SHIP"
        "QtyDelta" = 2
        "Location" = "DOCK"
        "Condition" = "GOOD"
        "AttributesJson" = "{}"
        "Note" = "Reference=BOL-VIEWER;Item=Viewer Shipment Item;UOM=each"
    }
    Add-ListObjectRow -ListObject $inventoryLog -Values @{
        "EventID" = "EVT-VIEWER-SHIP-REMOVE"
        "AppliedSeq" = 5
        "EventType" = "SHIP_RELEASE"
        "OccurredAtUTC" = [datetime]::UtcNow
        "AppliedAtUTC" = [datetime]::UtcNow
        "WarehouseId" = $warehouseId
        "StationId" = $stationId
        "UserId" = $testUser
        "System_Key" = "SYS-LIVE-SHIP"
        "SKU" = "SKU-SHIP"
        "QtyDelta" = 0
        "Location" = "DOCK"
        "Condition" = "GOOD"
        "AttributesJson" = "{}"
        "Note" = "Reference=SHIP-REMOVE-VIEWER;Item=Viewer Shipment Item;UOM=each"
    }
    $inventoryWb.Save()
    $opened.Add($configWb) | Out-Null
    $opened.Add($authWb) | Out-Null
    $opened.Add($inventoryWb) | Out-Null

    $step = "open packaged add-ins"
    $packages = @{}
    foreach ($packageName in @(
        "invSys.Core.xlam", "invSys.Inventory.Domain.xlam",
        "invSys.Designs.Domain.xlam", "invSys.Operations.xlam"
    )) {
        $package = $excel.Workbooks.Open((Join-Path $deploy $packageName))
        $opened.Add($package) | Out-Null
        $packages[$packageName] = $package
    }
    $coreName = [string]$packages["invSys.Core.xlam"].Name
    $operationsName = [string]$packages["invSys.Operations.xlam"].Name

    $step = "configure signed-in target"
    [void](Run-WorkbookMacro -Excel $excel -WorkbookName $coreName `
        -MacroName "modRuntimeWorkbooks.SetCoreDataRootOverride" -Arguments @($runtimeRoot))
    $configLoaded = [bool](Run-WorkbookMacro -Excel $excel -WorkbookName $coreName `
        -MacroName "modConfig.LoadConfig" -Arguments @($warehouseId, $stationId))
    $authLoaded = [bool](Run-WorkbookMacro -Excel $excel -WorkbookName $coreName `
        -MacroName "modAuth.LoadAuth" -Arguments @($warehouseId))
    $targetResult = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $coreName `
        -MacroName "modNasConnection.SelectWarehouseTargetForAutomation" `
        -Arguments @($runtimeRoot, $runtimeRoot, $stationId, $true))
    $targetPathsSet = [bool](Run-WorkbookMacro -Excel $excel -WorkbookName $coreName `
        -MacroName "modNasConnection.SetCurrentTargetPathsForTest" `
        -Arguments @("\\inventory-viewer-test\warehouse", $runtimeRoot))
    $signInResult = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $coreName `
        -MacroName "modAuth.SignInCurrentTargetForAutomation" `
        -Arguments @($testUser, $testPin, "RECEIVE_POST"))

    $step = "publish isolated snapshot"
    $snapshotCreated = [bool](Run-WorkbookMacro -Excel $excel -WorkbookName $coreName `
        -MacroName "modWarehouseSync.GenerateWarehouseSnapshot" `
        -Arguments @($warehouseId, $inventoryWb, $snapshotPath))
    if (-not $snapshotCreated) { throw "The isolated inventory snapshot was not created." }
    $inventoryWb.Close($true)
    $authWb.Close($false)
    $configWb.Close($false)
    $snapshotHashBefore = (Get-FileHash -LiteralPath $snapshotPath -Algorithm SHA256).Hash

    $step = "invoke public Viewer action twice"
    [void](Run-WorkbookMacro -Excel $excel -WorkbookName $operationsName `
        -MacroName "modOperationsInit.Auto_Open")
    $firstReport = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $operationsName `
        -MacroName "modInventoryViewer.RunInventoryViewerActionForTest")
    $secondReport = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $operationsName `
        -MacroName "modInventoryViewer.RunInventoryViewerActionForTest")
    $filterReport = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $operationsName `
        -MacroName "modInventoryViewer.RunInventoryViewerFilterForTest" `
        -Arguments @("SKU-SHIP"))
    $eventsReport = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $operationsName `
        -MacroName "modInventoryViewer.RunInventoryViewerEventsForTest")

    $step = "publish a newer event and refresh the open Events page"
    $refreshInventoryWb = $excel.Workbooks.Open($inventoryPath)
    $opened.Add($refreshInventoryWb) | Out-Null
    $refreshInventoryLog = $refreshInventoryWb.Worksheets.Item("InventoryLog").ListObjects.Item("tblInventoryLog")
    Add-ListObjectRow -ListObject $refreshInventoryLog -Values @{
        "EventID" = "EVT-VIEWER-NEW-RECEIVE"
        "AppliedSeq" = 6
        "EventType" = "RECEIVE"
        "OccurredAtUTC" = [datetime]::UtcNow
        "AppliedAtUTC" = [datetime]::UtcNow
        "WarehouseId" = $warehouseId
        "StationId" = $stationId
        "UserId" = $testUser
        "System_Key" = "SYS-VIEWER-NEW-RECEIVE"
        "SKU" = "SKU-COMP"
        "QtyDelta" = 7
        "Location" = "DOCK"
        "Condition" = "GOOD"
        "AttributesJson" = "{}"
        "Note" = "Reference=BOL-VIEWER-NEW;Item=New Viewer Receipt;UOM=each"
    }
    $refreshInventoryWb.Save()
    $newSnapshotCreated = [bool](Run-WorkbookMacro -Excel $excel -WorkbookName $coreName `
        -MacroName "modWarehouseSync.GenerateWarehouseSnapshot" `
        -Arguments @($warehouseId, $refreshInventoryWb, $snapshotPath))
    if (-not $newSnapshotCreated) { throw "The newer isolated event snapshot was not created." }
    $refreshInventoryWb.Close($true)
    $snapshotHashPublished = (Get-FileHash -LiteralPath $snapshotPath -Algorithm SHA256).Hash
    $refreshedEventsReport = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $operationsName `
        -MacroName "modInventoryViewer.RunInventoryViewerEventsForTest")
    [void](Run-WorkbookMacro -Excel $excel -WorkbookName $operationsName `
        -MacroName "modInventoryViewer.CloseInventoryViewerForTest")
    $snapshotHashAfter = (Get-FileHash -LiteralPath $snapshotPath -Algorithm SHA256).Hash

    $firstOk = $firstReport -match '^OK\|' -and $firstReport -match '(?:^|\|)VisibleRows=3(?:\||$)' -and
        $firstReport -match '(?:^|\|)Generation=1(?:\||$)'
    $secondReused = $secondReport -match '^OK\|' -and
        $secondReport -match '(?:^|\|)VisibleRows=3(?:\||$)' -and
        $secondReport -match '(?:^|\|)Generation=1(?:\||$)'
    $filterOk = $filterReport -match '^OK\|' -and
        $filterReport -match '(?:^|\|)VisibleRows=1(?:\||$)' -and
        $filterReport -match '(?:^|\|)Generation=1(?:\||$)'
    $snapshotAdvanced = $snapshotHashBefore -ne $snapshotHashPublished
    $snapshotUnchanged = $snapshotHashPublished -eq $snapshotHashAfter
    $eventsOk = $eventsReport -match '^OK\|' -and
        $eventsReport -match '(?:^|\|)TabCount=2(?:\||$)' -and
        $eventsReport -match '(?:^|\|)TabCaptions=Inventory,Events(?:\||$)' -and
        $eventsReport -match '(?:^|\|)SelectedTab=Events(?:\||$)' -and
        $eventsReport -match '(?:^|\|)Title=Inventory and shipping events(?:\||$)' -and
        $eventsReport -match '(?:^|\|)VisibleRows=2(?:\||$)' -and
        $eventsReport -match '(?:^|\|)ReadableDates=2(?:\||$)' -and
        $eventsReport -match '(?:^|\|)FirstReference=SHIP-REMOVE-VIEWER(?:\||$)' -and
        $eventsReport -match '(?:^|\|)RemoveRows=1(?:\||$)' -and
        $eventsReport -match '(?:^|\|)Columns=10(?:\||$)' -and
        $eventsReport -match '(?:^|\|)ReadOnly=True(?:\||$)'
    $refreshedEventsOk = $refreshedEventsReport -match '^OK\|' -and
        $refreshedEventsReport -match '(?:^|\|)VisibleRows=3(?:\||$)' -and
        $refreshedEventsReport -match '(?:^|\|)ReadableDates=3(?:\||$)' -and
        $refreshedEventsReport -match '(?:^|\|)FirstReference=BOL-VIEWER-NEW(?:\||$)'

    $facts.ConfigLoaded = $configLoaded
    $facts.AuthLoaded = $authLoaded
    $facts.TargetSelected = $targetResult.StartsWith("OK|")
    $facts.TargetPathsSet = $targetPathsSet
    $facts.SignedIn = $signInResult.StartsWith("OK|")
    $facts.SnapshotCreated = $snapshotCreated
    $facts.FirstActionRows = if ($firstOk) { 3 } else { 0 }
    $facts.RepeatedLaunchReusedGeneration = $secondReused
    $facts.FilterVisibleRows = if ($filterOk) { 1 } else { 0 }
    $facts.EventsVisibleRows = if ($eventsOk) { 2 } else { 0 }
    $facts.RefreshedEventsVisibleRows = if ($refreshedEventsOk) { 3 } else { 0 }
    $facts.NewestPublishedReference = if ($refreshedEventsOk) { "BOL-VIEWER-NEW" } else { "Unexpected" }
    $facts.ReadableEventDates = $eventsOk -and $refreshedEventsOk
    $facts.ViewerTabCount = if ($eventsOk) { 2 } else { 0 }
    $facts.ViewerTabCaptions = if ($eventsOk) { "Inventory,Events" } else { "Unexpected" }
    $facts.SelectedViewerTab = if ($eventsOk) { "Events" } else { "Unexpected" }
    $facts.RemoveEventsVisible = $eventsOk
    $facts.EventsReadOnly = $eventsOk
    $facts.SnapshotHashUnchanged = $snapshotUnchanged
    $facts.NewPublicationChangedSnapshot = $snapshotAdvanced

    $passed = $configLoaded -and $authLoaded -and
        $targetResult.StartsWith("OK|") -and $targetPathsSet -and
        $signInResult.StartsWith("OK|") -and $snapshotCreated -and
        $firstOk -and $secondReused -and $filterOk -and $eventsOk -and
        $refreshedEventsOk -and $snapshotAdvanced -and $snapshotUnchanged
    $detail = if ($passed) {
        "The public Operations Viewer action loaded readable Receipt and Shipping Remove events, refreshed the already-open Events page to show a newly published receipt first, reused one form generation, kept Events read-only, and left the new snapshot byte-for-byte unchanged."
    } else {
        "The packaged Viewer contract failed at step '$step'."
    }
}
catch {
    $detail = "Harness exception at step '$step': $($_.Exception.Message)"
    $facts.HarnessException = $_.Exception.Message
}
finally {
    $lines = @(
        "# Inventory Viewer Packaged Results", "",
        "- Status: **$(if ($passed) { 'PASS' } else { 'FAIL' })**",
        "- Runtime: isolated generated test warehouse"
    )
    foreach ($entry in $facts.GetEnumerator()) {
        $lines += "- $($entry.Key): $($entry.Value)"
    }
    $lines += ""
    $lines += "## Observed result"
    $lines += ""
    $lines += $detail
    [IO.File]::WriteAllText($resultPath, (($lines -join "`n") + "`n"), (New-Object Text.UTF8Encoding($false)))

    foreach ($wb in $opened) {
        try { $wb.Close($false) } catch {}
        Release-ComObject $wb
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        Release-ComObject $excel
    }
    $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
    $resolvedRuntime = [IO.Path]::GetFullPath($runtimeRoot)
    if ($resolvedRuntime.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase) -and
        (Split-Path -Leaf $resolvedRuntime) -like "invsys-inventory-viewer-*") {
        Remove-Item -LiteralPath $resolvedRuntime -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Write-Host $detail
Write-Host "Evidence: $resultPath"
if (-not $passed) { exit 1 }
