[CmdletBinding()]
param(
    [string]$RepoRoot = ".",
    [string]$DeployRoot = "deploy/current",
    [string]$OutputDirectory = "reports/runtime/plan022-slice0",
    [ValidateSet("", "Receiving", "Production", "Shipping")]
    [string]$CallbackFilter = "",
    [ValidateSet("NoEligible", "ConfigActive", "SavedEligible", "UnrelatedActive", "CapturedClosed", "ReceivingDurability", "ReceivingFormClosed", "ShippingLayout", "ProductionDesignRed")]
    [string]$WorkbookState = "NoEligible"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Import-LiveValidationHelpers {
    param([string]$ScriptPath)

    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile(
        $ScriptPath,
        [ref]$tokens,
        [ref]$errors
    )
    if ($errors.Count -gt 0) {
        throw "Unable to parse validation helper source: $($errors[0].Message)"
    }

    $definitions = $ast.FindAll({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst]
    }, $true)
    foreach ($definition in $definitions) {
        $scriptScopedDefinition = $definition.Extent.Text -replace (
            '^(?i)function\s+' + [regex]::Escape($definition.Name)
        ), ('function script:' + $definition.Name)
        . ([scriptblock]::Create($scriptScopedDefinition))
    }

    $required = @(
        "Release-ComObject",
        "Get-InvSysCredentialHash",
        "Run-WorkbookMacro",
        "New-ConfigWorkbook",
        "New-AuthWorkbook",
        "New-InventoryWorkbook",
        "New-OperationalWorkbook",
        "Get-WorksheetSafe",
        "Get-ListObjectSafe",
        "Get-ColumnIndexSafe",
        "Get-RowCountSafe"
    )
    $missing = @($required | Where-Object { -not (Get-Command $_ -CommandType Function -ErrorAction SilentlyContinue) })
    if ($missing.Count -gt 0) {
        throw "Missing imported validation helpers: $($missing -join ', ')"
    }
}

function Get-ExcelProcessId {
    param([object]$Excel)

    [uint32]$processId = 0
    [void][Plan022LauncherWindow]::GetWindowThreadProcessId(
        [intptr]$Excel.Hwnd,
        [ref]$processId
    )
    return [int]$processId
}

function Get-OpenWorkbookNames {
    param([object]$Excel)

    $names = New-Object System.Collections.Generic.List[string]
    $count = [int]$Excel.Workbooks.Count
    for ($i = 1; $i -le $count; $i++) {
        try {
            $name = [string]$Excel.Workbooks.Item($i).Name
            if (-not [string]::IsNullOrWhiteSpace($name)) {
                $names.Add($name) | Out-Null
            }
        }
        catch {}
    }
    return @($names)
}

function Start-DialogCaptureAndDismiss {
    param(
        [int]$ExcelProcessId,
        [int]$TimeoutSeconds = 20
    )

    return Start-Job -ScriptBlock {
        param($processId, $timeoutSeconds)

        Add-Type -AssemblyName UIAutomationClient
        Add-Type -AssemblyName UIAutomationTypes
        Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class Plan022DialogDismiss
{
    [DllImport("user32.dll")]
    public static extern IntPtr SendMessage(
        IntPtr hWnd,
        uint message,
        IntPtr wParam,
        IntPtr lParam
    );
}
"@
        $shell = New-Object -ComObject WScript.Shell
        $stopAt = (Get-Date).AddSeconds($timeoutSeconds)
        $seen = @{}
        try {
            while ((Get-Date) -lt $stopAt) {
                Start-Sleep -Milliseconds 200
                $desktop = [System.Windows.Automation.AutomationElement]::RootElement
                $processCondition = New-Object System.Windows.Automation.PropertyCondition(
                    [System.Windows.Automation.AutomationElement]::ProcessIdProperty,
                    $processId
                )
                $windows = $desktop.FindAll(
                    [System.Windows.Automation.TreeScope]::Children,
                    $processCondition
                )
                foreach ($window in $windows) {
                    $modalCondition = New-Object System.Windows.Automation.PropertyCondition(
                        [System.Windows.Automation.AutomationElement]::ClassNameProperty,
                        "#32770"
                    )
                    $modalWindows = $window.FindAll(
                        [System.Windows.Automation.TreeScope]::Descendants,
                        $modalCondition
                    )
                    foreach ($modalWindow in $modalWindows) {
                        $modalName = [string]$modalWindow.Current.Name
                        $modalKey = "WINDOW|" + $modalName + "|#32770"
                        if (-not $seen.ContainsKey($modalKey)) {
                            $seen[$modalKey] = $true
                            Write-Output $modalKey
                        }
                        $modalElements = $modalWindow.FindAll(
                            [System.Windows.Automation.TreeScope]::Descendants,
                            [System.Windows.Automation.Condition]::TrueCondition
                        )
                        $modalOkHandle = [IntPtr]::Zero
                        foreach ($modalElement in $modalElements) {
                            $modalText = [string]$modalElement.Current.Name
                            if ([string]::IsNullOrWhiteSpace($modalText)) { continue }
                            $modalType = [string]$modalElement.Current.ControlType.ProgrammaticName
                            $modalElementKey = $modalName + "|" + $modalType + "|" + $modalText
                            if (-not $seen.ContainsKey($modalElementKey)) {
                                $seen[$modalElementKey] = $true
                                Write-Output ("WINDOW_ELEMENT|" + $modalName + "|" + $modalType + "|" + $modalText)
                            }
                            if ($modalText -eq "OK") {
                                try {
                                    $nativeHandle = [IntPtr]$modalElement.Current.NativeWindowHandle
                                    if ($nativeHandle -ne [IntPtr]::Zero) {
                                        $modalOkHandle = $nativeHandle
                                    }
                                }
                                catch {}
                            }
                        }
                        if ($modalOkHandle -ne [IntPtr]::Zero) {
                            try {
                                [void][Plan022DialogDismiss]::SendMessage(
                                    $modalOkHandle,
                                    0x00F5,
                                    [IntPtr]::Zero,
                                    [IntPtr]::Zero
                                )
                            }
                            catch {}
                        }
                        try {
                            if ($shell.AppActivate($processId)) {
                                $shell.SendKeys("~")
                            }
                        }
                        catch {}
                    }

                    $windowName = [string]$window.Current.Name
                    $windowClass = [string]$window.Current.ClassName
                    $isDialogWindow = $windowClass -ne "XLMAIN"
                    $hasDismissButton = $false
                    if ($isDialogWindow) {
                        $windowKey = "WINDOW|" + $windowName + "|" + $windowClass
                        if (-not $seen.ContainsKey($windowKey)) {
                            $seen[$windowKey] = $true
                            Write-Output $windowKey
                        }
                    }
                    $elements = $window.FindAll(
                        [System.Windows.Automation.TreeScope]::Descendants,
                        [System.Windows.Automation.Condition]::TrueCondition
                    )
                    foreach ($element in $elements) {
                        if (-not $isDialogWindow) { continue }
                        $textValue = [string]$element.Current.Name
                        if ([string]::IsNullOrWhiteSpace($textValue)) { continue }
                        $controlType = [string]$element.Current.ControlType.ProgrammaticName
                        $key = $windowName + "|" + $controlType + "|" + $textValue
                        if (-not $seen.ContainsKey($key)) {
                            $seen[$key] = $true
                            Write-Output ("WINDOW_ELEMENT|" + $windowName + "|" + $controlType + "|" + $textValue)
                        }
                    }

                    $buttonCondition = New-Object System.Windows.Automation.PropertyCondition(
                        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
                        [System.Windows.Automation.ControlType]::Button
                    )
                    $buttons = $window.FindAll(
                        [System.Windows.Automation.TreeScope]::Descendants,
                        $buttonCondition
                    )
                    foreach ($button in $buttons) {
                        if (-not $isDialogWindow -or [string]$button.Current.Name -ne "OK") { continue }
                        $hasDismissButton = $true
                        try {
                            $pattern = $button.GetCurrentPattern(
                                [System.Windows.Automation.InvokePattern]::Pattern
                            )
                            $pattern.Invoke()
                        }
                        catch {}
                    }
                    if ($hasDismissButton) {
                        try {
                            if ($shell.AppActivate($processId)) {
                                $shell.SendKeys("~")
                            }
                        }
                        catch {}
                    }
                }
            }
        }
        finally {
            if ($null -ne $shell) {
                try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($shell) } catch {}
            }
        }
    } -ArgumentList $ExcelProcessId, $TimeoutSeconds
}

function Invoke-PackagedCallback {
    param(
        [object]$Excel,
        [string]$WorkbookName,
        [string]$MacroName
    )

    $job = Start-DialogCaptureAndDismiss -ExcelProcessId (Get-ExcelProcessId $Excel)
    $errorText = ""
    try {
        [void](Run-WorkbookMacro -Excel $Excel -WorkbookName $WorkbookName -MacroName $MacroName)
    }
    catch {
        $errorText = $_.Exception.Message
    }
    finally {
        Start-Sleep -Milliseconds 600
        Stop-Job -Job $job -ErrorAction SilentlyContinue
        Wait-Job -Job $job -Timeout 2 -ErrorAction SilentlyContinue | Out-Null
    }

    $captured = @(
        Receive-Job -Job $job -ErrorAction SilentlyContinue |
            ForEach-Object { [string]$_ } |
            Sort-Object -Unique
    )
    Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
    return [pscustomobject]@{
        Macro = $MacroName
        Error = $errorText
        WindowText = $captured
    }
}

function Add-Evidence {
    param(
        [System.Collections.Generic.List[object]]$Rows,
        [string]$Callback,
        [string]$Expected,
        [bool]$Passed,
        [string]$Observed
    )

    $Rows.Add([pscustomobject]@{
        Callback = $Callback
        Expected = $Expected
        Passed = $Passed
        Observed = $Observed
    }) | Out-Null
}

function Get-WorkbookSurfaceCounts {
    param([object]$Workbook)

    if ($null -eq $Workbook) {
        return [pscustomobject]@{ Worksheets = 0; Tables = 0 }
    }
    $tableCount = 0
    foreach ($worksheet in $Workbook.Worksheets) {
        $tableCount += [int]$worksheet.ListObjects.Count
    }
    return [pscustomobject]@{
        Worksheets = [int]$Workbook.Worksheets.Count
        Tables = $tableCount
    }
}

function Get-FileHashMap {
    param([string[]]$Paths)

    $result = @{}
    foreach ($path in $Paths) {
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
            $result[$path] = "<missing>"
            continue
        }
        $result[$path] = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
    }
    return $result
}

function Compare-FileHashMaps {
    param(
        [hashtable]$Before,
        [hashtable]$After
    )

    foreach ($path in $Before.Keys) {
        if (-not $After.ContainsKey($path) -or $Before[$path] -ne $After[$path]) {
            return $false
        }
    }
    return $true
}

function Get-ChangedFileNames {
    param(
        [hashtable]$Before,
        [hashtable]$After
    )

    $changed = @()
    foreach ($path in $Before.Keys) {
        if (-not $After.ContainsKey($path) -or $Before[$path] -ne $After[$path]) {
            $changed += [IO.Path]::GetFileName($path)
        }
    }
    return @($changed)
}

function Get-RoleFormWindowCount {
    param([int]$ExcelProcessId)

    Add-Type -AssemblyName UIAutomationClient
    Add-Type -AssemblyName UIAutomationTypes
    $processCondition = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ProcessIdProperty,
        $ExcelProcessId
    )
    $classCondition = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ClassNameProperty,
        "ThunderDFrame"
    )
    $condition = New-Object System.Windows.Automation.AndCondition(
        $processCondition,
        $classCondition
    )
    $windows = [System.Windows.Automation.AutomationElement]::RootElement.FindAll(
        [System.Windows.Automation.TreeScope]::Descendants,
        $condition
    )
    return [int]$windows.Count
}

function Close-RoleFormWindow {
    param(
        [int]$ExcelProcessId,
        [string]$Caption
    )

    Add-Type -AssemblyName UIAutomationClient
    Add-Type -AssemblyName UIAutomationTypes
    $processCondition = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ProcessIdProperty,
        $ExcelProcessId
    )
    $classCondition = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ClassNameProperty,
        "ThunderDFrame"
    )
    $nameCondition = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::NameProperty,
        $Caption
    )
    $condition = New-Object System.Windows.Automation.AndCondition(
        $processCondition,
        $classCondition,
        $nameCondition
    )
    $windows = [System.Windows.Automation.AutomationElement]::RootElement.FindAll(
        [System.Windows.Automation.TreeScope]::Descendants,
        $condition
    )
    if ($windows.Count -ne 1) {
        return $false
    }
    $nativeHandle = [IntPtr]$windows.Item(0).Current.NativeWindowHandle
    if ($nativeHandle -eq [IntPtr]::Zero) {
        return $false
    }
    [void][Plan022LauncherWindow]::SendMessage(
        $nativeHandle,
        0x0010,
        [IntPtr]::Zero,
        [IntPtr]::Zero
    )
    Start-Sleep -Milliseconds 750
    return $true
}

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class Plan022LauncherWindow
{
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
    [DllImport("user32.dll")]
    public static extern IntPtr SendMessage(
        IntPtr hWnd,
        uint message,
        IntPtr wParam,
        IntPtr lParam
    );
}
"@

$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$deployPath = (Resolve-Path -LiteralPath (Join-Path $repo $DeployRoot)).Path
$helperSource = Join-Path $repo "tools/validate_phase6_live_role_workflows.ps1"
Import-LiveValidationHelpers -ScriptPath $helperSource

$outputPath = Join-Path $repo $OutputDirectory
New-Item -ItemType Directory -Path $outputPath -Force | Out-Null
$progressPath = Join-Path $outputPath "progress.txt"
$runtimeRoot = Join-Path ([IO.Path]::GetTempPath()) (
    "invsys-plan022-launcher-red-" + [guid]::NewGuid().ToString("N")
)
$warehouseId = "WHL" + [guid]::NewGuid().ToString("N").Substring(0, 6).ToUpperInvariant()
$stationId = "S1"
$testUser = if ([string]::IsNullOrWhiteSpace($env:USERNAME)) { "user1" } else { $env:USERNAME }
$testPin = [guid]::NewGuid().ToString("N")
$testPinHash = Get-InvSysCredentialHash -Credential $testPin
$configPath = Join-Path $runtimeRoot ($warehouseId + ".invSys.Config.xlsb")
$authPath = Join-Path $runtimeRoot ($warehouseId + ".invSys.Auth.xlsb")
$inventoryPath = Join-Path $runtimeRoot ($warehouseId + ".invSys.Data.Inventory.xlsb")
$snapshotPath = Join-Path $runtimeRoot ($warehouseId + ".invSys.Snapshot.Inventory.xlsb")
$operatorRoot = Join-Path $runtimeRoot "operator-workbooks"
$callbacks = @(
    @{
        Name = "Receiving"
        Macro = "modTS_Received.ShowReceivingForm"
        Expected = "Creates or opens one station-local saved Receiving operator workbook and opens the modeless form."
    },
    @{
        Name = "Production"
        Macro = "mProduction.BtnOpenProductionForm"
        Expected = "Creates or opens one station-local saved Production operator workbook and opens one modeless form."
    },
    @{
        Name = "Shipping"
        Macro = "modTS_Shipments.BtnOpenShipmentsForm"
        Expected = "Creates or opens one station-local saved Shipping operator workbook and opens one modeless form."
    }
)
if (-not [string]::IsNullOrWhiteSpace($CallbackFilter)) {
    $callbacks = @($callbacks | Where-Object { $_.Name -eq $CallbackFilter })
}
if ($WorkbookState -eq "CapturedClosed" -and [string]::IsNullOrWhiteSpace($CallbackFilter)) {
    throw "CapturedClosed requires -CallbackFilter so each role starts from its own captured workbook."
}
if ($WorkbookState -eq "ReceivingDurability") {
    if ([string]::IsNullOrWhiteSpace($CallbackFilter)) {
        $callbacks = @($callbacks | Where-Object { $_.Name -eq "Receiving" })
    }
    elseif ($CallbackFilter -ne "Receiving") {
        throw "ReceivingDurability supports only -CallbackFilter Receiving."
    }
}
if ($WorkbookState -eq "ReceivingFormClosed") {
    if ([string]::IsNullOrWhiteSpace($CallbackFilter)) {
        $callbacks = @($callbacks | Where-Object { $_.Name -eq "Receiving" })
    }
    elseif ($CallbackFilter -ne "Receiving") {
        throw "ReceivingFormClosed supports only -CallbackFilter Receiving."
    }
}
if ($WorkbookState -eq "ShippingLayout") {
    if ([string]::IsNullOrWhiteSpace($CallbackFilter)) {
        $callbacks = @($callbacks | Where-Object { $_.Name -eq "Shipping" })
    }
    elseif ($CallbackFilter -ne "Shipping") {
        throw "ShippingLayout supports only -CallbackFilter Shipping."
    }
}
if ($WorkbookState -eq "ProductionDesignRed") {
    if ([string]::IsNullOrWhiteSpace($CallbackFilter)) {
        $callbacks = @($callbacks | Where-Object { $_.Name -eq "Production" })
    }
    elseif ($CallbackFilter -ne "Production") {
        throw "ProductionDesignRed supports only -CallbackFilter Production."
    }
}
$packageNames = @(
    "invSys.Core.xlam",
    "invSys.Inventory.Domain.xlam",
    "invSys.Designs.Domain.xlam",
    "invSys.Operations.xlam"
)

$excel = $null
$excelProcessId = 0
$opened = New-Object System.Collections.Generic.List[object]
$packages = @{}
$evidence = New-Object System.Collections.Generic.List[object]
$setupEvidence = New-Object System.Collections.Generic.List[string]
$currentStep = "startup"
$stateWb = $null
$inventoryWb = $null
$canonicalPaths = @()
$canonicalHashesBefore = @{}

try {
    [IO.File]::WriteAllText($progressPath, "create runtime root")
    New-Item -ItemType Directory -Path $runtimeRoot -Force | Out-Null
    [IO.File]::WriteAllText($progressPath, "start Excel")
    $excel = New-Object -ComObject Excel.Application
    $excelProcessId = Get-ExcelProcessId $excel
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $true
    $excel.AutomationSecurity = 1

    $currentStep = "create isolated config and auth"
    [IO.File]::WriteAllText($progressPath, $currentStep)
    $configWb = New-ConfigWorkbook -Excel $excel -Path $configPath `
        -WarehouseId $warehouseId -StationId $stationId -RuntimeRoot $runtimeRoot
    $authWb = New-AuthWorkbook -Excel $excel -Path $authPath `
        -WarehouseId $warehouseId -StationId $stationId `
        -CurrentUserIds @($testUser) -CredentialHash $testPinHash
    $opened.Add($configWb) | Out-Null
    $opened.Add($authWb) | Out-Null
    if ($WorkbookState -eq "ReceivingDurability") {
        $inventoryWb = New-InventoryWorkbook -Excel $excel -Path $inventoryPath `
            -WarehouseId $warehouseId -SkuRows @("SKU-R1-DURABILITY")
        $opened.Add($inventoryWb) | Out-Null
    }

    $currentStep = "open packaged add-ins"
    [IO.File]::WriteAllText($progressPath, $currentStep)
    foreach ($packageName in $packageNames) {
        $packagePath = Join-Path $deployPath $packageName
        $packageWb = $excel.Workbooks.Open($packagePath)
        $opened.Add($packageWb) | Out-Null
        $packages[$packageName] = $packageWb
    }

    $coreName = [string]$packages["invSys.Core.xlam"].Name
    $operationsName = [string]$packages["invSys.Operations.xlam"].Name
    [IO.File]::WriteAllText($progressPath, "configure and sign in")
    [void](Run-WorkbookMacro -Excel $excel -WorkbookName $coreName `
        -MacroName "modRuntimeWorkbooks.SetCoreDataRootOverride" -Arguments @($runtimeRoot))
    $configLoaded = [bool](Run-WorkbookMacro -Excel $excel -WorkbookName $coreName `
        -MacroName "modConfig.LoadConfig" -Arguments @($warehouseId, $stationId))
    $authLoaded = [bool](Run-WorkbookMacro -Excel $excel -WorkbookName $coreName `
        -MacroName "modAuth.LoadAuth" -Arguments @($warehouseId))
    $targetResult = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $coreName `
        -MacroName "modNasConnection.SelectWarehouseTargetForAutomation" `
        -Arguments @($runtimeRoot, $runtimeRoot, $stationId, $true))
    $testHubRoot = "\\plan022-test\warehouse"
    $targetPathsSet = [bool](Run-WorkbookMacro -Excel $excel -WorkbookName $coreName `
        -MacroName "modNasConnection.SetCurrentTargetPathsForTest" `
        -Arguments @($testHubRoot, $runtimeRoot))
    $operatorRootOverrideSet = $false
    try {
        $operatorRootOverrideSet = [bool](Run-WorkbookMacro -Excel $excel `
            -WorkbookName $coreName `
            -MacroName "modWarehouseBootstrap.SetLocalOperatorRootOverrideForAutomation" `
            -Arguments @($operatorRoot))
    }
    catch {}
    $signInResult = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $coreName `
        -MacroName "modAuth.SignInCurrentTargetForAutomation" `
        -Arguments @($testUser, $testPin, "RECEIVE_POST"))
    $setupEvidence.Add("ConfigLoaded=$configLoaded") | Out-Null
    $setupEvidence.Add("AuthLoaded=$authLoaded") | Out-Null
    $setupEvidence.Add("TargetSelected=$($targetResult.StartsWith('OK|'))") | Out-Null
    $setupEvidence.Add("TestNasTargetApplied=$targetPathsSet") | Out-Null
    $setupEvidence.Add("OperatorRootOverrideSet=$operatorRootOverrideSet") | Out-Null
    $setupEvidence.Add("SignedIn=$($signInResult.StartsWith('OK|'))") | Out-Null

    if ($WorkbookState -eq "ReceivingDurability") {
        $currentStep = "generate isolated canonical snapshot"
        [IO.File]::WriteAllText($progressPath, $currentStep)
        $snapshotCreated = [bool](Run-WorkbookMacro -Excel $excel `
            -WorkbookName $coreName `
            -MacroName "modWarehouseSync.GenerateWarehouseSnapshot" `
            -Arguments @($warehouseId, $inventoryWb, $snapshotPath))
        if (-not $snapshotCreated) {
            throw "Unable to generate the isolated canonical snapshot."
        }
        $setupEvidence.Add("SnapshotCreated=True") | Out-Null
    }

    [void](Run-WorkbookMacro -Excel $excel -WorkbookName $operationsName `
        -MacroName "modOperationsInit.Auto_Open")

    $currentStep = "prepare workbook state"
    [IO.File]::WriteAllText($progressPath, $currentStep)
    if ($WorkbookState -in @("NoEligible", "ReceivingDurability", "ReceivingFormClosed", "ShippingLayout", "ProductionDesignRed")) {
        if ($null -ne $inventoryWb) {
            $inventoryWb.Close($true)
        }
        $authWb.Close($false)
        $configWb.Close($false)
    }
    elseif ($WorkbookState -eq "ConfigActive") {
        $authWb.Close($false)
        $configWb.Activate()
        $stateWb = $configWb
    }
    else {
        $authWb.Close($false)
        $configWb.Close($false)
        if ($WorkbookState -in @("SavedEligible", "CapturedClosed")) {
            $stateFileName = "Plan022.Role.Operator.xlsb"
        }
        else {
            $stateFileName = "Plan022.Unrelated.xlsb"
        }
        $statePath = Join-Path $runtimeRoot $stateFileName
        $stateWb = New-OperationalWorkbook -Excel $excel `
            -NameHint ([IO.Path]::GetFileName($statePath)) -Path $statePath
        $opened.Add($stateWb) | Out-Null
        if ($WorkbookState -in @("SavedEligible", "CapturedClosed")) {
            foreach ($surfaceMacro in @(
                "modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface",
                "modRoleWorkbookSurfaces.EnsureProductionWorkbookSurface",
                "modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface"
            )) {
                $surfaceOk = [bool](Run-WorkbookMacro -Excel $excel `
                    -WorkbookName $coreName -MacroName $surfaceMacro `
                    -Arguments @($stateWb))
                if (-not $surfaceOk) {
                    throw "Unable to create saved eligible role surface: $surfaceMacro"
                }
            }
            $stateWb.Save()
        }
        $stateWb.Activate()
    }
    if ($WorkbookState -eq "ReceivingDurability") {
        $canonicalPaths = @($configPath, $authPath, $inventoryPath, $snapshotPath)
        $canonicalHashesBefore = Get-FileHashMap -Paths $canonicalPaths
    }

    $currentStep = "invoke packaged callbacks for state $WorkbookState"
    [IO.File]::WriteAllText($progressPath, $currentStep)
    foreach ($callback in $callbacks) {
        [IO.File]::WriteAllText($progressPath, "invoke " + $callback.Macro)
        $beforeNames = @(Get-OpenWorkbookNames -Excel $excel)
        $beforeSurface = Get-WorkbookSurfaceCounts -Workbook $stateWb
        $capture = Invoke-PackagedCallback -Excel $excel `
            -WorkbookName $operationsName -MacroName $callback.Macro
        $afterNames = @(Get-OpenWorkbookNames -Excel $excel)
        $afterSurface = Get-WorkbookSurfaceCounts -Workbook $stateWb
        $observedText = (@($capture.WindowText) -join " || ")
        if (-not [string]::IsNullOrWhiteSpace($capture.Error)) {
            $observedText += " || COM_ERROR=" + $capture.Error
        }
        $newWorkbooks = @($afterNames | Where-Object { $_ -notin $beforeNames })
        $newWorkbookPaths = @()
        foreach ($newWorkbookName in $newWorkbooks) {
            try {
                $newWorkbookPaths += [string]$excel.Workbooks.Item($newWorkbookName).FullName
            }
            catch {}
        }
        if ($newWorkbooks.Count -gt 0) {
            $observedText += " || NEW_WORKBOOKS=" + ($newWorkbooks -join ",")
            $observedText += " || NEW_WORKBOOKS_STATION_LOCAL=" +
                [bool](@($newWorkbookPaths | Where-Object {
                    [IO.Path]::GetFullPath($_).StartsWith(
                        [IO.Path]::GetFullPath($operatorRoot),
                        [StringComparison]::OrdinalIgnoreCase
                    )
                }).Count -eq $newWorkbookPaths.Count)
        }
        $surfaceChanged = $beforeSurface.Worksheets -ne $afterSurface.Worksheets -or
            $beforeSurface.Tables -ne $afterSurface.Tables
        if ($null -ne $stateWb) {
            $observedText += " || SURFACE_BEFORE=$($beforeSurface.Worksheets)/$($beforeSurface.Tables)"
            $observedText += " || SURFACE_AFTER=$($afterSurface.Worksheets)/$($afterSurface.Tables)"
            $observedText += " || SURFACE_CHANGED=$surfaceChanged"
        }

        $passed = $false
        if ($WorkbookState -eq "ShippingLayout" -and $callback.Name -eq "Shipping") {
            $statusLayoutReport = [string](Run-WorkbookMacro -Excel $excel `
                -WorkbookName $operationsName `
                -MacroName "modTS_Shipments.RunShippingStatusAnchorTest")
            $boxingLayoutReport = [string](Run-WorkbookMacro -Excel $excel `
                -WorkbookName $operationsName `
                -MacroName "modTS_Shipments.RunShippingBoxingLayoutTest")
            $shippingIdentityReport = [string](Run-WorkbookMacro -Excel $excel `
                -WorkbookName $operationsName `
                -MacroName "modTS_Shipments.RunShippingSystemKeyIdentityTest")
            $observedText += " || STATUS_LAYOUT=" + $statusLayoutReport
            $observedText += " || BOXING_LAYOUT=" + $boxingLayoutReport
            $observedText += " || SHIPPING_IDENTITY=" + $shippingIdentityReport
            $passed = [string]::IsNullOrWhiteSpace($capture.Error) -and
                $statusLayoutReport -match '^OK\|' -and
                $statusLayoutReport -match '(?:^|\|)TopStable=True(?:\||$)' -and
                $statusLayoutReport -match '(?:^|\|)HeightStable=True(?:\||$)' -and
                $statusLayoutReport -match '(?:^|\|)AboveSearch=True(?:\||$)' -and
                $boxingLayoutReport -match '^OK\|' -and
                $boxingLayoutReport -match '(?:^|\|)BuilderInventoryGrew=True(?:\||$)' -and
                $boxingLayoutReport -match '(?:^|\|)MakerDesignsGrew=True(?:\||$)' -and
                $boxingLayoutReport -match '(?:^|\|)BoxingHeaderWidthsMatchLists=True(?:\||$)' -and
                $boxingLayoutReport -match '(?:^|\|)HeadersMatch=True(?:\||$)' -and
                $boxingLayoutReport -match '(?:^|\|)SearchFiltered=True(?:\||$)' -and
                $boxingLayoutReport -match '(?:^|\|)NonVersionIsNA=True(?:\||$)' -and
                $shippingIdentityReport -match '^OK\|' -and
                $shippingIdentityReport -match '(?:^|\|)ControlReady=True(?:\||$)' -and
                $shippingIdentityReport -match '(?:^|\|)ValuePreserved=True(?:\||$)' -and
                $shippingIdentityReport -match '(?:^|\|)ReservationUsesKey=True(?:\||$)'
        }
        elseif ($WorkbookState -eq "ReceivingDurability" -and $callback.Name -eq "Receiving") {
            $operatorPath = @($newWorkbookPaths | Where-Object {
                [IO.Path]::GetFullPath($_).StartsWith(
                    [IO.Path]::GetFullPath($operatorRoot),
                    [StringComparison]::OrdinalIgnoreCase
                )
            } | Select-Object -First 1)
            $operatorWb = if ($operatorPath.Count -eq 1) {
                $excel.Workbooks.Item([IO.Path]::GetFileName($operatorPath[0]))
            }
            else {
                $null
            }
            $customHeader = "Custom_R1_Launcher_Persistence"
            $customValue = "PRESERVE"
            $customAdded = $false
            $historyRowsAdded = 0
            $receivingHistoryReport = ""
            $receivingControlReport = ""
            if ($null -ne $operatorWb) {
                $inventorySheet = Get-WorksheetSafe -Workbook $operatorWb `
                    -WorksheetName "InventoryManagement"
                $inventoryTable = Get-ListObjectSafe -Worksheet $inventorySheet `
                    -TableName "invSys"
                if ($null -ne $inventoryTable) {
                    $customColumn = $inventoryTable.ListColumns.Add()
                    $customColumn.Name = $customHeader
                    if ((Get-RowCountSafe $inventoryTable) -gt 0) {
                        $customColumn.DataBodyRange.Cells.Item(1, 1).Value2 = $customValue
                    }
                    $operatorWb.Save()
                    $customAdded = (Get-ColumnIndexSafe -ListObject $inventoryTable -ColumnName $customHeader) -gt 0
                }
                $historySheet = Get-WorksheetSafe -Workbook $operatorWb `
                    -WorksheetName "ReceivedLog"
                $historyTable = Get-ListObjectSafe -Worksheet $historySheet `
                    -TableName "ReceivedLog"
                if ($null -ne $historyTable) {
                    foreach ($historyOrdinal in 1..2) {
                        $historyRow = $historyTable.ListRows.Add()
                        $historyValues = @{
                            "ENTRY_DATE" = [DateTime]::UtcNow.AddMinutes(-$historyOrdinal).ToOADate()
                            "USER" = $testUser
                            "REF_NUMBER" = "R1-HISTORY-$historyOrdinal"
                            "ITEMS" = "R1 durability item"
                            "QUANTITY" = $historyOrdinal
                            "UOM" = "EA"
                            "VENDOR" = "R1 test vendor"
                            "LOCATION" = "TEST"
                            "ITEM_CODE" = "SKU-R1-DURABILITY"
                            "System_Key" = "SYS-R1-HISTORY-$historyOrdinal"
                            "EventId" = "EVT-R1-HISTORY-$historyOrdinal"
                        }
                        foreach ($historyColumn in $historyValues.Keys) {
                            $historyColumnIndex = Get-ColumnIndexSafe `
                                -ListObject $historyTable -ColumnName $historyColumn
                            if ($historyColumnIndex -gt 0) {
                                $historyRow.Range.Cells.Item(1, $historyColumnIndex).Value2 = `
                                    $historyValues[$historyColumn]
                            }
                        }
                        $historyRowsAdded++
                    }
                    $operatorWb.Save()
                }
                $receivingHistoryReport = [string](Run-WorkbookMacro -Excel $excel `
                    -WorkbookName $operationsName `
                    -MacroName "modTS_Received.RunReceivingRefreshFormActionForTest" `
                    -Arguments @([string]$operatorWb.Name, ""))
                $receivingControlReport = [string](Run-WorkbookMacro -Excel $excel `
                    -WorkbookName $operationsName `
                    -MacroName "modTS_Received.RunReceivingSearchAndHeaderContractTest")
            }
            $formsBeforeClose = Get-RoleFormWindowCount -ExcelProcessId $excelProcessId
            if ($null -ne $operatorWb) {
                $operatorWb.Close($true)
            }

            $unrelatedPath = Join-Path $runtimeRoot "Plan022.Durability.Unrelated.xlsb"
            $unrelatedWb = New-OperationalWorkbook -Excel $excel `
                -NameHint ([IO.Path]::GetFileName($unrelatedPath)) -Path $unrelatedPath
            $opened.Add($unrelatedWb) | Out-Null
            $unrelatedBefore = Get-WorkbookSurfaceCounts -Workbook $unrelatedWb
            $unrelatedWb.Activate()

            $reopenNamesBefore = @(Get-OpenWorkbookNames -Excel $excel)
            $reopenCapture = Invoke-PackagedCallback -Excel $excel `
                -WorkbookName $operationsName -MacroName $callback.Macro
            $reopenNamesAfter = @(Get-OpenWorkbookNames -Excel $excel)
            $reopenNewNames = @($reopenNamesAfter | Where-Object { $_ -notin $reopenNamesBefore })
            $operatorWb = if ($operatorPath.Count -eq 1) {
                try { $excel.Workbooks.Item([IO.Path]::GetFileName($operatorPath[0])) } catch { $null }
            }
            else {
                $null
            }
            $refreshOk = $false
            $customPreserved = $false
            if ($null -ne $operatorWb) {
                $refreshOk = [bool](Run-WorkbookMacro -Excel $excel `
                    -WorkbookName $coreName `
                    -MacroName "modOperationsPrimitiveBridge.RefreshInventoryReadModel" `
                    -Arguments @([string]$operatorWb.Name, $warehouseId, "LOCAL"))
                $inventorySheet = Get-WorksheetSafe -Workbook $operatorWb `
                    -WorksheetName "InventoryManagement"
                $inventoryTable = Get-ListObjectSafe -Worksheet $inventorySheet `
                    -TableName "invSys"
                $customIndex = Get-ColumnIndexSafe -ListObject $inventoryTable -ColumnName $customHeader
                $customPreserved = $customIndex -gt 0
                if ($customPreserved -and (Get-RowCountSafe $inventoryTable) -gt 0) {
                    $customPreserved = (
                        [string]$inventoryTable.DataBodyRange.Cells.Item(1, $customIndex).Value2
                    ) -eq $customValue
                }
            }

            $formCountAfterReopen = Get-RoleFormWindowCount -ExcelProcessId $excelProcessId
            $unrelatedWb.Activate()
            Start-Sleep -Milliseconds 300
            $formCountAfterUnrelatedActivation = Get-RoleFormWindowCount `
                -ExcelProcessId $excelProcessId
            $unrelatedAfter = Get-WorkbookSurfaceCounts -Workbook $unrelatedWb
            $unrelatedChanged = $unrelatedBefore.Worksheets -ne $unrelatedAfter.Worksheets -or
                $unrelatedBefore.Tables -ne $unrelatedAfter.Tables
            $canonicalHashesAfter = Get-FileHashMap -Paths $canonicalPaths
            $canonicalHashesUnchanged = Compare-FileHashMaps `
                -Before $canonicalHashesBefore -After $canonicalHashesAfter
            $changedCanonicalNames = @(
                Get-ChangedFileNames `
                    -Before $canonicalHashesBefore -After $canonicalHashesAfter
            )
            $operatorFileCount = @(
                Get-ChildItem -LiteralPath $operatorRoot -Filter "*.Receiving.Operator.xlsm" `
                    -File -Recurse -ErrorAction SilentlyContinue
            ).Count
            $reopenText = @($reopenCapture.WindowText) -join " // "
            $observedText += " || CUSTOM_ADDED=$customAdded"
            $observedText += " || RECEIVING_HISTORY_ROWS_ADDED=$historyRowsAdded"
            $observedText += " || RECEIVING_HISTORY=" + $receivingHistoryReport
            $observedText += " || RECEIVING_CONTROLS=" + $receivingControlReport
            $observedText += " || REFRESH_OK=$refreshOk"
            $observedText += " || CUSTOM_PRESERVED=$customPreserved"
            $observedText += " || CANONICAL_HASHES_UNCHANGED=$canonicalHashesUnchanged"
            if ($changedCanonicalNames.Count -gt 0) {
                $observedText += " || CHANGED_CANONICAL_FILES=" +
                    ($changedCanonicalNames -join ",")
            }
            $observedText += " || FORMS_BEFORE_CLOSE=$formsBeforeClose"
            $observedText += " || FORMS_AFTER_REOPEN=$formCountAfterReopen"
            $observedText += " || FORMS_AFTER_UNRELATED_ACTIVATION=$formCountAfterUnrelatedActivation"
            $observedText += " || UNRELATED_SURFACE_CHANGED=$unrelatedChanged"
            $observedText += " || OPERATOR_FILES=$operatorFileCount"
            $observedText += " || REOPEN_NEW_WORKBOOKS=" + ($reopenNewNames -join ",")
            if (-not [string]::IsNullOrWhiteSpace($reopenText)) {
                $observedText += " || REOPEN_DIALOGS=" + $reopenText
            }
            $passed = $operatorRootOverrideSet -and
                $newWorkbooks.Count -eq 1 -and
                $customAdded -and
                $historyRowsAdded -eq 2 -and
                $receivingHistoryReport -match '^OK\|' -and
                $receivingHistoryReport -match '(?:^|\|)HistoryRows=2(?:\||$)' -and
                $receivingHistoryReport -match '(?:^|\|)LoaderHistoryRowsBefore=2(?:\||$)' -and
                $receivingHistoryReport -match '(?:^|\|)LoaderHistoryRowsAfter=2(?:\||$)' -and
                $receivingHistoryReport -match '(?:^|\|)DirectHistoryRows=2(?:\||$)' -and
                $receivingControlReport -match '^OK\|' -and
                $receivingControlReport -match '(?:^|\|)DedicatedItemResults=True(?:\||$)' -and
                $receivingControlReport -match '(?:^|\|)Location=True(?:\||$)' -and
                $receivingControlReport -match '(?:^|\|)OptionalLot=True(?:\||$)' -and
                $receivingControlReport -match '(?:^|\|)ReceivingHeaderColumnsAligned=True(?:\||$)' -and
                $refreshOk -and
                $customPreserved -and
                $canonicalHashesUnchanged -and
                $formsBeforeClose -eq 1 -and
                $formCountAfterReopen -eq 1 -and
                $formCountAfterUnrelatedActivation -eq 1 -and
                -not $unrelatedChanged -and
                $operatorFileCount -eq 1 -and
                $reopenNewNames.Count -eq 1 -and
                $reopenText -notmatch "(?i)(failed|Type mismatch|operator workbook)"
        }
        elseif ($WorkbookState -eq "ReceivingFormClosed") {
            $formCountBeforeClose = Get-RoleFormWindowCount `
                -ExcelProcessId $excelProcessId
            $formClosed = Close-RoleFormWindow -ExcelProcessId $excelProcessId `
                -Caption "Receiving"
            $formCountAfterClose = Get-RoleFormWindowCount `
                -ExcelProcessId $excelProcessId
            $secondBeforeNames = @(Get-OpenWorkbookNames -Excel $excel)
            $secondCapture = Invoke-PackagedCallback -Excel $excel `
                -WorkbookName $operationsName -MacroName $callback.Macro
            $automationDisconnectedAfterRelaunch = $false
            try {
                $secondAfterNames = @(Get-OpenWorkbookNames -Excel $excel)
            }
            catch {
                $automationDisconnectedAfterRelaunch = $true
                $secondAfterNames = @($secondBeforeNames)
            }
            $secondNewWorkbooks = @(
                $secondAfterNames | Where-Object { $_ -notin $secondBeforeNames }
            )
            $formCountAfterRelaunch = Get-RoleFormWindowCount `
                -ExcelProcessId $excelProcessId
            $ownedExcelProcess = Get-Process -Id $excelProcessId -ErrorAction SilentlyContinue
            $excelAliveAfterRelaunch = $null -ne $ownedExcelProcess -and
                $ownedExcelProcess.ProcessName -eq "EXCEL" -and
                $ownedExcelProcess.Responding
            $operatorFileCount = @(
                Get-ChildItem -LiteralPath $operatorRoot `
                    -Filter "*.Receiving.Operator.xlsm" -File -Recurse `
                    -ErrorAction SilentlyContinue
            ).Count
            $secondText = @($secondCapture.WindowText) -join " // "
            $observedText += " || FORM_CLOSED=$formClosed"
            $observedText += " || FORMS_BEFORE_CLOSE=$formCountBeforeClose"
            $observedText += " || FORMS_AFTER_CLOSE=$formCountAfterClose"
            $observedText += " || FORMS_AFTER_RELAUNCH=$formCountAfterRelaunch"
            $observedText += " || RELAUNCH_NEW_WORKBOOKS=" + $secondNewWorkbooks.Count
            $observedText += " || OPERATOR_FILES=$operatorFileCount"
            $observedText += " || EXCEL_RESPONDING_AFTER_RELAUNCH=$excelAliveAfterRelaunch"
            $observedText += " || AUTOMATION_CLIENT_DISCONNECTED_AFTER_RELAUNCH=" +
                $automationDisconnectedAfterRelaunch
            if (-not [string]::IsNullOrWhiteSpace($secondCapture.Error)) {
                $observedText += " || RELAUNCH_COM_ERROR=" + $secondCapture.Error
            }
            if (-not [string]::IsNullOrWhiteSpace($secondText)) {
                $observedText += " || RELAUNCH_DIALOGS=" + $secondText
            }
            $passed = $newWorkbooks.Count -eq 1 -and
                $formCountBeforeClose -eq 1 -and
                $formClosed -and
                $formCountAfterClose -eq 0 -and
                $secondNewWorkbooks.Count -eq 0 -and
                $formCountAfterRelaunch -eq 1 -and
                $operatorFileCount -eq 1 -and
                $excelAliveAfterRelaunch -and
                [string]::IsNullOrWhiteSpace($capture.Error) -and
                [string]::IsNullOrWhiteSpace($secondCapture.Error) -and
                $secondText -notmatch "(?i)(failed|automation error)"
        }
        elseif ($WorkbookState -in @("NoEligible", "ProductionDesignRed")) {
            $workflowControlReport = ""
            $workflowControlPassed = $true
            $secondBeforeNames = @(Get-OpenWorkbookNames -Excel $excel)
            $secondCapture = Invoke-PackagedCallback -Excel $excel `
                -WorkbookName $operationsName -MacroName $callback.Macro
            $secondAfterNames = @(Get-OpenWorkbookNames -Excel $excel)
            $secondNewWorkbooks = @($secondAfterNames | Where-Object { $_ -notin $secondBeforeNames })
            $observedText += " || SECOND_LAUNCH_NEW_WORKBOOKS=" + $secondNewWorkbooks.Count
            if ($callback.Name -eq "Production") {
                $workflowControlReport = [string](Run-WorkbookMacro -Excel $excel `
                    -WorkbookName $operationsName `
                    -MacroName "mProduction.RunProductionBatchScaleContractTest")
                $workflowControlPassed = $workflowControlReport -match '^OK\|' -and
                    $workflowControlReport -match '(?:^|\|)Min=\.001%(?:\||$)' -and
                    $workflowControlReport -match '(?:^|\|)Default=100%(?:\||$)' -and
                    $workflowControlReport -match '(?:^|\|)Max=1000%(?:\||$)' -and
                    $workflowControlReport -match '(?:^|\|)BoundsRejected=True(?:\||$)'
                $observedText += " || PRODUCTION_BATCH_SCALE=" + $workflowControlReport
                if ($WorkbookState -eq "ProductionDesignRed") {
                    $productionDesignReport = [string](Run-WorkbookMacro -Excel $excel `
                        -WorkbookName $operationsName `
                        -MacroName "mProduction.RunReusableProductionSurfaceContractTest")
                    $productionDesignPassed = $productionDesignReport -match '^OK\|' -and
                        $productionDesignReport -match '(?:^|\|)Pages=5(?:\||$)' -and
                        $productionDesignReport -match '(?:^|\|)ProcessDesigner=True(?:\||$)' -and
                        $productionDesignReport -match '(?:^|\|)RecipeDesigner=True(?:\||$)' -and
                        $productionDesignReport -match '(?:^|\|)LegacyRecipeBuilder=False(?:\||$)'
                    $workflowControlPassed = $workflowControlPassed -and $productionDesignPassed
                    $observedText += " || PRODUCTION_REUSABLE_DESIGN=" + $productionDesignReport
                }
            }
            if (@($secondCapture.WindowText).Count -gt 0) {
                $observedText += " || SECOND_LAUNCH=" + (@($secondCapture.WindowText) -join " // ")
            }
            $stationLocalCount = @($newWorkbookPaths | Where-Object {
                [IO.Path]::GetFullPath($_).StartsWith(
                    [IO.Path]::GetFullPath($operatorRoot),
                    [StringComparison]::OrdinalIgnoreCase
                )
            }).Count
            $passed = $operatorRootOverrideSet -and
                $stationLocalCount -eq 1 -and
                $secondNewWorkbooks.Count -eq 0 -and
                $workflowControlPassed -and
                [string]::IsNullOrWhiteSpace($capture.Error) -and
                [string]::IsNullOrWhiteSpace($secondCapture.Error) -and
                $observedText -notmatch "(?i)(failed|Type mismatch|operator workbook)"
        }
        elseif ($WorkbookState -in @("ConfigActive", "UnrelatedActive")) {
            $stationLocalCount = @($newWorkbookPaths | Where-Object {
                [IO.Path]::GetFullPath($_).StartsWith(
                    [IO.Path]::GetFullPath($operatorRoot),
                    [StringComparison]::OrdinalIgnoreCase
                )
            }).Count
            $passed = $operatorRootOverrideSet -and
                -not $surfaceChanged -and
                $stationLocalCount -eq 1 -and
                $observedText -notmatch "(?i)(failed|Type mismatch|operator workbook)"
        }
        elseif ($WorkbookState -in @("SavedEligible", "CapturedClosed")) {
            $passed = [string]::IsNullOrWhiteSpace($capture.Error) -and
                $observedText -notmatch "(?i)(failed|Type mismatch|operator workbook)"
        }

        if ($WorkbookState -eq "CapturedClosed") {
            $initialLaunchPassed = $passed
            $stateWb.Close($false)
            $recoveryPath = Join-Path $runtimeRoot "Plan022.AfterCapturedClose.Unrelated.xlsb"
            $recoveryWb = New-OperationalWorkbook -Excel $excel `
                -NameHint ([IO.Path]::GetFileName($recoveryPath)) -Path $recoveryPath
            $opened.Add($recoveryWb) | Out-Null
            $recoveryWb.Activate()
            $recoveryBefore = Get-WorkbookSurfaceCounts -Workbook $recoveryWb
            $recoveryNamesBefore = @(Get-OpenWorkbookNames -Excel $excel)
            $recoveryCapture = Invoke-PackagedCallback -Excel $excel `
                -WorkbookName $operationsName -MacroName $callback.Macro
            $recoveryNamesAfter = @(Get-OpenWorkbookNames -Excel $excel)
            $recoveryAfter = Get-WorkbookSurfaceCounts -Workbook $recoveryWb
            $recoveryNewNames = @($recoveryNamesAfter | Where-Object { $_ -notin $recoveryNamesBefore })
            $recoveryText = (@($recoveryCapture.WindowText) -join " // ")
            $recoverySurfaceChanged = $recoveryBefore.Worksheets -ne $recoveryAfter.Worksheets -or
                $recoveryBefore.Tables -ne $recoveryAfter.Tables
            $observedText += " || AFTER_CAPTURED_CLOSE=" + $recoveryText
            $observedText += " || RECOVERY_SURFACE_CHANGED=" + $recoverySurfaceChanged
            $observedText += " || RECOVERY_NEW_WORKBOOKS=" + ($recoveryNewNames -join ",")

            $recoveryLocalCount = 0
            foreach ($recoveryName in $recoveryNewNames) {
                try {
                    $recoveryFullName = [string]$excel.Workbooks.Item($recoveryName).FullName
                    If ([IO.Path]::GetFullPath($recoveryFullName).StartsWith(
                            [IO.Path]::GetFullPath($operatorRoot),
                            [StringComparison]::OrdinalIgnoreCase)) {
                        $recoveryLocalCount += 1
                    }
                }
                catch {}
            }
            $passed = $initialLaunchPassed -and
                -not $recoverySurfaceChanged -and
                $recoveryLocalCount -eq 1 -and
                $recoveryText -notmatch "(?i)(failed|Type mismatch|operator workbook)"
        }
        Add-Evidence -Rows $evidence -Callback $callback.Macro `
            -Expected $callback.Expected -Passed $passed -Observed $observedText
        [IO.File]::WriteAllText($progressPath, "completed " + $callback.Macro)
    }
}
catch {
    Add-Evidence -Rows $evidence -Callback "HARNESS" `
        -Expected "The isolated packaged harness reaches the launcher callbacks." `
        -Passed $false -Observed (
            "Step=" + $currentStep + "; " + $_.Exception.Message +
            "; Location=" + ($_.InvocationInfo.PositionMessage -replace '\r?\n', ' ')
        )
}
finally {
    $reportLines = New-Object System.Collections.Generic.List[string]
    $reportTitle = if ($WorkbookState -eq "ReceivingDurability") {
        "# Plan 022 Slice 1 Receiving Durability Evidence"
    }
    elseif ($WorkbookState -eq "ShippingLayout") {
        "# Plan 022 Slices 4g-4h Shipping Layout Evidence"
    }
    elseif ($WorkbookState -eq "ProductionDesignRed") {
        "# Plan 022 Slice 4x Packaged Reusable Production RED Evidence"
    }
    else {
        "# Plan 022 Slice 0 Packaged Launcher Evidence"
    }
    $reportLines.Add($reportTitle) | Out-Null
    $reportLines.Add("") | Out-Null
    $reportLines.Add("- Captured: $([DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ'))") | Out-Null
    $reportLines.Add("- Runtime: isolated temporary test runtime (not NAS acceptance)") | Out-Null
    $reportLines.Add("- Package source: $DeployRoot") | Out-Null
    $reportLines.Add("- Workbook state: connected and signed in; $WorkbookState") | Out-Null
    foreach ($line in $setupEvidence) {
        $reportLines.Add("- $line") | Out-Null
    }
    $reportLines.Add("") | Out-Null
    $reportLines.Add("| Callback | Result | Expected | Observed |") | Out-Null
    $reportLines.Add("|---|---|---|---|") | Out-Null
    foreach ($row in $evidence) {
        $result = if ($row.Passed) { "PASS" } else { "RED" }
        $expected = ([string]$row.Expected).Replace("|", "/")
        $observed = ([string]$row.Observed).Replace("|", "/")
        $reportLines.Add("| $($row.Callback) | $result | $expected | $observed |") | Out-Null
    }
    $reportFile = if ($WorkbookState -eq "NoEligible") {
        "packaged-launcher-noeligible.md"
    }
    elseif ($WorkbookState -eq "ShippingLayout") {
        "shipping-layout.md"
    }
    elseif ($WorkbookState -eq "CapturedClosed") {
        "packaged-launcher-capturedclosed-$($CallbackFilter.ToLowerInvariant()).md"
    }
    elseif ($WorkbookState -eq "ReceivingDurability") {
        "receiving-launcher-durability.md"
    }
    elseif ($WorkbookState -eq "ReceivingFormClosed") {
        "receiving-launcher-formclosed.md"
    }
    elseif ($WorkbookState -eq "ProductionDesignRed") {
        "production-reusable-design-red.md"
    }
    else {
        "packaged-launcher-$($WorkbookState.ToLowerInvariant()).md"
    }
    $reportPath = Join-Path $outputPath $reportFile
    [IO.File]::WriteAllLines($reportPath, $reportLines)

    foreach ($wb in $opened) {
        try { $wb.Close($false) } catch {}
        Release-ComObject $wb
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        Release-ComObject $excel
    }
    if ($excelProcessId -gt 0) {
        Start-Sleep -Milliseconds 500
        $ownedExcelProcess = Get-Process -Id $excelProcessId -ErrorAction SilentlyContinue
        if ($null -ne $ownedExcelProcess -and $ownedExcelProcess.ProcessName -eq "EXCEL") {
            Stop-Process -Id $excelProcessId -Force
        }
    }

    $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
    $resolvedRuntime = [IO.Path]::GetFullPath($runtimeRoot)
    if ($resolvedRuntime.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase) -and
        (Split-Path -Leaf $resolvedRuntime) -like "invsys-plan022-launcher-red-*") {
        Remove-Item -LiteralPath $resolvedRuntime -Recurse -Force -ErrorAction SilentlyContinue
    }
}

$redCount = @($evidence | Where-Object { -not $_.Passed }).Count
if ($redCount -eq 0) {
    Write-Output "PLAN022_PACKAGED_LAUNCHER_GREEN"
}
else {
    Write-Output "PLAN022_PACKAGED_LAUNCHER_RED"
}
Write-Output "REPORT=$reportPath"
Write-Output "PASS=$($evidence.Count - $redCount) RED=$redCount TOTAL=$($evidence.Count)"
if ($redCount -eq 0) {
    exit 0
}
exit 1
