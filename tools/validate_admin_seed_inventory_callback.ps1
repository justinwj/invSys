[CmdletBinding()]
param(
    [string]$RepoRoot = ".",
    [string]$DeployRoot = "deploy/current",
    [ValidateSet("RED", "GREEN")]
    [string]$EvidencePhase = "RED",
    [switch]$KeepArtifacts
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$expectedDemoCount = 19

function Import-FunctionDefinitions {
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
}

function Get-ExcelProcessId {
    param([object]$Excel)

    [uint32]$processId = 0
    [void][InvSysAdminSeedWindow]::GetWindowThreadProcessId(
        [intptr]$Excel.Hwnd,
        [ref]$processId
    )
    return [int]$processId
}

function Start-SeedUiDriver {
    param(
        [int]$ExcelProcessId,
        [int]$TimeoutSeconds = 45
    )

    return Start-Job -ScriptBlock {
        param($processId, $timeoutSeconds)

        Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

public static class InvSysSeedNative
{
    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumChildWindows(IntPtr parent, EnumWindowsProc callback, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr hWnd, StringBuilder className, int maxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int maxCount);

    private static string WindowClass(IntPtr hWnd)
    {
        StringBuilder value = new StringBuilder(256);
        GetClassName(hWnd, value, value.Capacity);
        return value.ToString();
    }

    public static string WindowText(IntPtr hWnd)
    {
        StringBuilder value = new StringBuilder(2048);
        GetWindowText(hWnd, value, value.Capacity);
        return value.ToString();
    }

    public static IntPtr FindWindow(int processId, string className, string title)
    {
        IntPtr found = IntPtr.Zero;
        IntPtr excelMain = IntPtr.Zero;
        EnumWindows(delegate(IntPtr hWnd, IntPtr lParam) {
            uint candidatePid;
            GetWindowThreadProcessId(hWnd, out candidatePid);
            if (candidatePid != (uint)processId) return true;
            string candidateClass = WindowClass(hWnd);
            string candidateTitle = WindowText(hWnd);
            if (candidateClass == "XLMAIN") excelMain = hWnd;
            if (candidateClass == className &&
                (String.IsNullOrEmpty(title) || candidateTitle == title)) {
                found = hWnd;
                return false;
            }
            return true;
        }, IntPtr.Zero);
        if (found != IntPtr.Zero || excelMain == IntPtr.Zero) return found;

        EnumChildWindows(excelMain, delegate(IntPtr hWnd, IntPtr lParam) {
            if (WindowClass(hWnd) == className &&
                (String.IsNullOrEmpty(title) || WindowText(hWnd) == title)) {
                found = hWnd;
                return false;
            }
            return true;
        }, IntPtr.Zero);
        return found;
    }

    public static string[] ChildTexts(IntPtr parent)
    {
        List<string> values = new List<string>();
        EnumChildWindows(parent, delegate(IntPtr hWnd, IntPtr lParam) {
            string value = WindowText(hWnd);
            if (!String.IsNullOrWhiteSpace(value) && !values.Contains(value)) {
                values.Add(value);
            }
            return true;
        }, IntPtr.Zero);
        return values.ToArray();
    }
}
"@
        $shell = New-Object -ComObject WScript.Shell
        $stopAt = (Get-Date).AddSeconds($timeoutSeconds)
        $seen = @{}
        $formAbsentChecks = 0
        Write-Output "DRIVER|Started|ProcessId=$processId"
        try {
            while ((Get-Date) -lt $stopAt) {
                Start-Sleep -Milliseconds 200
                if (-not $seen.ContainsKey("ACTION|SeedFormOK")) {
                    if ($shell.AppActivate("invSys Admin - Seed Inventory")) {
                        Start-Sleep -Milliseconds 200
                        $shell.SendKeys("{TAB 4}~")
                        $seen["ACTION|SeedFormOK"] = $true
                        Write-Output "WINDOW|ThunderDFrame|invSys Admin - Seed Inventory"
                        Write-Output "ACTION|SeedFormOK"
                    }
                }
                elseif (-not $seen.ContainsKey("ACTION|ResultDialogOK")) {
                    if ($shell.AppActivate("invSys Admin - Seed Inventory")) {
                        $formAbsentChecks = 0
                    }
                    else {
                        $formAbsentChecks++
                        if ($formAbsentChecks -ge 3 -and $shell.AppActivate("invSys Admin")) {
                            $shell.SendKeys("~")
                            $seen["ACTION|ResultDialogOK"] = $true
                            Write-Output "ACTION|ResultDialogOK"
                        }
                    }
                }

                $seedForm = [InvSysSeedNative]::FindWindow(
                    $processId,
                    "ThunderDFrame",
                    "invSys Admin - Seed Inventory"
                )
                if ($seedForm -ne [IntPtr]::Zero -and
                    -not $seen.ContainsKey("ACTION|SeedFormOK")) {
                    Write-Output "WINDOW|ThunderDFrame|invSys Admin - Seed Inventory"
                    if ($shell.AppActivate("invSys Admin - Seed Inventory")) {
                        Start-Sleep -Milliseconds 200
                        $shell.SendKeys("{TAB 4}~")
                        $seen["ACTION|SeedFormOK"] = $true
                        Write-Output "ACTION|SeedFormOK"
                    }
                }

                $messageBox = [InvSysSeedNative]::FindWindow($processId, "#32770", "")
                if ($messageBox -ne [IntPtr]::Zero) {
                    $messageTitle = [InvSysSeedNative]::WindowText($messageBox)
                    $messageKey = "WINDOW|#32770|" + $messageTitle
                    if (-not $seen.ContainsKey($messageKey)) {
                        $seen[$messageKey] = $true
                        Write-Output $messageKey
                        foreach ($messageText in [InvSysSeedNative]::ChildTexts($messageBox)) {
                            Write-Output ("ELEMENT|" + $messageTitle + "|" + $messageText)
                        }
                    }
                    if ($shell.AppActivate($messageTitle)) {
                        $shell.SendKeys("~")
                    }
                }
            }
        }
        finally {
            if ((Get-Date) -ge $stopAt) {
                Write-Output "TIMEOUT|Public callback did not complete within ${timeoutSeconds}s."
                Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue
            }
            if ($null -ne $shell) {
                try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($shell) } catch {}
            }
        }
    } -ArgumentList $ExcelProcessId, $TimeoutSeconds
}

function Get-FileHashValue {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return "<missing>" }
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
}

function Find-ListObjectInWorkbook {
    param(
        [object]$Workbook,
        [string]$TableName
    )

    if ($null -eq $Workbook) { return $null }
    foreach ($worksheet in $Workbook.Worksheets) {
        foreach ($listObject in $worksheet.ListObjects) {
            if ([string]$listObject.Name -eq $TableName) {
                return $listObject
            }
        }
    }
    return $null
}

function Find-OpenWorkbookByPath {
    param(
        [object]$Excel,
        [string]$Path
    )

    foreach ($workbook in $Excel.Workbooks) {
        try {
            if ([string]::Equals(
                [string]$workbook.FullName,
                $Path,
                [StringComparison]::OrdinalIgnoreCase
            )) {
                return $workbook
            }
        }
        catch {}
    }
    return $null
}

function Get-WorkbookSurfaceCounts {
    param([object]$Workbook)

    $tableCount = 0
    foreach ($worksheet in $Workbook.Worksheets) {
        $tableCount += [int]$worksheet.ListObjects.Count
    }
    return [pscustomobject]@{
        Worksheets = [int]$Workbook.Worksheets.Count
        Tables = $tableCount
    }
}

function Get-ColumnValues {
    param(
        [object]$ListObject,
        [string]$ColumnName
    )

    $values = @()
    if ($null -eq $ListObject -or $null -eq $ListObject.DataBodyRange) {
        return $values
    }
    $columnIndex = Get-ColumnIndexSafe -ListObject $ListObject -ColumnName $ColumnName
    if ($columnIndex -le 0) { return $values }
    for ($rowIndex = 1; $rowIndex -le $ListObject.ListRows.Count; $rowIndex++) {
        $values += [string]$ListObject.DataBodyRange.Cells.Item($rowIndex, $columnIndex).Value2
    }
    return $values
}

function Get-WorkbookTableDataFingerprint {
    param([object]$Workbook)

    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($worksheet in $Workbook.Worksheets) {
        foreach ($listObject in $worksheet.ListObjects) {
            $parts.Add("TABLE=" + [string]$listObject.Name) | Out-Null
            foreach ($column in $listObject.ListColumns) {
                $parts.Add("HEADER=" + [string]$column.Name) | Out-Null
            }
            if ($null -ne $listObject.DataBodyRange) {
                for ($rowIndex = 1; $rowIndex -le $listObject.ListRows.Count; $rowIndex++) {
                    for ($columnIndex = 1; $columnIndex -le $listObject.ListColumns.Count; $columnIndex++) {
                        $parts.Add(
                            "CELL=" + $rowIndex + "," + $columnIndex + "=" +
                            [string]$listObject.DataBodyRange.Cells.Item($rowIndex, $columnIndex).Value2
                        ) | Out-Null
                    }
                }
            }
        }
    }
    $bytes = [Text.Encoding]::UTF8.GetBytes(($parts -join "`n"))
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        return ([BitConverter]::ToString($sha.ComputeHash($bytes))).Replace("-", "").ToLowerInvariant()
    }
    finally {
        $sha.Dispose()
    }
}

function Write-ResultArtifact {
    param(
        [string]$Path,
        [string]$Phase,
        [string]$Status,
        [string]$Detail,
        [string[]]$Windows,
        [hashtable]$Facts,
        [string[]]$SensitiveValues = @()
    )

    function Get-RedactedEvidenceText {
        param([AllowEmptyString()][string]$Value)

        $safeValue = $Value -replace '(?i)ProcessId=\d+', 'ProcessId=<redacted-session>'
        $safeValue = $safeValue -replace '(?i)(?:[A-Z]:\\|\\\\)[^|`"\r\n]*', '<redacted-path>'
        foreach ($sensitiveValue in $SensitiveValues) {
            if (-not [string]::IsNullOrWhiteSpace($sensitiveValue)) {
                $safeValue = $safeValue -replace [regex]::Escape($sensitiveValue), '<redacted-value>'
            }
        }
        return $safeValue
    }

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("# Admin Seed Demo Inventory Packaged Callback $Phase") | Out-Null
    $lines.Add("") | Out-Null
    $lines.Add("- Status: **$Status**") | Out-Null
    $lines.Add("- Callback: ``modAdmin.Seed_DemoInventory``") | Out-Null
    $lines.Add("- Runtime: isolated generated test warehouse") | Out-Null
    foreach ($key in @($Facts.Keys | Sort-Object)) {
        $safeFact = Get-RedactedEvidenceText -Value ([string]$Facts[$key])
        $lines.Add("- ${key}: $safeFact") | Out-Null
    }
    $lines.Add("") | Out-Null
    $lines.Add("## Observed result") | Out-Null
    $lines.Add("") | Out-Null
    $lines.Add((Get-RedactedEvidenceText -Value $Detail)) | Out-Null
    $lines.Add("") | Out-Null
    $lines.Add("## Captured UI") | Out-Null
    $lines.Add("") | Out-Null
    if ($Windows.Count -eq 0) {
        $lines.Add("- No non-Excel windows were captured.") | Out-Null
    }
    else {
        foreach ($window in $Windows) {
            $safeWindow = Get-RedactedEvidenceText -Value $window
            $lines.Add("- ``$safeWindow``") | Out-Null
        }
    }
    [IO.File]::WriteAllLines($Path, $lines)
}

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class InvSysAdminSeedWindow
{
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
"@

$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$deployPath = (Resolve-Path -LiteralPath (Join-Path $repo $DeployRoot)).Path
$helperSource = Join-Path $repo "tools/validate_phase6_live_role_workflows.ps1"
Import-FunctionDefinitions -ScriptPath $helperSource

$requiredHelpers = @(
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
    "Add-ListObjectRow"
)
$missingHelpers = @($requiredHelpers | Where-Object {
    -not (Get-Command $_ -CommandType Function -ErrorAction SilentlyContinue)
})
if ($missingHelpers.Count -gt 0) {
    throw "Missing imported helpers: $($missingHelpers -join ', ')"
}

$resultPath = Join-Path $repo (
    "tests/integration/admin_seed_callback_" + $EvidencePhase.ToLowerInvariant() + "_results.md"
)
$runtimeRoot = Join-Path ([IO.Path]::GetTempPath()) (
    "invsys-admin-seed-callback-" + [guid]::NewGuid().ToString("N")
)
$warehouseId = "WHS" + [guid]::NewGuid().ToString("N").Substring(0, 6).ToUpperInvariant()
$stationId = "S1"
$testUser = if ([string]::IsNullOrWhiteSpace($env:USERNAME)) { "user1" } else { $env:USERNAME }
$testPin = [guid]::NewGuid().ToString("N")
$testPinHash = Get-InvSysCredentialHash -Credential $testPin
$configPath = Join-Path $runtimeRoot ($warehouseId + ".invSys.Config.xlsb")
$authPath = Join-Path $runtimeRoot ($warehouseId + ".invSys.Auth.xlsb")
$inventoryPath = Join-Path $runtimeRoot ($warehouseId + ".invSys.Data.Inventory.xlsb")
$snapshotPath = Join-Path $runtimeRoot ($warehouseId + ".invSys.Snapshot.Inventory.xlsb")
$operatorPath = Join-Path $runtimeRoot ($warehouseId + "." + $stationId + ".Receiving.Operator.xlsb")
$packageNames = @(
    "invSys.Core.xlam",
    "invSys.Inventory.Domain.xlam",
    "invSys.Designs.Domain.xlam",
    "invSys.Operations.xlam",
    "invSys.Admin.xlam"
)

$excel = $null
$opened = New-Object System.Collections.Generic.List[object]
$packages = @{}
$uiJob = $null
$uiEvidence = @()
$currentStep = "startup"
$callbackError = ""
$callbackResult = ""
$timedOut = $false
$passed = $false
$detail = ""
$facts = @{}
$uncRuntimeRoot = ""

try {
    New-Item -ItemType Directory -Path $runtimeRoot -Force | Out-Null
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $true
    $excel.AutomationSecurity = 1

    $currentStep = "create isolated runtime"
    $configWb = New-ConfigWorkbook -Excel $excel -Path $configPath `
        -WarehouseId $warehouseId -StationId $stationId -RuntimeRoot $runtimeRoot
    $authWb = New-AuthWorkbook -Excel $excel -Path $authPath `
        -WarehouseId $warehouseId -StationId $stationId `
        -CurrentUserIds @($testUser) -CredentialHash $testPinHash
    $inventoryWb = New-InventoryWorkbook -Excel $excel -Path $inventoryPath `
        -WarehouseId $warehouseId -SkuRows @()
    $opened.Add($configWb) | Out-Null
    $opened.Add($authWb) | Out-Null
    $opened.Add($inventoryWb) | Out-Null

    $capabilitySheet = Get-WorksheetSafe -Workbook $authWb -WorksheetName "Capabilities"
    $capabilityTable = Get-ListObjectSafe -Worksheet $capabilitySheet -TableName "tblCapabilities"
    Add-ListObjectRow -ListObject $capabilityTable -Values @{
        UserId = $testUser
        Capability = "ADMIN_MAINT"
        WarehouseId = $warehouseId
        StationId = $stationId
        Status = "ACTIVE"
        ValidFrom = ""
        ValidTo = ""
    }
    $authWb.Save()

    $currentStep = "open Core package and isolate its runtime root"
    $packageName = "invSys.Core.xlam"
    $packagePath = Join-Path $deployPath $packageName
    $packageWb = $excel.Workbooks.Open($packagePath)
    $opened.Add($packageWb) | Out-Null
    $packages[$packageName] = $packageWb
    $coreName = [string]$packageWb.Name
    [void](Run-WorkbookMacro -Excel $excel -WorkbookName $coreName `
        -MacroName "modRuntimeWorkbooks.SetCoreDataRootOverride" `
        -Arguments @($runtimeRoot))

    $currentStep = "open remaining packaged add-ins"
    foreach ($packageName in @($packageNames | Where-Object { $_ -ne "invSys.Core.xlam" })) {
        $packagePath = Join-Path $deployPath $packageName
        $packageWb = $excel.Workbooks.Open($packagePath)
        $opened.Add($packageWb) | Out-Null
        $packages[$packageName] = $packageWb
    }

    $operationsName = [string]$packages["invSys.Operations.xlam"].Name
    $adminName = [string]$packages["invSys.Admin.xlam"].Name
    $currentStep = "configure isolated target and sign in"
    $configLoaded = [bool](Run-WorkbookMacro -Excel $excel -WorkbookName $coreName `
        -MacroName "modConfig.LoadConfig" -Arguments @($warehouseId, $stationId))
    $authLoaded = [bool](Run-WorkbookMacro -Excel $excel -WorkbookName $coreName `
        -MacroName "modAuth.LoadAuth" -Arguments @($warehouseId))
    $targetResult = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $coreName `
        -MacroName "modNasConnection.SelectWarehouseTargetForAutomation" `
        -Arguments @($runtimeRoot, $runtimeRoot, $stationId, $true))
    $uncRuntimeRoot = "\\localhost\C$\" + $runtimeRoot.Substring(3)
    $pathsSet = [bool](Run-WorkbookMacro -Excel $excel -WorkbookName $coreName `
        -MacroName "modNasConnection.SetCurrentTargetPathsForTest" `
        -Arguments @($uncRuntimeRoot, $runtimeRoot))
    $signInResult = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $coreName `
        -MacroName "modAuth.SignInCurrentTargetForAutomation" `
        -Arguments @($testUser, $testPin, "ADMIN_MAINT"))

    $facts.ConfigLoaded = $configLoaded
    $facts.AuthLoaded = $authLoaded
    $facts.TargetSelected = $targetResult.StartsWith("OK|")
    $facts.TargetPathsSet = $pathsSet
    $facts.SignedIn = $signInResult.StartsWith("OK|")
    if (-not $configLoaded -or -not $authLoaded -or
        -not $targetResult.StartsWith("OK|") -or -not $pathsSet -or
        -not $signInResult.StartsWith("OK|")) {
        throw "Isolated packaged runtime setup did not reach a signed-in Admin state."
    }

    $configWb.Save()
    $authWb.Save()
    $inventoryWb.Save()
    $authDataBefore = Get-WorkbookTableDataFingerprint -Workbook $authWb
    $inventoryWb.Close($false)
    $authWb.Close($false)
    $configWb.Close($false)
    $configHashBefore = Get-FileHashValue -Path $configPath
    $authHashBefore = Get-FileHashValue -Path $authPath
    $inventoryHashBefore = Get-FileHashValue -Path $inventoryPath
    $configWb = $excel.Workbooks.Open($configPath, 0, $true)
    $opened.Add($configWb) | Out-Null
    $configSurfaceBefore = Get-WorkbookSurfaceCounts -Workbook $configWb
    $configWb.Activate()

    $currentStep = "invoke public Admin seed callback"
    if ($EvidencePhase -eq "GREEN") {
        try {
            [void](Run-WorkbookMacro -Excel $excel `
                -WorkbookName $adminName `
                -MacroName "modAdmin.SetSeedInventorySelectionForAutomation" `
                -Arguments @($warehouseId, $stationId, $testUser))
            [void](Run-WorkbookMacro -Excel $excel -WorkbookName $adminName `
                -MacroName "modAdmin.Seed_DemoInventory")
            $callbackResult = "OK|Returned"
            $uiEvidence = @("ACTION|InjectedFormSelectionThroughSeed_DemoInventory")
        }
        catch {
            $callbackError = $_.Exception.Message
        }
    }
    else {
        $uiJob = Start-SeedUiDriver -ExcelProcessId (Get-ExcelProcessId -Excel $excel)
        try {
            [void](Run-WorkbookMacro -Excel $excel -WorkbookName $adminName `
                -MacroName "modAdmin.Seed_DemoInventory")
        }
        catch {
            $callbackError = $_.Exception.Message
        }
        Start-Sleep -Milliseconds 800
        Stop-Job -Job $uiJob -ErrorAction SilentlyContinue
        Wait-Job -Job $uiJob -Timeout 2 -ErrorAction SilentlyContinue | Out-Null
        $uiEvidence = @(Receive-Job -Job $uiJob -ErrorAction SilentlyContinue |
            ForEach-Object { [string]$_ } | Sort-Object -Unique)
        $uiEvidence += @($uiJob.ChildJobs[0].Error |
            ForEach-Object { "DRIVER_ERROR|" + [string]$_ })
        $uiEvidence = @($uiEvidence | Sort-Object -Unique)
        Remove-Job -Job $uiJob -Force -ErrorAction SilentlyContinue
        $uiJob = $null
        $timedOut = @($uiEvidence | Where-Object { $_ -like "TIMEOUT|*" }).Count -gt 0
    }

    if ($timedOut) {
        $facts.CallbackTimedOut = $true
        $facts.CallbackError = if ([string]::IsNullOrWhiteSpace($callbackError)) { "<none>" } else { $callbackError }
        $detail = "The public callback did not complete within 45 seconds while resolving or presenting its interactive context."
    }
    else {
        $configSurfaceAfter = Get-WorkbookSurfaceCounts -Workbook $configWb
        $configSurfaceChanged = $configSurfaceBefore.Worksheets -ne $configSurfaceAfter.Worksheets -or
            $configSurfaceBefore.Tables -ne $configSurfaceAfter.Tables
        $configWb.Close($false)

        $runtimeInventoryWb = Find-OpenWorkbookByPath -Excel $excel -Path $inventoryPath
        if ($null -ne $runtimeInventoryWb) {
            $runtimeInventoryWb.Save()
            $runtimeInventoryWb.Close($false)
        }
        $configHashAfter = Get-FileHashValue -Path $configPath
        $authHashAfter = Get-FileHashValue -Path $authPath
        $inventoryHashAfter = Get-FileHashValue -Path $inventoryPath
        $authInspectionWb = $excel.Workbooks.Open($authPath, 0, $true)
        $opened.Add($authInspectionWb) | Out-Null
        $authDataAfter = Get-WorkbookTableDataFingerprint -Workbook $authInspectionWb
        $authInspectionWb.Close($false)
        $inspectionWb = $excel.Workbooks.Open($inventoryPath, 0, $true)
        $opened.Add($inspectionWb) | Out-Null
        $entities = Find-ListObjectInWorkbook -Workbook $inspectionWb `
            -TableName "tblInventoryEntities"
        $entityCount = if ($null -eq $entities) { 0 } else { [int]$entities.ListRows.Count }
        $keys = @(Get-ColumnValues -ListObject $entities -ColumnName "System_Key" |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        $conditions = @(Get-ColumnValues -ListObject $entities -ColumnName "Condition")
        $uniqueKeyCount = @($keys | Sort-Object -Unique).Count
        $allGood = $conditions.Count -eq $expectedDemoCount -and @($conditions | Where-Object { $_ -ne "GOOD" }).Count -eq 0
        $catalog = Find-ListObjectInWorkbook -Workbook $inspectionWb `
            -TableName "tblSkuCatalog"
        $catalogCount = if ($null -eq $catalog) { 0 } else { [int]$catalog.ListRows.Count }
        $catalogCategories = @(Get-ColumnValues -ListObject $catalog -ColumnName "CATEGORY")
        $requiredCategories = @("raw", "wip", "shippable", "packaging.ship")
        $catalogCoverage = $catalogCount -eq $expectedDemoCount -and
            @($requiredCategories | Where-Object { $_ -notin $catalogCategories }).Count -eq 0

        $currentStep = "inspect published snapshot"
        $snapshotWb = Find-OpenWorkbookByPath -Excel $excel -Path $snapshotPath
        if ($null -eq $snapshotWb -and (Test-Path -LiteralPath $snapshotPath -PathType Leaf)) {
            $snapshotWb = $excel.Workbooks.Open($snapshotPath, 0, $true)
            $opened.Add($snapshotWb) | Out-Null
        }
        $snapshotTable = Find-ListObjectInWorkbook -Workbook $snapshotWb `
            -TableName "tblInventorySnapshot"
        $snapshotCount = if ($null -eq $snapshotTable) { 0 } else { [int]$snapshotTable.ListRows.Count }
        $snapshotKeys = @(Get-ColumnValues -ListObject $snapshotTable -ColumnName "System_Key" |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        $snapshotConditions = @(Get-ColumnValues -ListObject $snapshotTable -ColumnName "Condition")
        $snapshotCategories = @(Get-ColumnValues -ListObject $snapshotTable -ColumnName "CATEGORY")
        $snapshotUniqueKeyCount = @($snapshotKeys | Sort-Object -Unique).Count
        $snapshotAllGood = $snapshotConditions.Count -eq $expectedDemoCount -and
            @($snapshotConditions | Where-Object { $_ -ne "GOOD" }).Count -eq 0
        $snapshotCategoryCoverage = @($requiredCategories | Where-Object {
            $_ -notin $snapshotCategories
        }).Count -eq 0
        $snapshotMatchesCanonical = $snapshotUniqueKeyCount -eq $uniqueKeyCount -and
            @(Compare-Object -ReferenceObject @($keys | Sort-Object) `
                -DifferenceObject @($snapshotKeys | Sort-Object)).Count -eq 0

        $currentStep = "refresh saved Receiving operator read model"
        $operatorWb = New-OperationalWorkbook -Excel $excel `
            -NameHint "SeedReceivingOps" -Path $operatorPath
        $opened.Add($operatorWb) | Out-Null
        $surfaceEnsured = [bool](Run-WorkbookMacro -Excel $excel -WorkbookName $coreName `
            -MacroName "modOperationsPrimitiveBridge.EnsureReceivingWorkbookSurface" `
            -Arguments @($operatorWb.Name, ""))
        $refreshSucceeded = [bool](Run-WorkbookMacro -Excel $excel -WorkbookName $coreName `
            -MacroName "modOperationsPrimitiveBridge.RefreshInventoryReadModel" `
            -Arguments @($operatorWb.Name, $warehouseId, "LOCAL"))
        $operatorTable = Find-ListObjectInWorkbook -Workbook $operatorWb -TableName "invSys"
        $operatorCount = if ($null -eq $operatorTable) { 0 } else { [int]$operatorTable.ListRows.Count }
        $operatorKeys = @(Get-ColumnValues -ListObject $operatorTable -ColumnName "System_Key" |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        $operatorConditions = @(Get-ColumnValues -ListObject $operatorTable -ColumnName "Condition")
        $operatorCategories = @(Get-ColumnValues -ListObject $operatorTable -ColumnName "CATEGORY")
        $operatorUniqueKeyCount = @($operatorKeys | Sort-Object -Unique).Count
        $operatorAllGood = $operatorConditions.Count -eq $expectedDemoCount -and
            @($operatorConditions | Where-Object { $_ -ne "GOOD" }).Count -eq 0
        $operatorCategoryCoverage = @($requiredCategories | Where-Object {
            $_ -notin $operatorCategories
        }).Count -eq 0
        $operatorMatchesSnapshot = $operatorUniqueKeyCount -eq $snapshotUniqueKeyCount -and
            @(Compare-Object -ReferenceObject @($snapshotKeys | Sort-Object) `
                -DifferenceObject @($operatorKeys | Sort-Object)).Count -eq 0
        $formRefreshReport = [string](Run-WorkbookMacro -Excel $excel `
            -WorkbookName $operationsName `
            -MacroName "modTS_Received.RunReceivingRefreshFormActionForTest" `
            -Arguments @($operatorWb.Name, "DEMO-"))
        $formVisibleRows = if ($formRefreshReport -match '(?:^|\|)VisibleRows=(\d+)(?:\||$)') {
            [int]$Matches[1]
        }
        else {
            0
        }
        $operatorWb.Save()

        $callbackSucceeded = $callbackResult.StartsWith("OK|")
        $successUi = if ($EvidencePhase -eq "GREEN") {
            $callbackSucceeded
        }
        else {
            @($uiEvidence | Where-Object { $_ -eq "ACTION|ResultDialogOK" }).Count -gt 0
        }

        $facts.CallbackTimedOut = $false
        $facts.CallbackError = if ([string]::IsNullOrWhiteSpace($callbackError)) { "<none>" } else { $callbackError }
        $facts.CallbackResult = if ($callbackSucceeded) { "OK|<redacted-detail>" } elseif ($callbackResult -eq "") { "<none>" } else { "FAIL|<redacted-detail>" }
        $facts.EntityCount = $entityCount
        $facts.UniqueSystemKeys = $uniqueKeyCount
        $facts.AllConditionsGood = $allGood
        $facts.CatalogRows = $catalogCount
        $facts.CatalogCategoryCoverage = $catalogCoverage
        $facts.SnapshotFileCreated = Test-Path -LiteralPath $snapshotPath -PathType Leaf
        $facts.SnapshotRows = $snapshotCount
        $facts.SnapshotUniqueSystemKeys = $snapshotUniqueKeyCount
        $facts.SnapshotAllConditionsGood = $snapshotAllGood
        $facts.SnapshotCategoryCoverage = $snapshotCategoryCoverage
        $facts.SnapshotMatchesCanonical = $snapshotMatchesCanonical
        $facts.ReceivingSurfaceEnsured = $surfaceEnsured
        $facts.OperatorRefreshSucceeded = $refreshSucceeded
        $facts.OperatorRowsAfterRefresh = $operatorCount
        $facts.OperatorUniqueSystemKeys = $operatorUniqueKeyCount
        $facts.OperatorAllConditionsGood = $operatorAllGood
        $facts.OperatorCategoryCoverage = $operatorCategoryCoverage
        $facts.OperatorMatchesSnapshot = $operatorMatchesSnapshot
        $facts.ReceivingRefreshFormAction = if ($formRefreshReport.StartsWith("OK|")) { "OK|<redacted-detail>" } else { $formRefreshReport }
        $facts.ReceivingVisibleDemoRows = $formVisibleRows
        $facts.ConfigSurfaceChanged = $configSurfaceChanged
        $facts.ConfigHashUnchanged = $configHashBefore -eq $configHashAfter
        $facts.AuthHashUnchanged = $authHashBefore -eq $authHashAfter
        $facts.AuthTableDataUnchanged = $authDataBefore -eq $authDataAfter
        $facts.InventoryHashChanged = $inventoryHashBefore -ne $inventoryHashAfter

        $passed = [string]::IsNullOrWhiteSpace($callbackError) -and
            $successUi -and -not $configSurfaceChanged -and
            $entityCount -eq $expectedDemoCount -and $uniqueKeyCount -eq $expectedDemoCount -and
            $allGood -and $catalogCoverage -and
            $snapshotCount -eq $expectedDemoCount -and
            $snapshotUniqueKeyCount -eq $expectedDemoCount -and $snapshotAllGood -and
            $snapshotCategoryCoverage -and
            $snapshotMatchesCanonical -and $surfaceEnsured -and $refreshSucceeded -and
            $operatorCount -eq $expectedDemoCount -and $operatorUniqueKeyCount -eq $expectedDemoCount -and
            $operatorAllGood -and $operatorCategoryCoverage -and
            $operatorMatchesSnapshot -and
            $formRefreshReport.StartsWith("OK|") -and $formVisibleRows -eq $expectedDemoCount -and
            $configHashBefore -eq $configHashAfter -and
            $authDataBefore -eq $authDataAfter -and
            $inventoryHashBefore -ne $inventoryHashAfter
        if ($passed) {
            $detail = "The public ribbon callback seeded the complete 19-entity R1 workflow kit, published the same immutable identities to the snapshot, and exposed them through a refreshed saved Receiving operator workbook without using the active canonical config workbook as an Admin surface."
        }
        else {
            $detail = "The public ribbon callback failed its packaged behavioral contract at step '$currentStep'."
        }
    }
}
catch {
    $detail = "Harness exception at step '$currentStep': $($_.Exception.Message)"
    $facts.HarnessException = $_.Exception.Message
}
finally {
    if ($null -ne $uiJob) {
        Stop-Job -Job $uiJob -ErrorAction SilentlyContinue
        Remove-Job -Job $uiJob -Force -ErrorAction SilentlyContinue
    }
    Write-ResultArtifact -Path $resultPath -Phase $EvidencePhase `
        -Status $(if ($passed) { "PASS" } else { "FAIL" }) `
        -Detail $detail -Windows $uiEvidence -Facts $facts `
        -SensitiveValues @($testUser, $testPin, $testPinHash, $runtimeRoot, $uncRuntimeRoot)

    if ($null -ne $excel) {
        for ($index = $opened.Count - 1; $index -ge 0; $index--) {
            try { $opened[$index].Close($false) } catch {}
            Release-ComObject $opened[$index]
        }
        try { $excel.Quit() } catch {}
        Release-ComObject $excel
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if (-not $KeepArtifacts -and
        $runtimeRoot.StartsWith([IO.Path]::GetTempPath(), [StringComparison]::OrdinalIgnoreCase) -and
        [IO.Path]::GetFileName($runtimeRoot).StartsWith("invsys-admin-seed-callback-")) {
        Remove-Item -LiteralPath $runtimeRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Write-Output $detail
Write-Output "Evidence: $resultPath"
if (-not $passed) { exit 1 }
